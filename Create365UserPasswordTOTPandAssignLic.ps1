# Requires: Microsoft.Graph PowerShell module
# Install-Module Microsoft.Graph -Scope CurrentUser

# -------------------------------------------------------
# CONFIGURE
# -------------------------------------------------------
$domain        = "Put your Microsoft 365 domain here"
$usageLocation = "US"
$pbiSkuId      = "Put your license SKU that you want to assign to users here"

# -------------------------------------------------------
# HELPER: Generate a random secure password
# -------------------------------------------------------
function New-RandomPassword {
    $length  = 16
    $lower   = 'abcdefghijklmnopqrstuvwxyz'
    $upper   = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits  = '0123456789'
    $special = '!@#$%^&*'
    $all     = $lower + $upper + $digits + $special

    $pwd  = $lower[(Get-Random -Maximum $lower.Length)]
    $pwd += $upper[(Get-Random -Maximum $upper.Length)]
    $pwd += $digits[(Get-Random -Maximum $digits.Length)]
    $pwd += $special[(Get-Random -Maximum $special.Length)]
    $pwd += -join ((1..($length - 4)) | ForEach-Object { $all[(Get-Random -Maximum $all.Length)] })

    return -join ($pwd.ToCharArray() | Get-Random -Count $pwd.Length)
}

# -------------------------------------------------------
# HELPER: Generate a base32-encoded TOTP secret
# -------------------------------------------------------
function New-TOTPSecret {
    $bytes = New-Object byte[] 20
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    $b32    = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'
    $result = ''
    $buffer   = 0
    $bitsLeft = 0
    foreach ($byte in $bytes) {
        $buffer    = ($buffer -shl 8) -bor $byte
        $bitsLeft += 8
        while ($bitsLeft -ge 5) {
            $bitsLeft -= 5
            $result   += $b32[($buffer -shr $bitsLeft) -band 0x1F]
        }
    }
    return $result
}

# -------------------------------------------------------
# HELPER: Compute current TOTP code from a secret
# -------------------------------------------------------
function Get-TOTPCode {
    param([string]$Secret)

    $b32  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'
    $bits = ''
    foreach ($char in $Secret.ToUpper().ToCharArray()) {
        $bits += [Convert]::ToString($b32.IndexOf($char), 2).PadLeft(5, '0')
    }
    $keyBytes = @()
    for ($i = 0; $i -lt ($bits.Length - 7); $i += 8) {
        $keyBytes += [Convert]::ToByte($bits.Substring($i, 8), 2)
    }

    $timeStep  = [long]([System.DateTimeOffset]::UtcNow.ToUnixTimeSeconds() / 30)
    $timeBytes = [BitConverter]::GetBytes($timeStep)
    [Array]::Reverse($timeBytes)

    $hmac     = New-Object System.Security.Cryptography.HMACSHA1
    $hmac.Key = $keyBytes
    $hash     = $hmac.ComputeHash($timeBytes)

    $offset = $hash[$hash.Length - 1] -band 0x0F
    $code   = (($hash[$offset]     -band 0x7F) -shl 24) -bor
              (($hash[$offset + 1] -band 0xFF) -shl 16) -bor
              (($hash[$offset + 2] -band 0xFF) -shl 8)  -bor
               ($hash[$offset + 3] -band 0xFF)

    return ($code % 1000000).ToString("000000")
}

# -------------------------------------------------------
# USER LIST
# -------------------------------------------------------
$users = @(
    @{ Display="Example User";      First="Example";  Last="User";   User="EUser"   },
    @{ Display="Bob Brown";         First="Bob";  Last="Brown";    User="BBrown"    }
)

# -------------------------------------------------------
# MAIN LOOP
# -------------------------------------------------------
$results = @()

foreach ($u in $users) {
    $upn        = "$($u.User)@$domain"
    $password   = New-RandomPassword
    $totpSecret = New-TOTPSecret

    Write-Host "Processing $upn ..." -ForegroundColor Cyan

    try {
        # Step 1: Create user
        Write-Host "  Step 1: Creating user..." -ForegroundColor Gray
        $userBody = @{
            displayName       = $u.Display
            givenName         = $u.First
            surname           = $u.Last
            userPrincipalName = $upn
            mailNickname      = $u.User
            usageLocation     = $usageLocation
            accountEnabled    = $true
            passwordProfile   = @{
                password                      = $password
                forceChangePasswordNextSignIn = $false
            }
        }
        $newUser = New-MgUser -BodyParameter $userBody
        Write-Host "  Step 1 OK - User ID: $($newUser.Id)" -ForegroundColor Gray

        # Step 2: Assign Power BI Pro license
        Write-Host "  Step 2: Assigning license..." -ForegroundColor Gray
        $licenseBody = @{
            addLicenses    = @(@{ skuId = $pbiSkuId })
            removeLicenses = @()
        } | ConvertTo-Json -Depth 3
        Invoke-MgGraphRequest -Method POST `
            -Uri         "https://graph.microsoft.com/v1.0/users/$($newUser.Id)/assignLicense" `
            -Body        $licenseBody `
            -ContentType "application/json" | Out-Null
        Write-Host "  Step 2 OK - License assigned" -ForegroundColor Gray

        # Step 3: Create and assign OATH token
        Write-Host "  Step 3: Creating and assigning TOTP token..." -ForegroundColor Gray
        $serialNumber = "$($u.User)-$(Get-Date -Format 'yyyyMMddHHmmss')"
        $tokenBody = @{
            displayName           = "$($u.Display) MFA Token"
            serialNumber          = $serialNumber
            manufacturer          = "Microsoft"
            model                 = "Software TOTP"
            secretKey             = $totpSecret
            timeIntervalInSeconds = 30
            hashFunction          = "hmacsha1"
            assignTo              = @{ id = $newUser.Id }
        } | ConvertTo-Json -Depth 3
        $tokenResponse = Invoke-MgGraphRequest -Method POST `
            -Uri         "https://graph.microsoft.com/beta/directory/authenticationMethodDevices/hardwareOathDevices" `
            -Body        $tokenBody `
            -ContentType "application/json"
        $deviceId = $tokenResponse.id
        Write-Host "  Step 3 OK - Device ID: $deviceId" -ForegroundColor Gray

        # Step 4: Get the hardwareOathMethod ID on the user
        Write-Host "  Step 4: Retrieving method ID for activation..." -ForegroundColor Gray
        Start-Sleep -Seconds 3
        $userMethods = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/beta/users/$($newUser.Id)/authentication/hardwareOathMethods"
        $methodId = ($userMethods.value | Where-Object { $_.device.id -eq $deviceId }).id
        if (-not $methodId) {
            throw "Could not find hardwareOathMethod ID for device $deviceId on user $upn"
        }
        Write-Host "  Step 4 OK - Method ID: $methodId" -ForegroundColor Gray

        # Step 5: Activate the token
        Write-Host "  Step 5: Activating TOTP token..." -ForegroundColor Gray
        $secondsInWindow = [System.DateTimeOffset]::UtcNow.ToUnixTimeSeconds() % 30
        if ($secondsInWindow -gt 25) {
            Write-Host "  Near window boundary, waiting for next TOTP window..." -ForegroundColor Gray
            Start-Sleep -Seconds (31 - $secondsInWindow)
        }
        $totpCode     = Get-TOTPCode -Secret $totpSecret
        $activateBody = @{
            verificationCode = $totpCode
            displayName      = "$($u.Display) MFA Token"
        } | ConvertTo-Json
        Invoke-MgGraphRequest -Method POST `
            -Uri         "https://graph.microsoft.com/beta/users/$($newUser.Id)/authentication/hardwareOathMethods/$methodId/activate" `
            -Body        $activateBody `
            -ContentType "application/json"
        Write-Host "  Step 5 OK - Token activated" -ForegroundColor Gray

        Write-Host "  Done: $upn" -ForegroundColor Green

        $results += [PSCustomObject]@{
            UPN        = $upn
            Password   = $password
            TOTPSecret = $totpSecret
            Status     = "Success"
        }
    }
    catch {
        Write-Host "  FAILED: $upn" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        $results += [PSCustomObject]@{
            UPN        = $upn
            Password   = $password
            TOTPSecret = $totpSecret
            Status     = "FAILED: $_"
        }
    }
}

# -------------------------------------------------------
# OUTPUT
# -------------------------------------------------------
$csvPath = "$env:USERPROFILE\Desktop\PBI_Users_Credentials.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host ""
Write-Host "Complete. Credentials saved to: $csvPath" -ForegroundColor Yellow
