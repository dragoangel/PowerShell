$mailcow = 'mail.moo.com'
$api_key = '000000-000000-000000-000000-000000'
$method = "https://${mailcow}/api/v1/add/tls-policy-map"
$headers = @{
    'X-API-Key' = "${api_key}"
    'Content-Type' = "application/json"
}

$policies = Get-Content $PSScriptRoot\policies.json | Out-String | ConvertFrom-Json

foreach ($policy in $policies) {
    $params = @{
        "dest" = $policy.dest
        "policy" = $policy.policy
        "parameters" = $policy.parameters
        "active" = "1"
    }
    Write-Host "Creating policy for:" $params.dest
    Invoke-RestMethod -Uri $method -Method POST -Headers $headers -Body $params
}
