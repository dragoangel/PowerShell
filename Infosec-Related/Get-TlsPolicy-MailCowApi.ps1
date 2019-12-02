$mailcow = 'mail.moo.com'
$api_key = '000000-000000-000000-000000-000000'
$method = "https://${mailcow}/api/v1/get/tls-policy-map/all"
$headers = @{
    'X-API-Key' = "${api_key}"
    'Content-Type' = "application/json"
}

$policies = Invoke-RestMethod -Uri $method -Method GET -Headers $headers -Body $params
foreach ($policy in $policies) {
    $policy
}
$policies | ConvertTo-Json | Out-File -FilePath $PSScriptRoot\current-policies.json
