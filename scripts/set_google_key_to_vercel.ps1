$ErrorActionPreference = "Stop"

param(
  [Parameter(Mandatory = $true)]
  [string]$KeyPath
)

if (-not (Test-Path -LiteralPath $KeyPath)) {
  throw "Key file not found: $KeyPath"
}

$raw = Get-Content -LiteralPath $KeyPath -Raw
$json = $raw | ConvertFrom-Json

if (-not $json.client_email -or -not $json.private_key) {
  throw "This does not look like a Google service account JSON key."
}

Write-Host "Service account email:" $json.client_email
Write-Host "Share the Google Sheet with this email as Editor before testing submissions."

$compact = $json | ConvertTo-Json -Depth 20 -Compress
$compact | npx vercel env add GOOGLE_SERVICE_ACCOUNT_JSON production --force --yes

Write-Host "GOOGLE_SERVICE_ACCOUNT_JSON added to Vercel production."
Write-Host "Redeploy with: npx vercel deploy --prod --yes"
