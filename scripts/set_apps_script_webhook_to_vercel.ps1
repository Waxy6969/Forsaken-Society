$ErrorActionPreference = "Stop"

param(
  [Parameter(Mandatory = $true)]
  [string]$WebhookUrl,

  [string]$Secret = "change-this-secret"
)

npx vercel env add GOOGLE_APPS_SCRIPT_WEBHOOK_URL production --value "$WebhookUrl" --force --yes
npx vercel env add GOOGLE_APPS_SCRIPT_SECRET production --value "$Secret" --force --yes

Write-Host "Apps Script webhook added to Vercel production."
Write-Host "Redeploy with: npx vercel deploy --prod --yes"
