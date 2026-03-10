param(
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$JsonPath
)

if (!(Test-Path $JsonPath)) {
  throw "JSON file not found: $JsonPath"
}

$json = Get-Content -Raw -Path $JsonPath | ConvertFrom-Json

Write-Host "Connecting to $SiteUrl..."
Write-Host "Importing scheduler tasks and milestones from $JsonPath"

foreach ($task in $json.tasks) {
  Write-Host "Import task: $($task.title)"
  # TODO: Replace with Add-PnPListItem targeting Scheduler Tasks.
}

foreach ($milestone in $json.milestones) {
  Write-Host "Import milestone: $($milestone.title)"
  # TODO: Replace with Add-PnPListItem targeting Scheduler Milestones.
}

Write-Host "JSON import script completed."
