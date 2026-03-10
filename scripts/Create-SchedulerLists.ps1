param(
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl
)

Write-Host "Connecting to $SiteUrl..."
Write-Host "Creating SharePoint lists for scheduler..."

$lists = @(
  @{ Name = "Scheduler Tasks"; Template = "GenericList" },
  @{ Name = "Scheduler Milestones"; Template = "GenericList" }
)

foreach ($list in $lists) {
  Write-Host "Ensure list exists: $($list.Name)"
  # TODO: Replace with Add-PnPList / Set-PnPField commands in tenant environment.
}

Write-Host "List provisioning script completed."
