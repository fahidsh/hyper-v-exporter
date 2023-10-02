param (
  [string] $VMName = "_NONE_", # Name of the VM to export
  [string] $ExportPath = "default", # Path to export VM to
  [switch] $AppendDate # Whether to append datetime to export path
)

# Get All Virtual Machines on the Hyper-V Host
$VirtualMachines = Get-VM

# Get the Date and Time in the format yyyyMMdd-HHmmss
$DateTime = (Get-Date -Format "yyyyMMdd-HHmmss")

# Default path to export VMs to
$DefaultExportPath = "d:\export"

# A list of VMs to export
$VirtualMachinesToExport = @()

# Integer to keep track of the number of VMs exported
$VMsExported = 0

###################################
### FUNCTIONS
###################################
# check if the given path is valid
function CheckPath {
  param ([Parameter(Mandatory = $true)] [string] $Path)
  if (-not (Test-Path $Path)) {
   Write-Host "Invalid path: $Path" -ForegroundColor Red
   exit
  }
}

# Check if the given VM name is valid
function CheckVMName {
  param ([Parameter(Mandatory = $true)] [string] $VMName)
  if (-not ($VirtualMachines.Name -contains $VMName)) {
   Write-Host "Invalid VM name: $VMName" -ForegroundColor Red
   return $false
  }
  return $true
}

# export the given VM
function ExportVM {
  param (
    [Parameter(Mandatory = $true)] [string] $VMName,
    [Parameter(Mandatory = $true)] [string] $ExportPath
  )
  # Export the VM
  Write-Host "Exporting VM: $VMName to $ExportPath" -ForegroundColor Green
  # Export-VM -Name $VMName -Path $ExportPath -Verbose
}
###################################
### END FUNCTIONS
###################################

# If the VMName parameter is not specified, prompt the user to select a VM to export
if ($VMName -eq "_NONE_") {
  $VirtualMachinesNames = @("All") + $VirtualMachines.Name
  $VMName = $VirtualMachinesNames | Out-GridView -Title "Select VM to export" -OutputMode Multiple
  
  # If no VMs were selected, stop the script
  if ($null -eq $VMName) {
    Write-Host "No VMs to export" -ForegroundColor Red
    exit
  }
  # If "All" was selected, set VMName to "All"
  if ($VMName -contains "All") {
    $VMName = "All"
  }}

# If VMName is "All", export all VMs
if ($VMName -eq "All") {
  $VirtualMachinesToExport = $VirtualMachines.Name
} elseif ($VMName.Contains(",")) {
  # If VMName contains comma, split the VMName string into an array and trim the whitespace
  $VirtualMachinesToExport = $VMName.Split(",") | ForEach-Object { $_.Trim() }
} elseif ($VMName.GetType().Name -eq "Object[]") {
  # If VMName is an array, take it as is
  $VirtualMachinesToExport = $VMName
} elseif (-not ($VMName -eq "_NONE_") -and (-not ($VMName -eq ""))) {
  # If VMName is not "_NONE_" or empty, add it to the array
  $VirtualMachinesToExport = @($VMName)
} else {
  # If VMName is "_NONE_" or empty, stop the script
  Write-Host "No VMs to export" -ForegroundColor Red
  exit
}

# If the ExportPath parameter is not specified, use the default export path
if ($ExportPath.ToLower() -eq "default") {
  $ExportPath = $DefaultExportPath
} else {
  # Check if the given path is valid
  CheckPath -Path $ExportPath
}

# If AppendDate is specified, append date to export path
if ($AppendDate) {
  $ExportPath = "$ExportPath\$DateTime"
}

# Export the VMs
foreach ($VM in $VirtualMachinesToExport) {
  # Check if the VM name is valid
  $IsValidVM = CheckVMName -VMName $VM
  if ($IsValidVM -eq $false) {
    # If the VM name is not valid, skip to the next VM
    continue
  }
  # Export the VM
  ExportVM -VMName $VM -ExportPath $ExportPath
  # Increment the number of VMs exported
  $VMsExported++
}

# If any VMs were exported, display a message
if ($VMsExported -gt 0) {
  Write-Host "$VMsExported Virtual Machine(s) were exported successully" -ForegroundColor Yellow
}
