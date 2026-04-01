#Requires -Version 5.1

<#
.SYNOPSIS
	Adds a "Run as administrator" context menu entry for .ps1 files in Windows Explorer.

.DESCRIPTION
	Creates or updates a shell\runas verb for the current .ps1 file association, enabling
	right-click "Run as administrator" for PowerShell scripts directly from Windows Explorer.

	By default, the entry is created per user under HKEY_CURRENT_USER\Software\Classes,
	so administrator rights are not required.

	When -AllUsers is specified, the entry is created under
	HKEY_LOCAL_MACHINE\Software\Classes for all users, which requires an elevated
	PowerShell session.

	Use -Remove to delete a previously created context menu entry.

	The script supports -WhatIf and -Confirm via ShouldProcess.

.PARAMETER AllUsers
	Creates or removes the context menu entry for all users under
	HKEY_LOCAL_MACHINE\Software\Classes. Requires an elevated PowerShell session.

.PARAMETER UsePwsh
	Uses PowerShell 7 (pwsh.exe) instead of Windows PowerShell (powershell.exe).

.PARAMETER BypassExecutionPolicy
	Adds -ExecutionPolicy Bypass to the launched command line.
	Use this only when you explicitly want that behavior.

.PARAMETER Remove
	Removes a previously created context menu entry instead of creating one.

.EXAMPLE
	.\Add-RunAsAdministratorContextMenu.ps1

	Creates the context menu entry for the current user using Windows PowerShell.

.EXAMPLE
	.\Add-RunAsAdministratorContextMenu.ps1 -UsePwsh

	Creates the context menu entry for the current user using PowerShell 7.

.EXAMPLE
	.\Add-RunAsAdministratorContextMenu.ps1 -AllUsers

	Creates the context menu entry for all users (requires elevation).

.EXAMPLE
	.\Add-RunAsAdministratorContextMenu.ps1 -Remove

	Removes the context menu entry for the current user.

.EXAMPLE
	.\Add-RunAsAdministratorContextMenu.ps1 -WhatIf

	Shows what would be changed without writing to the registry.

.OUTPUTS
	System.Management.Automation.PSCustomObject
	Returns an object with details about the created, updated, removed, or missing entry.

.NOTES
	Author : Zupfel
	License : MIT
	Platform: Windows only

.LINK
	https://github.com/Zupfel/PowerShellScripts/
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Add')]
[OutputType([PSCustomObject])]
param(
	[switch]$AllUsers,
	[Parameter(ParameterSetName = 'Add')]
	[switch]$UsePwsh,
	[Parameter(ParameterSetName = 'Add')]
	[switch]$BypassExecutionPolicy,
	[Parameter(ParameterSetName = 'Remove')]
	[switch]$Remove
)
 
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
 
# ---------------------------------------------------------------------------
# Guard: Windows only
# ---------------------------------------------------------------------------
if (-not ($env:OS -eq 'Windows_NT' -or $IsWindows)) {
	throw 'This script is supported on Windows only.'
}
 
# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
 
function Test-IsAdministrator {
	$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
	$principal = [Security.Principal.WindowsPrincipal]::new($identity)

	return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
 
function Get-Ps1ProgId {
	$candidates = @(
		'Registry::HKEY_CURRENT_USER\Software\Classes\.ps1',
		'Registry::HKEY_CLASSES_ROOT\.ps1'
	)

	foreach ($path in $candidates) {
		if (Test-Path -LiteralPath $path) {
			$value = (Get-Item -LiteralPath $path).GetValue('')

			if (-not [string]::IsNullOrWhiteSpace($value)) {
				return $value
			}
		}
	}
 
	return 'Microsoft.PowerShellScript.1'
}
 
function Get-ClassesRootPath {
	param([switch]$AllUsers)
 
	if ($AllUsers) {
		if (-not (Test-IsAdministrator)) {
			throw 'The -AllUsers parameter requires an elevated (administrator) PowerShell session.'
		}

		return 'Registry::HKEY_LOCAL_MACHINE\Software\Classes'
	}
 
	return 'Registry::HKEY_CURRENT_USER\Software\Classes'
}
 
function Resolve-PowerShellHostPath {
	param([switch]$UsePwsh)
 
	if (-not $UsePwsh) {
		$exePath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'
 
		if (-not (Test-Path -LiteralPath $exePath)) {
			throw "Windows PowerShell executable not found: $exePath"
		}
 
		return $exePath
	}
 
	# Try PATH first
	$cmd = Get-Command -Name 'pwsh.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
 
	if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Path)) {
		return $cmd.Path
	}
 
	# Fall back to well-known install locations
	$candidates = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Join-Path -Path $_ -ChildPath 'PowerShell\7\pwsh.exe' } | Select-Object -Unique
 
	foreach ($candidate in $candidates) {
		if (Test-Path -LiteralPath $candidate) {
			return $candidate
		}
	}
 
	throw 'PowerShell 7 (pwsh.exe) was not found. Install PowerShell 7 or omit -UsePwsh.'
}
 
# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
 
try {
	$progId      = Get-Ps1ProgId
	$classesRoot = Get-ClassesRootPath -AllUsers:$AllUsers
	$runAsKey    = Join-Path -Path $classesRoot -ChildPath "$progId\shell\runas"
	$scope       = if ($AllUsers) { 'AllUsers' } else { 'CurrentUser' }
 
	# --- Remove mode ---
	if ($Remove) {
		if (-not (Test-Path -LiteralPath $runAsKey)) {
			[PSCustomObject]@{
				Action      = 'NotFound'
				Scope       = $scope
				ProgId      = $progId
				RegistryKey = $runAsKey
			}
 
			return
		}
 
		if ($PSCmdlet.ShouldProcess($runAsKey, 'Remove the .ps1 Run as administrator context menu entry')) {
			Remove-Item -LiteralPath $runAsKey -Recurse -Force
 
			[PSCustomObject]@{
				Action      = 'Removed'
				Scope       = $scope
				ProgId      = $progId
				RegistryKey = $runAsKey
			}
		}
 
		return
	}
 
	# --- Create / update mode ---
	$hostExe = Resolve-PowerShellHostPath -UsePwsh:$UsePwsh
 
	$menuText = if ($UsePwsh) {
		'Run as administrator (PowerShell 7)'
	}
	else {
		'Run as administrator (PowerShell)'
	}
 
	$commandKey  = Join-Path -Path $runAsKey -ChildPath 'command'
	$entryExists = Test-Path -LiteralPath $runAsKey
 
	$commandLine = "`"$hostExe`" -NoProfile"

	if ($BypassExecutionPolicy) {
		$commandLine += ' -ExecutionPolicy Bypass'
	}

	$commandLine += ' -File "%1"'
 
	if ($PSCmdlet.ShouldProcess($runAsKey, 'Create or update the .ps1 Run as administrator context menu entry')) {
		# Ensure registry keys exist
		if (-not $entryExists) {
			New-Item -Path $runAsKey | Out-Null
		}
		if (-not (Test-Path -LiteralPath $commandKey)) {
			New-Item -Path $commandKey | Out-Null
		}
 
		# Set default values (display text and command line)
		Set-ItemProperty -LiteralPath $runAsKey   -Name '(default)' -Value $menuText    -Type String
		Set-ItemProperty -LiteralPath $commandKey -Name '(default)' -Value $commandLine -Type String
 
		# UAC shield icon and visual indicator
		Set-ItemProperty -LiteralPath $runAsKey -Name 'HasLUAShield' -Value '' -Type String
		Set-ItemProperty -LiteralPath $runAsKey -Name 'Icon'         -Value $hostExe -Type String
 
		[PSCustomObject]@{
			Action         = if ($entryExists) { 'Updated' } else { 'Created' }
			Scope          = $scope
			ProgId         = $progId
			MenuText       = $menuText
			HostExecutable = $hostExe
			RegistryKey    = $runAsKey
			Command        = $commandLine
		}
	}
}
catch {
	$PSCmdlet.ThrowTerminatingError($_)
}
