. "$PSScriptRoot\Helpers.ps1"


function Update-Database
{
	[CmdletBinding()]
	param (
		[String]$ProjectName,
		[String]$ConnectionName,
		[String]$Provider,
		[String]$Target,
		[Switch]$MigrateDown,
		[Switch]$Script = $false
		)

	$completed = Execute-MigrationsCommand "update-database" $ProjectName $ConnectionName $Provider $Target $MigrateDown.IsPresent $Script.IsPresent

	if ($completed)
	{
		Write-Host "Database update is completed."
	}
	else
	{
		Write-Host "Database update is not completed."
	}
}


function Add-Migration
{
	[CmdletBinding(DefaultParameterSetName = 'Name')]
	param (
		[parameter(Position = 0, Mandatory = $true)]
		[String] $Name,
		[String]$ProjectName
		)

	$timestamp = (Get-Date -Format yyyyMMddHHmmss)

	$project = Get-MigrationsProject $ProjectName $false
	$namespace = $project.Properties.Item("DefaultNamespace").Value.ToString() + ".Migrations"
	$projectPath = [System.IO.Path]::GetDirectoryName($project.FullName)
	$migrationsPath = Join-Path $projectPath "Migrations"
	$outputPath = Join-Path $migrationsPath ("$timestamp" + "_$name.cs")

	if (-not (Test-Path $migrationsPath))
	{
		[System.IO.Directory]::CreateDirectory($migrationsPath)
	}

	"using FluentMigrator;

namespace $namespace
{
	[Migration($timestamp)]
	public class $name : Migration
	{
		public override void Up()
		{
		}

		public override void Down()
		{
		}
	}
}" | Out-File -Encoding "UTF8" -Force $outputPath

	$DTE.ExecuteCommand("File.OpenFile", """" + $outputPath + """")
}


Export-ModuleMember @("Update-Database", "Add-Migration")
