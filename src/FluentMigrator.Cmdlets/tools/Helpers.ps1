function Execute-MigrationsCommand(
	[String]$Command,
	[String]$ProjectName,
	[String]$ConnectionName,
	[String]$Provider,
	[String]$TargetVersion,
	[Boolean]$MigrateDown = $false,
	[Boolean]$Script = $false
	)
{
	$verboseOutput = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent

	$parameters = New-MigrationParameters $ProjectName $ConnectionName
	if (!$parameters)
	{
		return $false
	}

	$migrationAssemblyPath = $parameters.MigrationsAssemblyPath
	$connectionString = $parameters.ConnectionString

	$migratorPath = Init-Migrator $Provider
	
	Write-Verbose "Using migration utility: $migratorPath"
	Write-Verbose "Using migrations from assembly: $migrationAssemblyPath"
	Write-Verbose "Using connectionString: $connectionString"

	if ($Command -eq "update-database")
	{
		if (!$Provider)
		{
			$Provider = "SqlServer2014"
		}
		$parameters = "-db " + $Provider + " -a """ + $migrationAssemblyPath + """ -c """ + $connectionString + """"

		if ($MigrateDown)
		{
			$parameters += " --task migrate:down"

			if (!$TargetVersion)
			{
				$revertAllAnswer = Read-Host -Prompt "You are about to revert ALL migrations at the database. Do you want to proceed [yes/no]?"
				if ($revertAllAnswer.ToLower() -ne "yes")
				{
					return $false;
				}
			}
		}
		if ($TargetVersion)
		{
			$parameters += " --version " + $TargetVersion
		}

		if ($Script)
		{
			$scriptFileName = [System.IO.Path]::GetFileNameWithoutExtension($migrationAssemblyPath) + ".sql"

			$tempDirectory =  $env:temp
			if ($tempDirectory)
			{
				$scriptFileName = (Join-Path $tempDirectory $scriptFileName)
			}
			$parameters += " -o -p --of """ + $scriptFileName + """"
		}

		if ($verboseOutput)
		{
			$parameters += " --verbose true"
		}

		Write-Verbose "Executing: `"$migratorPath`" $parameters"
		$output = Invoke-Expression "& dotnet `"$migratorPath`" $parameters";
		if ($verboseOutput)
		{
			for($i = 0; $i -lt $output.Length; $i++)
			{
				#Skip migrator label
				if ($i -le 6)
				{
					continue;
				}
				Write-Verbose $output[$i]
			}
		}

		if ($scriptFileName)
		{
			$DTE.ExecuteCommand("File.OpenFile", """" + $scriptFileName + """")
		}

		return $true
	}

	Write-Error "Migration command '$Command' is unknown."
	return $false
}

function New-MigrationParameters([String]$ProjectName, [String]$ConnectionName)
{
	$project = Get-MigrationsProject $ProjectName $false
	
	Write-Verbose "Migration project: `"$ProjectName`""
	
	$parameters = Build-Project $project
	if (!$parameters)
	{
		return $false
	}
	
	$migrationsAssemblyPath = $parameters.MigrationsAssemblyPath
	$outputPath = $parameters.OutputPath
	Write-Verbose "Migration AssemblyPath: `"$migrationsAssemblyPath`""
	Write-Verbose "OutputPath: `"$outputPath`""
	
	$configPath = Join-Path $outputPath "appsettings.json"
	$config = Get-Content $configPath | Out-String | ConvertFrom-Json
	
	Write-Verbose "Config: `"$config`""
	
	if ($config)
	{
		$connectionStrings = $config.ConnectionStrings
		if (!$ConnectionName)
		{
			$ConnectionName = "Default"
		}
		$connectionString = $connectionStrings.$ConnectionName
	}

	if (!$connectionString)
	{
		Write-Error "Can not get connection string from the migration project configuration. Name you connection string as 'Default' or specify it's name explicitly using '-ConnectionName' option."
		return $null
	}

	return @{
		MigrationsAssemblyPath = $migrationsAssemblyPath
		ConnectionString = $connectionString
	}
}


function Init-Migrator($provider, $WhatIf = $False)
{
	$packagesPath = Join-Path $PSScriptRoot "../../../../";
		
	$migrateExe = ((Get-ChildItem $packagesPath -Recurse "FluentMigrator.Migrate.dll") | where { $_.FullName -match "FluentMigrator.\d+.\d+.\d+(.\d+)?\\tools" }) | Sort-Object -Property FullName -Descending;
	if ($migrateExe -eq $null)
	{
		throw "Couldn't find migrate.exe anywhere. (Searched {0})" -f $packagesPath;
	}
	$migrateExePath = $migrateExe[0].FullName;

	$architecture = "x86";
	if ($env:PROCESSOR_ARCHITECTURE -match "64")
	{
		$architecture = "AnyCPU";
	}

	$version = "40";
	if ($PSVersionTable.CLRVersion.Major -lt 4)
	{
		$version = "35";
	}

	Write-Verbose "Looking for FluentMigrator.Runner.dll for architecture '$architecture' and CLR version '$version'.";
	# The default assembly works for x86, version 4.0. No need to copy anything.
	if ($architecture -eq "x86" -and $version -eq "40")
	{
		$fluentRunnerPath = ((Get-ChildItem $packagesPath -Recurse "FluentMigrator.Runner.dll") | where { $_.FullName -match "FluentMigrator.\d+.\d+.\d+(.\d+)?\\tools" }).FullName;
	}
	else
	{
		$fluentRunnerPath = ((Get-ChildItem $packagesPath -Recurse "FluentMigrator.Runner.dll") | where { $_.FullName -match "FluentMigrator.Tools.\d+.\d+.\d+(.\d+)?\\tools\\$architecture\\$version" }).FullName;
		$copyTo = Join-Path (Split-Path $migrateExePath) "FluentMigrator.Runner.dll";
		if ($WhatIf)
		{
				Write-Debug ("WHATIF: Would copy '{0}' to '{1}'" -f $fluentRunnerPath, $copyTo);
		}
		elseif ($fluentRunnerPath -eq $null)
		{
				Write-Warning "Couldn't find FluentMigrator.Runner.dll for architecture $architecture and version $version. Please install FluentMigrator.Tools. Will try to continue anyway!";
		}
		else
		{
				Copy-Item $fluentRunnerPath $copyTo "FluentMigrator.Runner.dll" -Force;
		}
	}

	if ($provider -eq "PostGres")
	{
		# Load Npgsql DLL because we will need it later!
		$npgsqlPath = (Join-Path (Split-Path $fluentRunnerPath) "Npgsql.dll");
		if ($WhatIf)
		{
			Write-Debug ("WHATIF: Would load PostGreSQL DLL at '{0}'" -f $npgsqlPath);
		}
		else
		{
			$assembly = [System.Reflection.Assembly]::LoadFrom($npgsqlPath);
		}
	}
	elseif ($provider -eq "SQLite")
	{
		# Load SQLite DLL because we will need it later!
		$sqlitePath = (Join-Path (Split-Path $fluentRunnerPath) "System.Data.SQLite");
		if ($WhatIf)
		{
			Write-Debug ("WHATIF: Would load SQLite DLL at '{0}'" -f $sqlitePath);
		}
		else
		{
			$assembly = [System.Reflection.Assembly]::LoadFrom($sqlitePath);
		}
	}

	return $migrateExePath;
}

function Get-MigrationsProject([String]$name, $hideMessage)
{
	if ($name -and $name.Length -gt 0)
	{
		return Get-SingleProject $name
	}

	$project = Get-Project

	if (!$hideMessage)
	{
		$projectName = $project.Name
		Write-Verbose "Using NuGet project '$projectName'."
	}

	return $project
}

function Get-SingleProject($name)
{
	$project = Get-Project $name
	if ($project -is [array])
	{
		throw "More than one project '$name' was found. Specify the full name of the one to use."
	}
	return $project
}

function Build-Project($project)
{
	$configuration = $DTE.Solution.SolutionBuild.ActiveConfiguration.Name

	$DTE.Solution.SolutionBuild.BuildProject($configuration, $project.UniqueName, $true)
	if ($DTE.Solution.SolutionBuild.LastBuildInfo)
	{
		$projectName = $project.Name
		throw "The project '$projectName' failed to build."
	}

	$outputPath = $project.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value
	$outputFileName = $project.Properties.Item("AssemblyName").Value + ".dll"
	$outputType = $project.Properties.Item("OutputType").Value

	$outputPath = Join-Path (Split-Path $project.FullName) $outputPath
	$outputAssemblyPath = Join-Path $outputPath $outputFileName


	return @{
		MigrationsAssemblyPath = $outputAssemblyPath
		OutputPath = $outputPath
	}
}
