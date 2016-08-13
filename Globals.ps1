#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

#region Get-ScriptDirectory
#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($hostinvocation -ne $null)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}
#endregion

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

#region ExecuteDbDataReaderSqlQuery
	<###############################################################
	#
	# Execute a query and populate the $Datatable_OBJ object with the data
	#
	###############################################################>

function ExecuteDbDataReaderSqlQuery ($DatabaseConnection_STR, $Database_NME, $SQLCommand_STR)
{
	#	Write-Output "Globals.ps1 `$DatabaseConnection_STR = $DatabaseConnection_STR"
	#	Write-Output "Globals.ps1 `$Database_NME           = $Database_NME"
	#	Write-Output "Globals.ps1 `$SQLCommand_STR         = $SQLCommand_STR"
	
	$Datatable_OBJ = New-Object System.Data.DataTable
	
	$Connection = New-Object System.Data.SQLClient.SQLConnection
	$Connection.ConnectionString = "server = '$DatabaseConnection_STR'; database = '$Database_NME'; trusted_connection = true;"
	
	# Open the database connection
	
	$Connection.Open()
	
	$Command = New-Object System.Data.SQLClient.SQLCommand
	$Command.Connection = $Connection
	$Command.CommandText = $SQLCommand_STR
	
	try
	{
		$DbDataReader_OBJ = $Command.ExecuteReader()
	}
	catch
	{
		$ErrorMessageTitle_STR = "Error in function ExecuteDbDataReaderSqlQuery"
		
		$ErrorMessage_STR = "
An error occurred while attempting to access the [$Database_NME] database in the SQL Server $DatabaseConnection_STR instance.

PowerShell Error: $error[0].ToString()"

		if ($Debug_SW -eq 1)
		{
		$ErrorMessage_STR = $ErrorMessage_STR + "
SQL Command:
$SQLCommand_STR"
		}

		[void][System.Windows.Forms.MessageBox]::Show($ErrorMessage_STR, $ErrorMessageTitle_STR)
	}
	
	$Datatable_OBJ.Load($DbDataReader_OBJ)
	
	# Close the database connection
	
	$Connection.Close()
	
	return $Datatable_OBJ
}
#endregion

#region Get-Datatable
	<###############################################################
	#
	# Execute a query and populate the $ds object with the data
	#
	# This code is based on https://www.sapien.com/forums/viewtopic.php?f=21&t=10337&p=55948#p55948
	#
	###############################################################>

function Get-Datatable
{
	Param (
		$DatabaseConnection_STR,
		$Database_NME,
		$SQLCommand_STR
	)
	
	Try
	{
		$connStr = 'Integrated Security=SSPI;Data Source={0};Initial Catalog="{1}";' -f $DatabaseConnection_STR, $Database_NME
		
		$Conn = New-Object System.Data.SQLClient.SQLConnection($connStr)
		
		$Conn.Open()
		
		$cmd = $Conn.CreateCommand()
		$cmd.CommandText = $SQLCommand_STR
		
		$ds = New-Object System.Data.DataSet
		$adapt = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
		$ret = $adapt.Fill($ds)
		
		$ds
		
		$Conn.Close()
	}
	catch
	{
		[void][System.Windows.Forms.MessageBox]::Show($_)
		
		Throw
	}
}
#endregion

#region clean_up_field
	<###############################################################
	#
	# Clean up a data field before writing into the database
	#
	###############################################################>

function clean_up_field ($EditType_SW, $InputField_OBJ)
{
	if ($Debug_SW -eq 1)
	{
		Write-Host ""
		write-host "$(get-date): Begin clean_up_field: $InputField"
		Write-Host ""
		Write-Host "    `$EditType_SW        = $EditType_SW"
		Write-Host "    `$InputField_OBJ     = $InputField_OBJ"
	}
	
	$InputFieldType_STR = $InputField_OBJ.GetType().Name

	$Long_QTY = $InputField_OBJ.length

	if ($Debug_SW -eq 1)
	{
		Write-Host "    `$InputFieldType_STR = $InputFieldType_STR"
		Write-Host "    `$Long_QTY           = $Long_QTY"
	}

	switch ($EditType_SW)
		{
		"DateTime"
			{
				if ($InputFieldType_STR -eq "DateTime")
				{
					# Add single quotes to the date time value
					
					[string]$OutputField_OBJ = "'" + [string]$InputField_OBJ + "'"
				} else {
					[string]$OutputField_OBJ = "null"
				}
			}
		"Text"
			{
				if ($Long_QTY -gt 0)
				{
					# Remove white space at end of variable
		
					$InputField_OBJ = $InputField_OBJ.trim()
	
					# Escape single quotes
					
					[string]$OutputField_OBJ = $InputField_OBJ.Replace('''', '''''')
					
					# Add single quotes to the string
					
					[string]$OutputField_OBJ = "'" + $OutputField_OBJ + "'"
				} else {
					[string]$OutputField_OBJ = "null"
				}
			}
		"Number"
			{
				if ($InputField_OBJ -match "[0-9]")
				{
					# Input is numeric
					[string]$OutputField_OBJ = $InputField_OBJ
				} else {
					# Input is not numeric--return a null
					[string]$OutputField_OBJ = "null"
				}
			}
		default
			{
				# Unsupported edit type--return $false to the caller
			
				return $false
			}
		}

	if ($Debug_SW -eq 1)
	{
		Write-Host ""
		Write-Host "    `$OutputField_OBJ    = $OutputField_OBJ"
		Write-Host ""
		write-host "$(get-date): End clean_up_field"
		Write-Host ""
	}
	
	# Return the cleaned up data to the calling code
	
	$OutputField_OBJ
}
#endregion

#region validation functions
function Validate-FileName
{
	<#
		.SYNOPSIS
			Validates if the file name has valid characters
	
		.DESCRIPTION
			Validates if the file name has valid characters
	
		.PARAMETER  FileName
			A string containing a file name
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param([string]$Filename)
	
	if ($Filename -eq $null -or $Filename -eq "")
    {
		return $false
	}
	 
	$invalidChars = [System.IO.Path]::GetInvalidFileNameChars();

    foreach ($fileChar in $Filename)
    {
        foreach ($invalid in $invalidChars)
        {
            if ($fileChar -eq $invalid)
            {
				return $false
			}
        }
    }
}

function Validate-IsDate
{
	<#
		.SYNOPSIS
			Validates if input is empty (ignores spaces).
	
		.DESCRIPTION
			Validates if input is empty (ignores spaces).
	
		.PARAMETER  Date
			A string containing a date
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#> 
	[OutputType([Boolean])]
	param([string]$Date)
	
	return [DateTime]::TryParse($Date,[ref](New-Object System.DateTime))	
}

function Validate-IsEmail
{
	<#
		.SYNOPSIS
			Validates if input is an Email
	
		.DESCRIPTION
			Validates if input is an Email
	
		.PARAMETER  Email
			A string containing an email address
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param ([string]$Email)
	
	return $Email -match "^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$"
}

function Validate-IsEmpty 
{
	<#
		.SYNOPSIS
			Validates if input is empty.
	
		.DESCRIPTION
			Validates if input is empty.
	
		.PARAMETER  Text
			A string containing an IP address
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param([string]$Text)
	
	return [string]::IsNullOrEmpty($Text)
}

function Validate-IsEmptyTrim 
{
	<#
		.SYNOPSIS
			Validates if input is empty (ignores spaces).
	
		.DESCRIPTION
			Validates if input is empty (ignores spaces).
	
		.PARAMETER  Text
			A string containing an IP address
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param([string]$Text)
	
	if($text -eq $null -or $text.Trim().Length -eq 0)
	{
		return $true	
	}
	
	return $false
}

function Validate-IsIP
{
	<#
		.SYNOPSIS
			Validates if input is an IP Address
	
		.DESCRIPTION
			Validates if input is an IP Address
	
		.PARAMETER  IP
			A string containing an IP address
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>	
	[OutputType([Boolean])]
	param([string] $IP)
	#Regular Express
	#return $IP -match "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
	#Parse using a IPAddress static method
	try
	{
		return ([System.Net.IPAddress]::Parse($IP) -ne $null)
	}
	catch
	{ }
	
	return $false
}

function Validate-IsName
{
	<#
		.SYNOPSIS
			Validates if input is english name
	
		.DESCRIPTION
			Validates if input is english name
	
		.PARAMETER  Name
			A string containing a name 
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param ([string] $Name, [int]$MaxLength)
	
	if($MaxLength -eq $null -or $MaxLength -le 0)
	{#Set default length to 40
		$MaxLength = 40
	}
	
	return $Name -match "^[a-zA-Z''-'\s]{1,$MaxLength}$"
}

function Validate-IsURL
{
	<#
		.SYNOPSIS
			Validates if input is an URL
	
		.DESCRIPTION
			Validates if input is an URL
	
		.PARAMETER  Url
			A string containing an URL address
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param ([string]$Url)
	
	if($Url -eq $null)
	{
		return $false	
	}
	
	return $Url -match "^(ht|f)tp(s?)\:\/\/[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*(\/?)([a-zA-Z0-9\-\.\?\,\'\/\\\+&amp;%\$#_]*)?$"
}

function Validate-IsUSPhone
{
	<#
		.SYNOPSIS
			Validates if input is an phone number
	
		.DESCRIPTION
			Validates if input is an phone number
	
		.PARAMETER  Phone
			A string containing an phone address
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#> 
	[OutputType([Boolean])]
	param([string]$Phone)
	
	return $Phone -match "^[01]?[- .]?(\([2-9]\d{2}\)|[2-9]\d{2})[- .]?\d{3}[- .]?\d{4}$"
}

function Validate-IsZipCode
{
	<#
		.SYNOPSIS
			Validates if input is a zipcode
	
		.DESCRIPTION
			Validates if input is a zipcode
	
		.PARAMETER  ZipCode
			A string containing an zipcode 
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param ([string]$ZipCode)
	
	return $ZipCode -match "^(\d{5}-\d{4}|\d{5}|\d{9})$|^([a-zA-Z]\d[a-zA-Z] \d[a-zA-Z]\d)$"	
}

function Validate-Path
{
	<#
		.SYNOPSIS
			Validates if path has valid characters
	
		.DESCRIPTION
			Validates if path has valid characters
	
		.PARAMETER  Path
			A string containing a directory or file path
	
		.INPUTS
			System.String
	
		.OUTPUTS
			System.Boolean
	#>
	[OutputType([Boolean])]
	param([string]$Path)
	
    if ($Path -eq $null -or $Path -eq "")
    {
		return $false
	}
	
    $invalidChars = [System.IO.Path]::GetInvalidPathChars();

    foreach ($pathChar in $Path)
    {
        foreach ($invalid in $invalidChars)
        {
            if ($pathChar -eq $invalid)
            {
				return $false
			}
        }
    }

    return $true
}
#endregion

	<###############################################################
	#
	# Function to display non-Production environment notification
	#
	###############################################################>

function display_nonproduction_environment
	{
		
		if ($Environment_NME -ne "Production")
		{
			$This.Text = "$Environment_NME " + $This.Text
			$This.BackColor = "$NonProductionFormBackgroundColor_NME"
		}
	}

	<###############################################################
	#
	# Load any additional assemblies
	#
	###############################################################>

	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") # This is required for Messagebox to work

	<###############################################################
	#
	# Call the profile if it is missing (because PowerShell Studio
	# does not call the profile by default)
	#
	###############################################################>

	if ($DBA_SCRIPTS -eq $null)
	{
		. $PSHOME\profile.ps1
	}

	<###############################################################
	#
	# Housekeeping
	#
	###############################################################>

	$OS_OBJ                        = Get-WmiObject -class Win32_OperatingSystem

	[string]$OSVersion_NME         = $OS_OBJ.Caption
	[string]$OSVersion_NBR         = $OS_OBJ.Version

	[string]$PowerShellVersion_STR = "$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
	[float]$PowerShellVersion_NBR  = $PowerShellVersion_STR

	$CurrentUser_NME               = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

	if ($Debug_SW -eq 1)
	{
		Write-Output ""
		Write-Output "`$Debug_SW              = $Debug_SW"
		Write-Output "`$User_SW               = $User_SW"

		Write-Output "`$OSVersion_NME         = $OSVersion_NME"
		Write-Output "`$OSVersion_NBR         = $OSVersion_NBR"
	
		Write-Output "`$PowerShellVersion_STR = $PowerShellVersion_STR"
		Write-Output "`$PowerShellVersion_NBR = $PowerShellVersion_NBR"
	
		Write-Output "`$CurrentUser_NME       = $CurrentUser_NME"
	}
	
	<###############################################################
	#
	# Build the name of the script configuration file
	#
	###############################################################>

	[string]$ScriptDirectory = Get-ScriptDirectory

	$ScriptConfigurationFile_NME = "$ScriptDirectory\settings.ps1"

	if ($Debug_SW -eq 1)
	{
		Write-Output "`$ScriptConfigurationFile_NME = $ScriptConfigurationFile_NME"
	}

	<###############################################################
	#
	# Create the configuration file, if it is missing
	#
	###############################################################>

	Write-Output ""

	if (test-path -path $ScriptConfigurationFile_NME)
	{
		# Settings.ps1 was found
	}
	else
	{
		# Settings.ps1 was not found
	
		Write-Output "$(get-date): Creating the missing `$ScriptConfigurationFile_NME file ($ScriptConfigurationFile_NME)"
	
		$Long_QTY = $CurrentUser_NME.Length
	
		$Admin_STR = "Admin"
		$Admin_STR = $Admin_STR.PadRight($Long_QTY, " ")
	
		$Dot_STR = "."
		$Dot_STR = $Dot_STR.PadRight($Long_QTY, ".")
	
		Write-Output "
	<###############################################################
	#
	# Filename:           settings.ps1
	#
	# Description:        This file is a configuration script
	#
	#                     `$DatabaseConnection_STR--the connection 
	#                     string for the issue tracker database (for
	#                     example, sql_server_network_name\my_instance_name)
	#
	#                     `$Database_NME--the name of the database
	#                     that contains the issue tracker data (for
	#                     example, issue_tracker)
	#
	#                     `$Environment_NME--the name of the 
	#                     processing environment; when this variable 
	#                     is set to something other than `"Production,`" 
	#                     the form title bars will display the 
	#                     contents of this variable and change the 
	#                     background form color to
	#                     `$NonProductionFormBackgroundColor_NME
	#
	#                     `$NonProductionFormBackgroundColor_NME--the
	#                     form background color for non-Production 
	#                     processing environments
	#
	# Windows Version:    $OSVersion_NME ($OSVersion_NBR)
	# PowerShell Version: $PowerShellVersion_STR
	# SQL Server Version: 2012, 2014
	#
	# Edit History:
	#
	# MM/DD/YYYY  $Admin_STR  Description
	# ..........  $Dot_STR  ......................................................................................
	# $(Get-Date -format "MM/dd/yyyy")  $CurrentUser_NME  Created command file
	#
	###############################################################>

	# Please confirm that the following variables are set correctly for this processing environment

	[string]`$DatabaseConnection_STR               = `"$DBA_MetricsDatabaseConnection_STR`" # SQL Server database connection string
	[string]`$Database_NME                         = `"issue_tracker`" # SQL Server database name
	[string]`$Environment_NME                      = `"Production`" # Application environment name (for example, Production, Test, or Development)
	[string]`$NonProductionFormBackgroundColor_NME = `"LightYellow`" # Non-Production form background color (for example, LightYellow, Tan, Pink)
" > $ScriptConfigurationFile_NME # End of Write-Output command
		
		Start-Process notepad.exe $ScriptConfigurationFile_NME -wait
	}

	<###############################################################
	#
	# Execute the configuration file to pick up the script settings
	#
	###############################################################>

	Write-Output ""
	Write-Output "$(get-date): Now `"dot sourcing`" the `$ScriptConfigurationFile_NME file located at $ScriptConfigurationFile_NME"

	. $ScriptConfigurationFile_NME
	
	<###############################################################
	#
	# Determine the assignee_cde of the current user
	#
	###############################################################>
	
	[array]$CurrentAssignee_CDE = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select assignee_cde from assignee where user_nme = '$CurrentUser_NME';"

	[string]$CurrentAssignee_CDE = $CurrentAssignee_CDE.assignee_cde
	
	if ($Debug_SW -eq 1)
	{
		Write-Host ""
		Write-Host "`$DatabaseConnection_STR               = $DatabaseConnection_STR"
		Write-Host "`$Database_NME                         = $Database_NME"
		Write-Host "`$Environment_NME                      = $Environment_NME"
		Write-Host "`$NonProductionFormBackgroundColor_NME = $NonProductionFormBackgroundColor_NME"
		Write-Host "`$CurrentUser_NME                      = $CurrentUser_NME"
		Write-Host "`$CurrentAssignee_CDE                  = $CurrentAssignee_CDE"
	
#		# Potentially use a version number to identify which script is running--see https://www.sapien.com/forums/viewtopic.php?t=10061
#		$ScriptVersion_NBR = [System.Windows.Forms.Application]::ProductVersion
#		Write-Host "`$ScriptVersion_NBR                    = $ScriptVersion_NBR"
	}

	<###############################################################
	#
	# Load arrays with combobox validation data
	#
	# A single set of loads takes place here so that multiple trips
	# to the database can be avoided for the same data
	#
	###############################################################>

	[array]$LocationCode_ARY     = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select location_cde, location_dsc   from location  order by display_order_nbr;"
	[array]$PriorityCode_ARY     = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select priority_cde, priority_dsc   from priority  order by display_order_nbr;"
	[array]$WorkTypeCode_ARY     = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select work_type_cde, work_type_dsc from work_type order by work_type_cde;"
	[array]$CustomerCode_ARY     = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select customer_cde, customer_nme   from customer  order by customer_cde;"

	[array]$StatusCode_ARY       = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select status_cde, status_dsc       from status    order by display_order_nbr;"
#	[array]$StatusCode_ARY       = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select 'All', 'All statuses', 0 as display_order_nbr union all select status_cde, status_dsc, display_order_nbr from status order by display_order_nbr;"
	
	[array]$StatusGroupCode_ARY  = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select distinct status_group_cde    from status    order by status_group_cde;"

	[array]$CategoryCode_ARY     = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select category_cde, category_dsc   from category  order by category_cde;"

	[array]$QueueOrder_ARY       = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select '' as queue_order_nbr, '(Empty)'            union all select convert(varchar, queue_order_nbr), queue_order_dsc from queue_order order by queue_order_nbr;"
	[array]$AssigneeCode_ARY     = ExecuteDbDataReaderSqlQuery -DatabaseConnection_STR $DatabaseConnection_STR -Database_NME $Database_NME -SQLCommand_STR "select '' as assignee_cde,    '(Empty)', '(Empty)' union all select assignee_cde, first_nme, last_nme as full_nme      from assignee    order by assignee_cde;"

#	<###############################################################
#	#
#	# Handle the debug temp.txt file
#	#
#	###############################################################>
#
#	if ($Debug_SW -eq 1)
#	{
#		$DebugFile_NME = "temp.txt"
#		
#		Write-Output "--- Globals.ps1 ------------------------------" >  $DebugFile_NME
#		Write-Output "`$StatusCode_ARY"                               >> $DebugFile_NME
#		$StatusCode_ARY                                               >> $DebugFile_NME
#		Write-Output "--- Globals.ps1 ------------------------------" >> $DebugFile_NME
#		Write-Output "`$StatusGroupCode_ARY"                          >> $DebugFile_NME
#		$StatusGroupCode_ARY                                          >> $DebugFile_NME
#		
#		notepad $DebugFile_NME
#	}
