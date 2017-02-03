# ***************************************************************************
# 
# File:      MDTDB.psm1
# 
# Author:    Michael Niehaus 
# 
# Purpose:   Provides a set of PowerShell advanced functions (cmdlets) to
#            manipulate the Microsoft Deployment Toolkit database contents.
#            This required at least PowerShell 2.0 CTP3.
#
# Usage:     This script must be imported using "import-module", e.g.:
#              import-module C:\MDTDB.psm1
#            After it has been imported, the indivual functions below can be
#            used.  For details on the parameters each one takes, you can
#            use "get-help", e.g. "get-help Connect-MDTDatabase".  Note that
#            there is no detailed help provided on the cmdlets.  Feel free to
#            add your own...
#
# ------------- DISCLAIMER -------------------------------------------------
# This script code is provided as is with no guarantee or waranty concerning
# the usability or impact on systems and may be used, distributed, and
# modified in any way provided the parties agree and acknowledge the 
# Microsoft or Microsoft Partners have neither accountabilty or 
# responsibility for results produced by use of this script.
#
# Microsoft will not provide any support through any means.
# ------------- DISCLAIMER -------------------------------------------------
#
# ***************************************************************************

# ---------------------------------------------------------------------
# Helper functions (not intended to be called directly)
# ---------------------------------------------------------------------

function Clear-MDTArray {

    PARAM
    (
        $id,
        $type,
        $table
    )

    # Build the delete command
    $delCommand = "DELETE FROM $table WHERE ID = $id and Type = '$type'"
        
    # Issue the delete command
    Write-Verbose "About to issue command: $delCommand"
    $cmd = New-Object System.Data.SqlClient.SqlCommand($delCommand, $mdtSQLConnection)
    $null = $cmd.ExecuteScalar()

    Write-Verbose "Removed all records from $table for Type = $type and ID = $id."
}

function Get-MDTArray {

    PARAM
    (
        $id,
        $type,
        $table,
        $column
    )

    # Build the select command
    $sql = "SELECT $column FROM $table WHERE ID = $id AND Type = '$type' ORDER BY Sequence"
        
    # Issue the select command and return the results
    Write-Verbose "About to issue command: $sql"
    $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
    $selectDataset = New-Object System.Data.Dataset
    $null = $selectAdapter.Fill($selectDataset, "$table")
    $selectDataset.Tables[0].Rows 
}

function Set-MDTArray {

    PARAM
    (
        $id,
        $type,
        $table,
        $column,
        $array
    )

    # First clear the existing array
    Clear-MDTArray $id $type $table
    
    # Now insert each row in the array
    $seq = 1
    foreach ($item in $array)
    {
        # Insert the  row
        $sql = "INSERT INTO $table (Type, ID, Sequence, $column) VALUES ('$type', $id, $seq, '$item')"
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()

        # Increment the counter
        $seq = $seq + 1
    }
        
    Write-Verbose "Added records to $table for Type = $type and ID = $id."
}

# ---------------------------------------------------------------------
# Connection function
# ---------------------------------------------------------------------

function Connect-MDTDatabase {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1)] $drivePath = "",
        [Parameter()] $sqlServer,
        [Parameter()] $instance = "",
        [Parameter()] $database
    )

    # Clear the results from any previous execution
    Clear-Variable -name mdtDatabase -errorAction SilentlyContinue

    # If a drive path is specified, use PowerShell to build the connection string.
    # Otherwise, build it from the other parameters
    if ($drivePath -ne "")
    {
        # Get the needed properties to build the connection string    
        $mdtProperties = get-itemproperty $drivePath

        $mdtSQLConnectString = "Server=$($mdtProperties.'Database.SQLServer')"
        if ($mdtProperties."Database.Instance" -ne "")
        {
            $mdtSQLConnectString = "$mdtSQLConnectString\$($mdtProperties.'Database.Instance')"
        }
        $mdtSQLConnectString = "$mdtSQLConnectString; Database='$($mdtProperties.'Database.Name')'; Integrated Security=true;"
    }
    else
    {
        $mdtSQLConnectString = "Server=$($sqlServer)"
        if ($instance -ne "")
        {
            $mdtSQLConnectString = "$mdtSQLConnectString\$instance"
        }
        $mdtSQLConnectString = "$mdtSQLConnectString; Database='$database'; Integrated Security=true;"
    }
    
    # Make the connection and save it in a global variable
    Write-Verbose "Connecting to: $mdtSQLConnectString"
    $global:mdtSQLConnection = new-object System.Data.SqlClient.SqlConnection
    $global:mdtSQLConnection.ConnectionString = $mdtSQLConnectString
    $global:mdtSQLConnection.Open()
}

# ---------------------------------------------------------------------
# Computer functions
# ---------------------------------------------------------------------

function New-MDTComputer {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $assetTag,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $macAddress,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $serialNumber,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $uuid,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $description,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $settings
    )

    Process
    {
        # Insert a new computer row and get the identity result
        $sql = "INSERT INTO ComputerIdentity (AssetTag, SerialNumber, MacAddress, UUID, Description) VALUES ('$assetTag', '$serialNumber', '$macAddress', '$uuid', '$description') SELECT @@IDENTITY"
        Write-Verbose "About to execute command: $sql"
        $identityCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $identity = $identityCmd.ExecuteScalar()
        Write-Verbose "Added computer identity record"
    
        # Insert the settings row, adding the values as specified in the hash table
        $settingsColumns = $settings.Keys -join ","
        $settingsValues = $settings.Values -join "','"
        $sql = "INSERT INTO Settings (Type, ID, $settingsColumns) VALUES ('C', $identity, '$settingsValues')"
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified computer"
        
        # Write the new record back to the pipeline
        Get-MDTComputer -ID $identity
    }
}

function Get-MDTComputer {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $id = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $assetTag = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $macAddress = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $serialNumber = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $uuid = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $description = ""
    )
    
    Process
    {
        # Build a select statement based on what parameters were specified
        if ($id -eq "" -and $assetTag -eq "" -and $macAddress -eq "" -and $serialNumber -eq "" -and $uuid -eq "" -and $description -eq "")
        {
            $sql = "SELECT * FROM ComputerSettings"
        }
        elseif ($id -ne "")
        {
            $sql = "SELECT * FROM ComputerSettings WHERE ID = $id"
        }
        else
        {
            # Specified the initial command
            $sql = "SELECT * FROM ComputerSettings WHERE "
        
            # Add the appropriate where clauses
            if ($assetTag -ne "")
            {
                $sql = "$sql AssetTag='$assetTag' AND"
            }
        
            if ($macAddress -ne "")
            {
                $sql = "$sql MacAddress='$macAddress' AND"
            }

            if ($serialNumber -ne "")
            {
                $sql = "$sql SerialNumber='$serialNumber' AND"
            }

            if ($uuid -ne "")
            {
                $sql = "$sql UUID='$uuid' AND"
            }

            if ($description -ne "")
            {
                $sql = "$sql Description='$description' AND"
            }
    
            # Chop off the last " AND"
            $sql = $sql.Substring(0, $sql.Length - 4)
        }
    
        $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
        $selectDataset = New-Object System.Data.Dataset
        $null = $selectAdapter.Fill($selectDataset, "ComputerSettings")
        $selectDataset.Tables[0].Rows
    }
}

function Set-MDTComputer {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(Mandatory=$true)] $settings
    )
    
    Process
    {
        # Add each each hash table entry to the update statement
        $sql = "UPDATE Settings SET"
        foreach ($setting in $settings.GetEnumerator())
        {
            $sql = "$sql $($setting.Key) = '$($setting.Value)', "
        }
        
        # Chop off the trailing ", "
        $sql = $sql.Substring(0, $sql.Length - 2)

        # Add the where clause
        $sql = "$sql WHERE ID = $id AND Type = 'C'"
        
        # Execute the command
        Write-Verbose "About to execute command: $sql"        
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified computer"
        
        # Write the updated record back to the pipeline
        Get-MDTComputer -ID $id
    }
}


function Set-MDTComputerIdentity {
[CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(Mandatory=$true)] [Hashtable]$settings
    )
    
    Process
    {
        # Add each each hash table entry to the update statement
        $sql = "UPDATE ComputerIdentity SET"
        foreach ($setting in $settings.GetEnumerator())
        {
            $sql = "$sql $($setting.Key) = '$($setting.Value)', "
        }
        
        # Chop off the trailing ", "
        $sql = $sql.Substring(0, $sql.Length - 2)

        # Add the where clause
        $sql = "$sql WHERE ID = $id"
        
        # Execute the command
        Write-Verbose "About to execute command: $sql"        
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Update settings for the specified computer"
        
        # Write the updated record back to the pipeline
        Get-MDTComputer -ID $id
    }
}

function Remove-MDTComputer {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        # Build the delete command
        $delCommand = "DELETE FROM ComputerIdentity WHERE ID = $id"
        
        # Issue the delete command
        Write-Verbose "About to issue command: $delCommand"
        $cmd = New-Object System.Data.SqlClient.SqlCommand($delCommand, $mdtSQLConnection)
        $null = $cmd.ExecuteScalar()

        Write-Verbose "Removed the computer with ID = $id."
    }
}

function Get-MDTComputerApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'C' 'Settings_Applications' 'Applications'
    }
}

function Clear-MDTComputerApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'C' 'Settings_Applications'
    }
}

function Set-MDTComputerApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $applications
    )

    Process
    {
        Set-MDTArray $id 'C' 'Settings_Applications' 'Applications' $applications
    }
}

function Get-MDTComputerPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'C' 'Settings_Packages' 'Packages'
    }
}

function Clear-MDTComputerPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'C' 'Settings_Packages'
    }
}

function Set-MDTComputerPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $packages
    )

    Process
    {
        Set-MDTArray $id 'C' 'Settings_Packages' 'Packages' $packages
    }
}

function Get-MDTComputerRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'C' 'Settings_Roles' 'Role'
    }
}

function Clear-MDTComputerRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'C' 'Settings_Roles'
    }
}

function Set-MDTComputerRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $roles
    )

    Process
    {
        Set-MDTArray $id 'C' 'Settings_Roles' 'Role' $roles
    }
}

function Get-MDTComputerAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'C' 'Settings_Administrators' 'Administrators'
    }
}

function Clear-MDTComputerAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'C' 'Settings_Administrators'
    }
}

function Set-MDTComputerAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $administrators
    )

    Process
    {
        Set-MDTArray $id 'C' 'Settings_Administrators' 'Administrators' $administrators
    }
}

# ---------------------------------------------------------------------
# Role functions
# ---------------------------------------------------------------------

function New-MDTRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $name,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $settings
    )

    Process
    {
        # Insert a new role row and get the identity result
        $sql = "INSERT INTO RoleIdentity (Role) VALUES ('$name') SELECT @@IDENTITY"
        Write-Verbose "About to execute command: $sql"
        $identityCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $identity = $identityCmd.ExecuteScalar()
        Write-Verbose "Added role identity record"
    
        # Insert the settings row, adding the values as specified in the hash table
        $settingsColumns = $settings.Keys -join ","
        $settingsValues = $settings.Values -join "','"
        $sql = "INSERT INTO Settings (Type, ID, $settingsColumns) VALUES ('R', $identity, '$settingsValues')"
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified role"
        
        # Write the new record back to the pipeline
        Get-MDTRole -ID $identity
    }
}

function Get-MDTRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $id = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $name = ""
    )
    
    Process
    {
        # Build a select statement based on what parameters were specified
        if ($id -eq "" -and $name -eq "")
        {
            $sql = "SELECT * FROM RoleSettings"
        }
        elseif ($id -ne "")
        {
            $sql = "SELECT * FROM RoleSettings WHERE ID = $id"
        }
        else
        {
            $sql = "SELECT * FROM RoleSettings WHERE Role = '$name'"
        }
        
        # Execute the statement and return the results    
        $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
        $selectDataset = New-Object System.Data.Dataset
        $null = $selectAdapter.Fill($selectDataset, "RoleSettings")
        $selectDataset.Tables[0].Rows
    }
}

function Set-MDTRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(Mandatory=$true)] $settings
    )
    
    Process
    {
        # Add each each hash table entry to the update statement
        $sql = "UPDATE Settings SET"
        foreach ($setting in $settings.GetEnumerator())
        {
            $sql = "$sql $($setting.Key) = '$($setting.Value)', "
        }
        
        # Chop off the trailing ", "
        $sql = $sql.Substring(0, $sql.Length - 2)

        # Add the where clause
        $sql = "$sql WHERE ID = $id AND Type = 'R'"
        
        # Execute the command
        Write-Verbose "About to execute command: $sql"        
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified role"
        
        # Write the updated record back to the pipeline
        Get-MDTRole -ID $id
    }
}

function Remove-MDTRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        # Build the delete command
        $delCommand = "DELETE FROM RoleIdentity WHERE ID = $id"
        
        # Issue the delete command
        Write-Verbose "About to issue command: $delCommand"
        $cmd = New-Object System.Data.SqlClient.SqlCommand($delCommand, $mdtSQLConnection)
        $null = $cmd.ExecuteScalar()

        Write-Host "Removed the role with ID = $id."
    }
}

function Get-MDTRoleApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'R' 'Settings_Applications' 'Applications'
    }
}

function Clear-MDTRoleApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'R' 'Settings_Applications'
    }
}

function Set-MDTRoleApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $applications
    )

    Process
    {
        Set-MDTArray $id 'R' 'Settings_Applications' 'Applications' $applications
    }
}

function Get-MDTRolePackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'R' 'Settings_Packages' 'Packages'
    }
}

function Clear-MDTRolePackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'R' 'Settings_Packages'
    }
}

function Set-MDTRolePackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $packages
    )

    Process
    {
        Set-MDTArray $id 'R' 'Settings_Packages' 'Packages' $packages
    }
}

function Get-MDTRoleRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'R' 'Settings_Roles' 'Role'
    }
}

function Clear-MDTRoleRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'R' 'Settings_Roles'
    }
}

function Set-MDTRoleRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $roles
    )

    Process
    {
        Set-MDTArray $id 'R' 'Settings_Roles' 'Role' $roles
    }
}

function Get-MDTRoleAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'R' 'Settings_Administrators' 'Administrators'
    }
}

function Clear-MDTRoleAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'R' 'Settings_Administrators'
    }
}

function Set-MDTRoleAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $administrators
    )

    Process
    {
        Set-MDTArray $id 'R' 'Settings_Administrators' 'Administrators' $administrators
    }
}

# ---------------------------------------------------------------------
# Location functions
# ---------------------------------------------------------------------

function New-MDTLocation {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $name,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $gateways,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $settings
    )

    Process
    {
        # Insert a new role row and get the identity result
        $sql = "INSERT INTO LocationIdentity (Location) VALUES ('$name') SELECT @@IDENTITY"
        Write-Verbose "About to execute command: $sql"
        $identityCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $identity = $identityCmd.ExecuteScalar()
        Write-Verbose "Added location identity record"
    
        # Set the gateways
        $null = Set-MDTLocation -id $identity -gateways $gateways
        
        # Insert the settings row, adding the values as specified in the hash table
        $settingsColumns = $settings.Keys -join ","
        $settingsValues = $settings.Values -join "','"
        $sql = "INSERT INTO Settings (Type, ID, $settingsColumns) VALUES ('L', $identity, '$settingsValues')"
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified location"
        
        # Write the new record back to the pipeline
        Get-MDTLocation -ID $identity
    }
}

function Get-MDTLocation {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $id = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $name = "",
        [Parameter()][switch] $detail = $false
    )
    
    Process
    {
        # Build a select statement based on what parameters were specified
        if ($id -eq "" -and $name -eq "")
        {
            if ($detail)
            {
                $sql = "SELECT * FROM LocationSettings"
            }
            else
            {
                $sql = "SELECT DISTINCT ID, Location FROM LocationSettings"
            }
        }
        elseif ($id -ne "")
        {
            if ($detail)
            {
                $sql = "SELECT * FROM LocationSettings WHERE ID = $id"
            }
            else
            {
                $sql = "SELECT DISTINCT ID, Location FROM LocationSettings WHERE ID = $id"
            }
        }
        else
        {
            if ($detail)
            {
                $sql = "SELECT * FROM LocationSettings WHERE Location = '$name'"
            }
            else
            {
                $sql = "SELECT DISTINCT ID, Location FROM LocationSettings WHERE Location = '$name'"
            }
        }
        
        # Execute the statement and return the results    
        $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
        $selectDataset = New-Object System.Data.Dataset
        $null = $selectAdapter.Fill($selectDataset, "LocationSettings")
        $selectDataset.Tables[0].Rows
    }
}

function Set-MDTLocation {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $gateways = $null,
        [Parameter()] $settings = $null
    )
    
    Process
    {
        # If there are some new settings save them
        if ($settings -ne $null)
        {
            # Add each each hash table entry to the update statement
            $sql = "UPDATE Settings SET"
            foreach ($setting in $settings.GetEnumerator())
            {
                $sql = "$sql $($setting.Key) = '$($setting.Value)', "
            }
        
            # Chop off the trailing ", "
            $sql = $sql.Substring(0, $sql.Length - 2)

            # Add the where clause
            $sql = "$sql WHERE ID = $id AND Type = 'L'"
        
            # Execute the command
            Write-Verbose "About to execute command: $sql"        
            $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
            $null = $settingsCmd.ExecuteScalar()
            
            Write-Verbose "Added settings for the specified location"
        }
        
        # If there are some gateways save them
        if ($gateways -ne $null)
        {
            # Build the delete command to remove the existing gateways
            $delCommand = "DELETE FROM LocationIdentity_DefaultGateway WHERE ID = $id"
        
            # Issue the delete command
            Write-Verbose "About to issue command: $delCommand"
            $cmd = New-Object System.Data.SqlClient.SqlCommand($delCommand, $mdtSQLConnection)
            $null = $cmd.ExecuteScalar()
            
            # Now insert the specified values
            foreach ($gateway in $gateways)
            {
                # Insert the  row
                $sql = "INSERT INTO LocationIdentity_DefaultGateway (ID, DefaultGateway) VALUES ($id, '$gateway')"
                Write-Verbose "About to execute command: $sql"
                $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
                $null = $settingsCmd.ExecuteScalar()

            }
            Write-Verbose "Set the default gateways for the location with ID = $id."    
        }
        
        # Write the updated record back to the pipeline
        Get-MDTLocation -ID $id
    }
}

function Remove-MDTLocation {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        # Build the delete command
        $delCommand = "DELETE FROM LocationIdentity WHERE ID = $id"
        
        # Issue the delete command
        Write-Verbose "About to issue command: $delCommand"
        $cmd = New-Object System.Data.SqlClient.SqlCommand($delCommand, $mdtSQLConnection)
        $null = $cmd.ExecuteScalar()

        Write-Verbose "Removed the location with ID = $id."
    }
}

function Get-MDTLocationApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'L' 'Settings_Applications' 'Applications'
    }
}

function Clear-MDTLocationApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'L' 'Settings_Applications'
    }
}

function Set-MDTLocationApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $applications
    )

    Process
    {
        Set-MDTArray $id 'L' 'Settings_Applications' 'Applications' $applications
    }
}

function Get-MDTLocationPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'L' 'Settings_Packages' 'Packages'
    }
}

function Clear-MDTLocationPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'L' 'Settings_Packages'
    }
}

function Set-MDTLocationPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $packages
    )

    Process
    {
        Set-MDTArray $id 'L' 'Settings_Packages' 'Packages' $packages
    }
}

function Get-MDTLocationRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'L' 'Settings_Roles' 'Role'
    }
}

function Clear-MDTLocationRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'L' 'Settings_Roles'
    }
}

function Set-MDTLocationRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $roles
    )

    Process
    {
        Set-MDTArray $id 'L' 'Settings_Roles' 'Role' $roles
    }
}

function Get-MDTLocationAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'L' 'Settings_Administrators' 'Administrators'
    }
}

function Clear-MDTLocationAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'L' 'Settings_Administrators'
    }
}

function Set-MDTLocationAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $administrators
    )

    Process
    {
        Set-MDTArray $id 'L' 'Settings_Administrators' 'Administrators' $administrators
    }
}

# ---------------------------------------------------------------------
# Make Model functions
# ---------------------------------------------------------------------

function New-MDTMakeModel {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $make,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $model,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $settings
    )

    Process
    {
        # Insert a new role row and get the identity result
        $sql = "INSERT INTO MakeModelIdentity (Make, Model) VALUES ('$make', '$model') SELECT @@IDENTITY"
        Write-Verbose "About to execute command: $sql"
        $identityCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $identity = $identityCmd.ExecuteScalar()
        Write-Verbose "Added make model identity record"
    
        # Insert the settings row, adding the values as specified in the hash table
        $settingsColumns = $settings.Keys -join ","
        $settingsValues = $settings.Values -join "','"
        $sql = "INSERT INTO Settings (Type, ID, $settingsColumns) VALUES ('M', $identity, '$settingsValues')"
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified make model"
        
        # Write the new record back to the pipeline
        Get-MDTMakeModel -ID $identity
    }
}

function Get-MDTMakeModel {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $id = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $make = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $model = ""
    )
    
    Process
    {
        # Build a select statement based on what parameters were specified
        if ($id -eq "" -and $make -eq "" -and $model -eq "")
        {
            $sql = "SELECT * FROM MakeModelSettings"
        }
        elseif ($id -ne "")
        {
            $sql = "SELECT * FROM MakeModelSettings WHERE ID = $id"
        }
        elseif ($make -ne "" -and $model -ne "")
        {
            $sql = "SELECT * FROM MakeModelSettings WHERE Make = '$make' AND Model = '$model'"
        }
        elseif ($make -ne "")
        {
            $sql = "SELECT * FROM MakeModelSettings WHERE Make = '$make'"
        }
        else
        {
            $sql = "SELECT * FROM MakeModelSettings WHERE Model = '$model'"
        }
        
        # Execute the statement and return the results    
        $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
        $selectDataset = New-Object System.Data.Dataset
        $null = $selectAdapter.Fill($selectDataset, "MakeModelSettings")
        $selectDataset.Tables[0].Rows
    }
}

function Set-MDTMakeModel {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(Mandatory=$true)] $settings
    )
    
    Process
    {
        # Add each each hash table entry to the update statement
        $sql = "UPDATE Settings SET"
        foreach ($setting in $settings.GetEnumerator())
        {
            $sql = "$sql $($setting.Key) = '$($setting.Value)', "
        }
        
        # Chop off the trailing ", "
        $sql = $sql.Substring(0, $sql.Length - 2)

        # Add the where clause
        $sql = "$sql WHERE ID = $id AND Type = 'M'"
        
        # Execute the command
        Write-Verbose "About to execute command: $sql"        
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
            
        Write-Verbose "Added settings for the specified make model"
        
        # Write the updated record back to the pipeline
        Get-MDTMakeModel -ID $id
    }
}

function Remove-MDTMakeModel {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        # Build the delete command
        $delCommand = "DELETE FROM MakeModelIdentity WHERE ID = $id"
        
        # Issue the delete command
        Write-Verbose "About to issue command: $delCommand"
        $cmd = New-Object System.Data.SqlClient.SqlCommand($delCommand, $mdtSQLConnection)
        $null = $cmd.ExecuteScalar()

        Write-Verbose "Removed the make model with ID = $id."
    }
}

function Get-MDTMakeModelApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'M' 'Settings_Applications' 'Applications'
    }
}

function Clear-MDTMakeModelApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'M' 'Settings_Applications'
    }
}

function Set-MDTMakeModelApplication {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $applications
    )

    Process
    {
        Set-MDTArray $id 'M' 'Settings_Applications' 'Applications' $applications
    }
}

function Get-MDTMakeModelPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'M' 'Settings_Packages' 'Packages'
    }
}

function Clear-MDTMakeModelPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'M' 'Settings_Packages'
    }
}

function Set-MDTMakeModelPackage {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $packages
    )

    Process
    {
        Set-MDTArray $id 'M' 'Settings_Packages' 'Packages' $packages
    }
}

function Get-MDTMakeModelRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'M' 'Settings_Roles' 'Role'
    }
}

function Clear-MDTMakeModelRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'M' 'Settings_Roles'
    }
}

function Set-MDTMakeModelRole {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $roles
    )

    Process
    {
        Set-MDTArray $id 'M' 'Settings_Roles' 'Role' $roles
    }
}

function Get-MDTMakeModelAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Get-MDTArray $id 'M' 'Settings_Administrators' 'Administrators'
    }
}

function Clear-MDTMakeModelAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id
    )

    Process
    {
        Clear-MDTArray $id 'M' 'Settings_Administrators'
    }
}

function Set-MDTMakeModelAdministrator {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $id,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $administrators
    )

    Process
    {
        Set-MDTArray $id 'M' 'Settings_Administrators' 'Administrators' $administrators
    }
}


# ---------------------------------------------------------------------
# Package mapping functions
# ---------------------------------------------------------------------

function New-MDTPackageMapping {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $ARPName,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $package
    )

    Process
    {
        # Insert a new row
        $sql = "INSERT INTO PackageMapping (ARPName, Packages) VALUES ('$ARPName','$package')"
        Write-Verbose "About to execute command: $sql"
        $identityCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $identityCmd.ExecuteScalar()
        Write-Verbose "Added package mapping record for $ARPName"
    
        # Write the new record back to the pipeline
        Get-MDTPackageMapping -ARPName $ARPName
    }
}

function Get-MDTPackageMapping {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $ARPName = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $package = ""
    )
    
    Process
    {
        # Build a select statement based on what parameters were specified
        if ($ARPName -eq "" -and $package -eq "")
        {
            $sql = "SELECT * FROM PackageMapping"
        }
        elseif ($ARPName -ne "" -and $package -ne "")
        {
            $sql = "SELECT * FROM PackageMapping WHERE ARPName = '$ARPName' AND Packages = '$package'"
        }
        elseif ($ARPName -ne "")
        {
            $sql = "SELECT * FROM PackageMapping WHERE ARPName = '$ARPName'"
        }
        else
        {
            $sql = "SELECT * FROM PackageMapping WHERE Packages = '$package'"
        }
        
        # Execute the statement and return the results    
        $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
        $selectDataset = New-Object System.Data.Dataset
        $null = $selectAdapter.Fill($selectDataset, "PackageMapping")
        $selectDataset.Tables[0].Rows
    }
}

function Set-MDTPackageMapping {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $ARPName,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $package = $null
    )
    
    Process
    {
        # Update the row
        $sql = "UPDATE PackageMapping SET Packages = '$package' WHERE ARPName = '$ARPName'"
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
        Write-Verbose "Updated the package mapping record for $ARPName to install package $package."    
        
        # Write the updated record back to the pipeline
        Get-MDTPackageMapping -ARPName $ARPName
    }
}

function Remove-MDTPackageMapping {

    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $ARPName = "",
        [Parameter(ValueFromPipelineByPropertyName=$true)] $package = ""
    )
    
    Process
    {
        # Build a delete statement based on what parameters were specified
        if ($ARPName -eq "" -and $package -eq "")
        {
            # Dangerous, delete them all
            $sql = "DELETE FROM PackageMapping"
        }
        elseif ($ARPName -ne "" -and $package -ne "")
        {
            $sql = "DELETE FROM PackageMapping WHERE ARPName = '$ARPName' AND Packages = '$package'"
        }
        elseif ($ARPName -ne "")
        {
            $sql = "DELETE FROM PackageMapping WHERE ARPName = '$ARPName'"
        }
        else
        {
            $sql = "DELETE FROM PackageMapping WHERE Packages = '$package'"
        }
        
        # Execute the delete command
        Write-Verbose "About to execute command: $sql"
        $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
        $null = $settingsCmd.ExecuteScalar()
        Write-Verbose "Removed package mapping records matching the specified parameters."    
    }
}
