function New-OledbConnection {
    <#
    .SYNOPSIS
    Returns a new System.Data.OleDb.OleDbConnection object.

    .DESCRIPTION
    Returns a new System.Data.OleDb.OleDbConnection object given a connection
    string or alias found in a file. Defaults to Module_Root\SQLConnections.txt.
    Queries the user for input when a connection string has parameters specified
    between percent signs are found, %Example Parameter%.

    .PARAMETER ConnectionString
    Any valid OleDB connection string. Prompts a user for input if a paramter
    specified between percent signs is found, %Example Parameter%.

    .PARAMETER List
    List the alias names and their connection string values found in the
    connection file.

    .PARAMETER File
    File used for resolving aliases to the full connection string.
    Defaults to SQLConnections.txt.

    .OUTPUTS
    System.Data.OleDb.OleDbConnection

    .NOTES

    SEE ALSO
        about_OledbSql

    #>
        [CmdletBinding(
            DefaultParameterSetName="receive"
        )]
        [OutputType([System.Data.OleDb.OleDbConnection])]
        param(
            [Parameter(
                Mandatory=$true,
                ParameterSetName="receive",
                ValueFromPipeline=$true,
                Position=0,
                HelpMessage='Enter an OLEDB connection string'
            )]
            [String]$ConnectionString,

            [Parameter(
                ParameterSetName="give"
            )]
            [Switch]$List,
            [String]$File = (Join-Path $PSScriptRoot Connections.txt)
        )
    <#
        Parse optional connection string file with the following format:
          # Comment
          alias        OleDB Connection String
    #>
        $connectionRegex = "^\s*(?<Alias>[^#]\w+)\s+(?<ConnectionString>.+)$"
        $Connections = @{}
        if (Test-Path $File) {
            Get-Content $File |
                Where-Object { $_ -match $connectionRegex } |
                    ForEach-Object { $Connections[$Matches['Alias']] = $Matches['ConnectionString'] }
        }
        if ($PsCmdlet.ParameterSetName -eq 'give' -and $List) { return $Connections }

        # Query user to replace any %variables%
        if ($Connections -is [Hashtable] -and $Connections.ContainsKey($ConnectionString)) {
            $ConnectionString = $Connections.Item($ConnectionString)
        }
        while ($ConnectionString -match "%([^%\s][^%]*)%") {
            $ConnectionString = $ConnectionString -replace $Matches[0],$(Read-Host $Matches[1])
        }

        if (-not $ConnectionString) {
            throw "No OleDB connection string specified."
        }
        Write-Verbose "Connection string: $ConnectionString"
        new-object System.Data.OleDb.OleDbConnection -ArgumentList $ConnectionString
    }
