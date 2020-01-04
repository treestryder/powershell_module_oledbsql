
function Invoke-OledbSql {
    <#
    .SYNOPSIS
    Invokes SQL commands through an OleDB Connection.

    .DESCRIPTION
    When a Select is invoked, returns System.Management.Automation.PSCustomObject
    objects. When an Update, Insert or Delete is invoked, it returns the affected
    count.

    .PARAMETER SQL
    SQL command to be invoked.

    .PARAMETER Connection
    Any valid OleDB connection string or a System.Data.OleDb.OleDbConnection object.
    Strings are passed to New-OledbConnection to be resolved.
    Defaults to using the variable $Conn if set to a SQLConnection object.

    .PARAMETER Timeout
    Sets the command timeout in seconds.

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    A custom type can be set with TypeName parameter.

    .EXAMPLE
    Invoke-OledbSql 'select 1 as Ping' 'Provider=sqloledb; Data Source=ServerName; Initial Catalog=DatabaseName;Integrated Security=SSPI;'

    Tests SQL connectivity.

    .EXAMPLE
    $cs = 'Provider=sqloledb; Data Source=ServerName; Initial Catalog=DatabaseName;Integrated Security=SSPI;'
    $c = New-OledbConnection $cs
    Get-Content 'Update_statements_one_per_line.txt' | WHERE { (Invoke-OledbSql -Sql $_ -Connection $c -KeepOpen) -eq 0 } | Set-Content 'failed.txt' -PassThru
    $c.Close()


    .NOTES
    Though this will handle more than one SQL statement in a single call, it is not recommended.

    SEE ALSO
        about_OledbSql

    #>
        [CmdletBinding()]
        [OutputType([PsCustomObject])]
        param(
            [Parameter(Mandatory=$True, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,
                HelpMessage='Enter SQL command'
            )]
            [string[]]$Sql,

            [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,
                HelpMessage='Enter a type name for the object returned
            ')]
            [string[]]$TypeName,

            [Parameter(HelpMessage='Enter an OLEDB connection string')]
            $Connection,

            [int]$Timeout,

            [switch]$KeepOpen
        )


        begin {
            # Will use the variable $Conn if it is set in the current context.
            if ((-not $Connection) -and ($Conn -is [System.Data.OleDb.OleDbConnection])) {$Connection = $Conn}

            if ($Connection -isnot [System.Data.OleDb.OleDbConnection]) {
                $Connection = New-OledbConnection -Connection:$Connection
            }

            if ($Connection -isnot [System.Data.OleDb.OleDbConnection]) {
                throw "No OleDB connection or string provided."
            }

            function Uniqueify ([string[]]$Columns) {
                <#
                .SYNOPSIS
                    Helper funciton that appends incrementing numbers to any duplicate
                    values in an incoming array.
                #>
                $rtn = @()
                for ($i = 0; $i -lt $Columns.Length; $i++) {
                    $cnt = 0
                    for ($j = 0; $j -lt $Columns.Length -and $j -lt $i; $j++) {
                        if ($Columns[$i] -eq $Columns[$j]) { $cnt++ }
                    }
                    if ($cnt -gt 0) {
                        $rtn += ('{0}{1}' -f $Columns[$i], $cnt)
                    } else {
                        $rtn += $Columns[$i]
                    }
                }
                return $rtn
            }

        }


        process {
            # Handles an array or piped list of sql statements.
            Foreach ($s in $SQL) {
                $command = New-Object System.Data.OleDb.OleDbCommand $s, $Connection
                if ($Timeout) {$command.CommandTimeOut = $Timeout}
                try {
                    # If the connection started open, keep it open.
                    if ($Connection.State -eq 'Open') {
                        $KeepOpen = $true
                    } else {
                        $Connection.Open()
                    }

                    Write-Verbose ('SQL: {0}' -f $s)
                    [System.Data.OleDB.OleDbDataReader]$reader = $command.ExecuteReader()

                    # while $reader.NextResult(), handles multiple queries/results in a single SQL request.
                    do {
                        # For scalar values, return the records affected.
                        if ( -not ($reader.HasRows)) { return $reader.RecordsAffected }

                        # Create an object template, with the uniqueified column names
                        # as properties and a custom PSTypename if specified.
                        #TODO: Look at maybe using $reader.GetSchemaTable()
                        $columns = 0 .. ($reader.VisibleFieldCount -1) | ForEach-Object {$reader.GetName($_)}
                        $columns = @(Uniqueify $columns)
                        $columnHash = [ordered]@{}
                        if ($TypeName) { $columnHash['Pstypename'] = $TypeName }
                        $columns | ForEach-Object { $columnHash[$_] = $null }
                        $objectTemplate = [pscustomobject]$columnHash

                        while ($reader.Read()) {
                            $o = $objectTemplate.PsObject.Copy()
                            for ($i=0; $i -lt $columns.Count; $i++) {
                                $column = $columns[$i]
                                Write-Debug ('{0} [{1}] ({2}) = [{3}]' -f $i, $column, $reader.GetFieldType($i), $reader.GetValue($i))
                                $o.$column = if ($reader.IsDBNull($i)) { $null } else { $reader.GetValue($i) }
                            }
                            Write-Output $o
                        }
                    } while ( $reader.NextResult() )

                    if ($reader -and (-not $reader.IsClosed)) {$reader.Close()}
                }
                catch {
                    Write-Verbose ('Message: {0}' -f $_.exception.InnerException.message)
                    throw $_
                }
                finally {
                    $reader = $null
                    $command = $null
                    if (-not $KeepOpen) {
                        $Connection.Close()
                        $Connection = $null
                    }
                }
            }
        }
    }
