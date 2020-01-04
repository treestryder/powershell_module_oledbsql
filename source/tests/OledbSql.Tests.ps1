
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

Get-Module OledbSql | Remove-Module
Import-Module (Join-Path $here '..\OledbSql.psd1')

$connectionString = 'Provider=MSPersist'

Describe 'New-OledbConnection' {
    It 'Create and open an OLEDB Connection.' {
        $connection = New-OledbConnection $connectionString
        $connection.Open()
        $connection.State | Should Be 'Open'
        $connection.Close()
    }
}

Describe 'Invoke-OledbSql' {
    Context 'Read from a file.' {
        $sqlString = (Join-Path $here 'test.adtg') -replace '\\','\\'
        $result = Invoke-OledbSql -Connection $connectionString -Sql $sqlString | Select-Object -First 1

        It 'Return the value of the first column.' {
            $result.'field' | Should Be 'value'
        }
    }
}