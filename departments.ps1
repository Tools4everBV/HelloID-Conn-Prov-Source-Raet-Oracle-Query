$config = ConvertFrom-Json $configuration

$DataSource = $config.dataSource
$Username = $config.username
$Password = $config.password

$OracleConnectionString = "User Id=$Username;Password=$Password;Data Source=$DataSource"

function Get-OracleDBData
{
    param(
        [parameter(Mandatory=$true)]
        $OracleConnectionString,

        [parameter(Mandatory=$true)]
        $OracleQuery,

        [parameter(Mandatory=$true)]
        [ref]$Data
    )
    try{
        $Data.value = $null

        # Initialize connection and execute query
        # Connect to the Oracle server
        $null =[Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")

        $OracleConnection = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
        $OracleConnection.Open()
        Write-Verbose -Verbose "Successfully connected Oracle to database '$DataSource'" 

        # Execute the command against the database, returning results.
        $OracleCmd = $OracleConnection.CreateCommand()
        $OracleCmd.CommandText = $OracleQuery

        $OracleAdapter = New-Object System.Data.OracleClient.OracleDataAdapter($cmd)
        $OracleAdapter.SelectCommand = $OracleCmd;

        # Execute the command against the database, returning results.
        $DataSet = New-Object system.Data.DataSet
        $null = $OracleAdapter.fill($DataSet)

        $Data.value =  $DataSet.Tables[0] | Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors;
        Write-Verbose -Verbose "Successfully performed Oracle query. Returned [$($DataSet.Tables[0].Columns.Count)] columns and [$($DataSet.Tables[0].Rows.Count)] rows"
    } catch {
        $Data.Value = $null
        Write-Error $_
    }finally{
        if($OracleConnection.State -eq "Open"){
            $OracleConnection.close()
        }
        Write-Verbose -Verbose "Successfully disconnected from Oracle database '$DataSource'"
    }
}


try{
    # Get Department data
    $OracleQuery = "SELECT 
            CAST(ou.dpib015_sl AS int) AS afdeling,
            ou.orgeenh_kd,
            ou.oe_kort_nm,
            ou.oe_vol_nm,
            CAST(ou.oe_hoger_n AS int) AS oe_hoger_n,
            CAST(m.pers_nr AS int) as pers_nr,
            m.rol_oe_kd
        FROM dpib015 ou
            LEFT JOIN dpib025 m ON m.dpib015_sl = ou.dpib015_sl
                AND m.rol_oe_kd = 'MGR'
                AND m.ingang_dt < CURRENT_TIMESTAMP
                AND (m.eind_dt IS NULL OR m.eind_dt > CURRENT_TIMESTAMP)
    "

    $departments = New-Object System.Collections.ArrayList
    Get-OracleDBData -OracleConnectionString $OracleConnectionString -OracleQuery $OracleQuery -Data ([ref]$departments)

    # Extend the departments with required and additional fields
    $departments | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
    $departments | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $null -Force
    $departments | Add-Member -MemberType NoteProperty -Name "Name" -Value $null -Force
    $departments | Add-Member -MemberType NoteProperty -Name "ManagerExternalId" -Value $null -Force
    $departments | Add-Member -MemberType NoteProperty -Name "ParentExternalId" -Value $null -Force
    $departments | ForEach-Object {
        $_.ExternalId = $_.afdeling
        $_.DisplayName = $_.oe_vol_nm
        $_.Name = $_.oe_vol_nm
        $_.ManagerExternalId = $_.pers_nr
        $_.ParentExternalId = $_.oe_hoger_n
    }

    # Make sure departments are unique
    $departments = $departments | Sort-Object ExternalId -Unique

    # Export and sanitize the persons in json format
    foreach($department in $departments){
        $json = $department | ConvertTo-Json -Depth 10
        $json = $json.Replace("._", "__")
        Write-Output $json
    }

    Write-Verbose -Verbose "Department import completed";
}catch{
    Write-Error $_
}