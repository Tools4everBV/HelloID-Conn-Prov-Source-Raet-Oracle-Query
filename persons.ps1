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
    # Get Person data
    $OracleQuery = "SELECT 
	    tpersoon.pers_nr,
	    tpersoon.e_titul, 
	    tpersoon.e_titul_na, 
	    tpersoon.e_vrlt, 
	    tpersoon.e_roepnaam, 
	    tpersoon.e_voornmn, 
	    tpersoon.e_vrvg, 
	    tpersoon.e_naam, 
	    tpersoon.p_vrvg, 
	    tpersoon.p_naam,
	    tpersoon.vrvg_samen, 
	    tpersoon.naam_samen, 
	    tpersoon.gbrk_naam, 
	    tpersoon.tssnvgsl_kd, 
	    tpersoon.geslacht,
	    tpersoon.mobiel_tel_nr, 
	    tpersoon.werk_tel_nr, 
	    tpersoon.email_werk
	FROM dpib010 tpersoon, dpic300 tcontract, dpic351 tfunctie, dpib015 tafdeling
	WHERE tpersoon.pers_nr = tcontract.pers_nr
	    AND tcontract.primfunc_kd = tfunctie.func_kd
	    AND tcontract.oe_hier_sl = tafdeling.dpib015_sl
	ORDER BY tpersoon.pers_nr,dv_vlgnr desc
    "

    $persons = New-Object System.Collections.ArrayList
    Get-OracleDBData -OracleConnectionString $OracleConnectionString -OracleQuery $OracleQuery -Data ([ref]$persons)

    # Get Contract data
    $OracleQuery = "SELECT
	    tcontract.pers_nr,
	    tcontract.dv_vlgnr,
	    tcontract.opdrgvr_nr,
	    tcontract.oe_hier_sl AS afdeling,
	    tcontract.uren_pw,
	    tcontract.deelb_perc,
	    tcontract.indnst_dt,  
	    tcontract.uitdnst_dt,
	    tcontract.arelsrt_kd,
	    tcontract.opdrgvr_nr,
	    tcontract.object_id,
	    tafdeling.kstpl_kd,
	    tafdeling.kstdrg_kd,
	    tafdeling.oe_kort_nm,
	    tafdeling.oe_vol_nm,
	    tafdeling.oe_hoger_n,
	    tfunctie.func_kd,
	    tfunctie.func_oms,
	    tfunctie.funtyp_kd
	FROM dpic300 tcontract
	LEFT JOIN dpib015 tafdeling ON tcontract.oe_hier_sl = tafdeling.dpib015_sl 
	LEFT JOIN dpic351 tfunctie ON tcontract.primfunc_kd = tfunctie.func_kd
    "
    $employments = New-Object System.Collections.ArrayList
    Get-OracleDBData -OracleConnectionString $OracleConnectionString -OracleQuery $OracleQuery -Data ([ref]$employments)

    # Group the employments
    $employments = $employments | Group-Object PERS_NR -AsHashTable  

    # Extend the persons with positions and required fields
    $persons | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
    $persons | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $null -Force
    $persons | Add-Member -MemberType NoteProperty -Name "Contracts" -Value $null -Force
    # Map required fields
    $persons | ForEach-Object {
        $_.ExternalId = $_.PERS_NR
        $_.DisplayName = $_.PERS_NR
    }

    # Make sure persons are unique
    $persons = $persons | Sort-Object ExternalId -Unique

    # Enhance person object
    foreach($person in $persons){
        # Add Contracts to person object
        if(-not([string]::IsNullOrEmpty($person.PERS_NR))){
            $contracts = $employments[$person.PERS_NR]
            if ( -not($null -eq $contracts) ){
                $person.Contracts = $contracts
            }
        }

	# Convert naming convention codes to standard
        Switch($person.GBRK_NAAM ){
            "E" {
                $person.GBRK_NAAM = "B"
            }
            "P" {
                $person.GBRK_NAAM = "P"
            }
            "C" {
                $person.GBRK_NAAM = "BP"
            }
            "B" {
                $person.GBRK_NAAM = "PB"
            }
            "D" {
                $person.GBRK_NAAM = "BP"
            }
        }

        # Export and sanitize the persons in json format
        $json = $person | ConvertTo-Json -Depth 10
        $json = $json.Replace("._", "__")
        Write-Output $json
    }

    Write-Verbose -Verbose "Person import completed";
}catch{
    Write-Error $_
}
