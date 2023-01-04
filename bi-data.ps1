$temp = Import-Excel D:\Privat\Applied-ML-main\temp.xlsx
$tryck = Import-Excel D:\Privat\Applied-ML-main\LuftTryck.xlsx
$fukt = Import-Excel D:\Privat\Applied-ML-main\LuftFuktighet.xlsx
$vikt = Import-Excel D:\Privat\Applied-ML-main\Vikt.xlsx

$temp = $null
$tryck = $null
$fukt =$null
$vikt = $null

#Get data file
$SqlServer = 'Localhost'
$User = 'sa'
$Pass = '6r!1l2$J$Mk%l0419'
$Database = 'BI-Data'
$Table = 'Temp'
#$import = 15
$Data = Import-Excel D:\Privat\Applied-ML-main\temp.xlsx
$x=1
$count = $data.Count
#Empty tblSouPerson
$truncate = 'truncate table '+$Table
Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Username $User -Password $pass -Query $truncate

foreach ($Item in $Data) { 
        write-host "Adding Item $x of $count"
        [DateTime]$Time = $Item.Time
        [decimal]$Value = $Item.Value

        $sqlCmd = @"
        INSERT INTO  $Table (
        [Time]
        ,[Value]
        )  
        VALUES ('$Time',$Value)
"@
  
        Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Username $User -Password $pass -Query $sqlCmd   
        $x++
}
3