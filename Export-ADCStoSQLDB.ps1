 
#Script Exports ADCS Database from locally from a CA and writes to database (could stop at exporting to CSV)
#This needs to be run on the CA Server and in a security context that has permissions to write to the SQL DB
#The script only exports Certificates where Issued Email Address is not empty  this behavior is easily changeable
#Script stores the highest (filtered) Request ID Exported / written to database in a file called WaterMarkFile 
#and only exports IDs higher than that the next time
#Why do this? SO you can query your DB for certs by team or have an email job that emails a report of expiring certs to the owner


$hWatermarkFile = "C:\scripts\hwatermark.txt"
$OutFile  = "C:\scripts\file1.csv"
[int]$hwm = Get-Content -Path $hWatermarkFile
$DBServerInstance = 'MYDBSERVER'
$DBName = 'InvDB'

Function Write-DebugLog{
PARAM(
    [string]$Message,
    [string]$file = "c:\Scripts\Logs\ADCSExport.log"
)
    $dt = Get-date -Format u 
    "$dt $Message"| Out-File -FilePath $file -Append
    Write-Debug "$dt $Message"
}



<#
.Synopsis
   Exports certificates from ADCS Database with Disposition Greater than 20
.DESCRIPTION
   Exports all certificates with RequestID Greater than  HighWaterMark to Outfile
.EXAMPLE
   Export-CertsDB -HighWaterMark $hwm -OutFile $OutFile
.EXAMPLE
   Export-CertsDB -HighWaterMark 83580 -OutFile "C:\scripts\file1.csv"
#>
Function Export-CertsDB{
Param(
    [Parameter(Mandatory=$true)]
    [int] $HighWaterMark = 0,
    [Parameter(Mandatory=$true)]
    [string] $OutFile
    
)

$cmd = @"
certutil -view -out Request.RequestID,SerialNumber,Request.CommonName,CommonName,NotBefore,NotAfter,CertificateHash,CertificateTemplate,Request.EndorsementKeyHash,Request.Disposition,Request.DispositionMessage,Request.EMail,Email -Restrict "Disposition>=20,Request.RequestID>$HighWaterMark" csv 
"@
    Invoke-Expression -Command $cmd | Out-File -FilePath $OutFile -Force
}
Function Get-CertDBcsvContent{
Param(
    [Parameter(Mandatory=$true)]
    [string] $csvFile
)
    Import-Csv -path $csvFile
}

Function Get-MaxID{
Param(
    [Parameter(Mandatory=$true)]
    $CertArray
)
    $maxID = $($CertArray.'Request ID') | Measure-object -Maximum
    return $($maxID.Maximum)
}

FUNCTION Invoke-Sqlcmd2 { 
<# 
.SYNOPSIS 
Runs a T-SQL script. 
.DESCRIPTION 
Runs a T-SQL script. Invoke-Sqlcmd2 only returns message output, such as the output of PRINT statements when -verbose parameter is specified 
.INPUTS 
None 
    You cannot pipe objects to Invoke-Sqlcmd2 
.OUTPUTS 
   System.Data.DataTable 
.EXAMPLE 
Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -Query "SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1" 
This example connects to a named instance of the Database Engine on a computer and runs a basic T-SQL query. 
StartTime 
----------- 
2010-08-12 21:21:03.593 
.EXAMPLE 
Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -InputFile "C:\MyFolder\tsqlscript.sql" | Out-File -filePath "C:\MyFolder\tsqlscript.rpt" 
This example reads a file containing T-SQL statements, runs the file, and writes the output to another file. 
.EXAMPLE 
Invoke-Sqlcmd2  -ServerInstance "MyComputer\MyInstance" -Query "PRINT 'hello world'" -Verbose 
This example uses the PowerShell -Verbose parameter to return the message output of the PRINT command. 
VERBOSE: hello world 
.NOTES 
Version History 
v1.0   - Chad Miller - Initial release 
v1.1   - Chad Miller - Fixed Issue with connection closing 
v1.2   - Chad Miller - Added inputfile, SQL auth support, connectiontimeout and output message handling. Updated help documentation 
v1.3   - Chad Miller - Added As parameter to control DataSet, DataTable or array of DataRow Output type 
#> 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance, 
    [Parameter(Position=1, Mandatory=$false)] [string]$Database, 
    [Parameter(Position=2, Mandatory=$false)] [string]$Query, 
    [Parameter(Position=3, Mandatory=$false)] [string]$Username, 
    [Parameter(Position=4, Mandatory=$false)] [string]$Password, 
    [Parameter(Position=5, Mandatory=$false)] [Int32]$QueryTimeout=600, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$ConnectionTimeout=15, 
    [Parameter(Position=7, Mandatory=$false)] [ValidateScript({test-path $_})] [string]$InputFile, 
    [Parameter(Position=8, Mandatory=$false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As="DataRow" 
    ) 
 
    if ($InputFile) 
    { 
        $filePath = $(resolve-path $InputFile).path 
        $Query =  [System.IO.File]::ReadAllText("$filePath") 
    } 
 
    $conn=new-object System.Data.SqlClient.SQLConnection 
      
    if ($Username) 
    { $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout } 
    else 
    { $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout } 
    #$ConnectionString
    $conn.ConnectionString=$ConnectionString 
     
    #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller 
    if ($PSBoundParameters.Verbose) 
    { 
        $conn.FireInfoMessageEventOnUserErrors=$true 
        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {Write-Verbose "$($_)"} 
        $conn.add_InfoMessage($handler) 
    } 
    try{ 
    $conn.Open() 
    
    $cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn) 
    $cmd.CommandTimeout=$QueryTimeout 
    $ds=New-Object system.Data.DataSet 
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 
    [void]$da.fill($ds) 
    $conn.Close() 
    switch ($As) 
    { 
        'DataSet'   { Write-Output ($ds) } 
        'DataTable' { Write-Output ($ds.Tables) } 
        'DataRow'   { Write-Output ($ds.Tables[0]) } 
    } 
    }
    catch{
    Write-Error $error[0]
    }
} 
Function Get-MDY{
    Get-Date -Format yyMMdd
}
Function Write-CertsToDB{
PARAM(
    $certArray,
    $ServerInstance,
    $DBName,
    [switch] $WhatIf
)
    if($certArray.count){
        foreach($cert in $certArray){
            $template = $cert.'Certificate Template' -replace '\d+\.','' -replace '^\d+\s',''
$qry = @"
insert INTO certs (reqid,SerialNum, CN,validfrom,validto,thumbprint,email,template,DispMsg)
values('$($cert.'Request ID')', '$($cert.'Serial Number')', '$($cert.'Issued Common Name')', 
'$($cert.'Certificate Effective Date')', '$($cert.'Certificate Expiration Date')',
    '$($cert.'Certificate Hash')', '$($cert.'Issued Email Address')', '$template','$($cert.'Request Disposition Message')')
"@
            Write-Debuglog -Message $qry
            if(-not $WhatIf){
                Invoke-Sqlcmd2 -ServerInstance $ServerInstance -Database $DBName -Query $qry
            }
        } #End foreach
    } #End of If
}#End of Function







Write-DebugLog -Message "Getting Requests with ID Greater than $hwm"
Remove-Item $OutFile
#Export to $outfile
Export-CertsDB -HighWaterMark $hwm -OutFile $OutFile
#Read contents of $outfile 
$certexport = @()
$certexport += Get-CertDBcsvContent -csvFile $OutFile | Where-Object {($_.'Issued Email Address' -ne 'EMPTY') -and $_.'Request Disposition' -eq '20 -- Issued'}
if(($certexport.count) -ge 1){
    Write-DebugLog -Message "Got $($certexport.Count) Certs to upload to the database"
    Write-CertsToDB -certArray $certexport -ServerInstance $DBServerInstance -DBName $DBName 
    foreach($cert in $certexport){
        Write-DebugLog -Message "$($cert.'Issued Common Name')"
    }
    $MaxID = Get-MaxID -CertArray $certexport
    Write-DebugLog -Message "highest ID in the list is $MaxID"
    $maxID | Out-File -FilePath $hWatermarkFile
}
