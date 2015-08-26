#delete files and (folders(+subfolders)) by Maxzor1908 on creation date. 
#folder 
$Now = Get-Date 
$Days = "10" 
$TargetFolder = "C:\todelete" 
$LastWrite = $Now.AddDays(-$days) 
$Files = get-childitem $TargetFolder -include *.*  -recurse -force 
     Where {$_.CreationTime -le "$LastWrite"}  
    foreach ($i in Get-ChildItem $TargetFolder -recurse) 
{ 
    if ($i.CreationTime -lt ($(Get-Date).AddDays(-10))) 
    { 
        Remove-Item $Files -recurse -force 
    } 
} 
    Write-Output $Files >> c:\powershell\.delete.log 