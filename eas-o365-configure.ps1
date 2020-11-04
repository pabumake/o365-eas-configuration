
$cfgx86 = "C:\Program Files (x86)\Microsoft Office\root\Office16\OLCFG.EXE"
$cfgx64 = "C:\Program Files\Microsoft Office\root\Office16\OLCFG.EXE"

if(Test-Path -Path $cfgx86){
    Start-Process -FilePath $cfgx86
} elseif(Test-Path -Path $cfgx64){
    Start-Process -FilePath $cfgx64
} else {
    Write-Host "Seems like there is no Office16 Folder in either x86 nor x64 Paths. Aborting."
}
Exit-PSSession