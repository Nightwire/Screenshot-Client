function Expand-ZIPFile($file, $destination)
{
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
}

$path = $env:LocalAppData + "\Niklas_Mollenhauer\"
$apikey = Read-Host("Bitte API Key eingeben")
$pluginPath = $env:AppData + "\HolzShots\Plugin\"
$ShortcutPath = ($env:AppData + "\Microsoft\Windows\Start Menu\Programs\Startup\HolzShots.lnk")

#Prepare Directorys
Write-Host("Erstelle HolzShots Programm Ordner")
Remove-Item ($env:LocalAppData + "\HolzShots") -recurse
$HZPath = (New-Item -ItemType Directory -Force -Path ($env:LocalAppData + "\HolzShots")).FullName
$HolzShotsBinPath = $HZPath + "\Holzshots.exe"

Write-Host("Erstelle HolzShots Plugin Ordner")
Remove-Item $pluginPath -recurse
New-Item -ItemType Directory -Force -Path $pluginPath

Write-Host("Erstelle HolzShots Settings Ordner")
Remove-Item $path -recurse
New-Item -ItemType Directory -Force -Path $path

#Download Holzshots
Write-Host("Lade HolzShots herunter..")
$url = "https://gitlab.com/nightwire/NWDE-Holzshots/raw/49e3085b851b63c5c9cf3b5d85c692ffda79157d/HolzShots.zip"
$output = "$HZPath\HolzShots.zip"

Invoke-WebRequest -Uri $url -OutFile $output

#Unzip Holzshots
Write-Host("Entpacke HolzShots..")
Expand-ZIPFile -File $output -Destination $HZPath

#Delete Zip
Write-Host("Lösche HolzShots Archiv..")
Remove-Item -Force $output

#Add Holzshots to autostart
Write-Host("Füge Holzshots zum Autostart hinzu..")
If (Test-Path $ShortcutPath){
	Remove-Item $ShortcutPath
}
$wshshell = New-Object -comObject WScript.Shell
$link = $wshshell.CreateShortcut($ShortcutPath)
$link.targetpath = $HolzShotsBinPath
$link.save()

#Copy CustomPost Plugin
Write-Host("Kopiere Custom Post Plugin..")
Copy-Item -Force ($HZPath + "\CustomPostUpload.dll") -Destination ($pluginPath + "\CustomPostUpload.dll")

#Start Holzshots once, wait a few seconds and kill it right away
Write-Host("Starte HolzShots, um entsprechenden Config Pfad zu erkennen.")

$Process = [Diagnostics.Process]::Start($HolzShotsBinPath)
$id = $Process.Id

Write-Host("Warte 10 Sekunden..")
Start-Sleep -Seconds 10

try {            
    Stop-Process -Id $id -ErrorAction stop            
    Write-Host "Erfolgreich Holzshots mit der ID: $id gestoppt."            
} catch {            
    Write-Host "FEHLER: Bitte Holzshots manuell neustarten."            
}

#Add Config
Get-ChildItem $path | ForEach-Object{

    Remove-Item $_.FullName -recurse     
    $ConfigFolder = New-Item -ItemType directory -Path ($_.FullName + "\" + "0.9.8.12")
        
    $configPath = $ConfigFolder.FullName + "\user.config"

    Write-Host("Lade HolzShots Config herunter..")
    $url = "https://gitlab.com/nightwire/NWDE-Holzshots/raw/master/user.config"
    Invoke-WebRequest -Uri $url -OutFile $configPath
	(Get-Content $configPath) | 
	Foreach-Object {$_ -replace 'INSERTAPIKEYHERE',$apikey}  | 
	Out-File $configPath -Encoding utf8
    Write-Host("HolzShots Config erfolgreich unter $configPath gespeichert")
    
}

Write-Host("Installation erfolgreich, starte Holzshots.")
[Diagnostics.Process]::Start($HolzShotsBinPath)