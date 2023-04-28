Remove-Variable * -ErrorAction SilentlyContinue
$ProgressPreference = "SilentlyContinue"
$curdir = Get-Location
Write-Output "Current working directory: $curdir"

$index_template = Get-Content -Path .\index_template.html -Encoding UTF8 -Raw


$index_template = $index_template -split '<!---PLACEHOLDER-->'

$body_content = ""

if (Test-Path "nuovi_eventi.html"){
    $nuovi_eventi = Get-Content -Path .\nuovi_eventi.html -Encoding UTF8 -Raw
    $body_content += "<h1>Nuovi Eventi</h1>`n" + $nuovi_eventi + "`n"
    Remove-Item -Path ".\nuovi_eventi.html"
}

$body_content += '<h1>Elenco Campetti</h1>' + "`n"

if (Test-Path "buonacaccia_comp_mod.html"){
    $body_content += '<p><a title="Eventi Competenza" href="buonacaccia_comp_mod.html">Eventi Competenza</a></p>' + "`n"
}

if (Test-Path "buonacaccia_spec_mod.html"){
    $body_content += '<p><a title="Eventi Specialità" href="buonacaccia_spec_mod.html">Eventi Specialit&agrave</a></p>' + "`n"
}


$index_source = $index_template[0] + $body_content + $index_template[1]

Write-Output "Writing index.html"
$index_source | Out-File -FilePath .\index.html -Encoding UTF8

