# install MailKit through: Install-Package -Name 'MailKit' -Source 'nuget.org'
function GetColorCode {

    param (
        $color1,
        $color2,
        [float]$weight
    )

    $w1 = $weight;
    $w2 = 1 - $w1;
    $rgb = 0,0,0
    $rgb[0] = [int32]($color1[0] * $w1 + $color2[0] * $w2)
    $rgb[1] = [int32]($color1[1] * $w1 + $color2[1] * $w2)
    $rgb[2] = [int32]($color1[2] * $w1 + $color2[2] * $w2)

    return $rgb
}

Remove-Variable * -ErrorAction SilentlyContinue
$ProgressPreference = "SilentlyContinue"
Invoke-WebRequest "https://buonacaccia.net/Events.aspx?RID=F&CID=21" -OutFile .\buonacaccia_spec.html
$camps_path = "$PSScriptRoot/buonacaccia_spec.json"

# $content = Get-Content -Encoding UTF8 -Path .\buonacaccia_spec.html
$content = [System.IO.File]::ReadAllText("$PSScriptRoot\buonacaccia_spec.html")
$template = [System.IO.File]::ReadAllText("$PSScriptRoot\output_template_spec.html")

$template = $template -split 'PLACEHOLDER'

$content = $content -replace '(<span id="MainContent_EventsGridView_Taken_\d+">\s*\d+\s*</span>).*?(<span id="MainContent_EventsGridView_SeatsMax_)', '$1 $2'
$content = $content -split '<img src="Images/light_.*?.png" /></td> </tr>'
# $content | Out-File -FilePath .\buonacaccia_spec_mod.html


$camps = [System.Collections.ArrayList]@()
$camp = @{}
for ($cont_sel = 0; $cont_sel -lt $content.count; $cont_sel++)
{
    $match = $content[$cont_sel] -match '<a href="(?<link>event\.aspx\?e=\d+)">(?<spec>.*?)\s*\|\s*(?<province2>.*?)</a></td>.*?(?<datestart>\d+/\d+/\d+).*?(?<dateend>\d+/\d+/\d+).*?MainContent_EventsGridView_Fee_\d+">(?<fee>.*?) \u20ac.*?MainContent_EventsGridView_Location_\d+".*?>(?<location>.*?) \((?<province>\w+)\).*?MainContent_EventsGridView_Taken_\d+">(?<taken>\d+).*\n*.*?MainContent_EventsGridView_SeatsMax_\d+">(?<seatsmax>\d+)'
    if ($match){
        $camp = @{}
        $camp.seatsmax = [int32]$Matches.seatsmax
        $camp.taken = [int32]$Matches.taken
        $camp.availability = $camp.taken/$camp.seatsmax
        $camp.spec = $Matches.spec
        $camp.location = $Matches.location
        $camp.province = $Matches.province
        $camp.datestart = $Matches.datestart
        $camp.dateend = $Matches.dateend
        $camp.fee = $Matches.fee
        $camp.link = 'https://buonacaccia.net/' + $Matches.link
        [void]$camps.Add($camp);
    }

}

$camps = $camps | Sort-Object {$_.availability}


# check if an old version of json is available
$old_json_available = $false
if (Test-Path $camps_path){
    $old_json_available = $true
    $camps_old_string = [System.IO.File]::ReadAllText($camps_path)
    $camps_old = ConvertFrom-Json $camps_old_string
}

for ($camp_sel = 0; $camp_sel -lt $camps.count; $camp_sel++)
{
    $rgb = GetColorCode -color1 (255,0,0) -color2 (0,255,0) -weight ([math]::Min($camps[$camp_sel].availability-0.15, 1))

    $new_availability = $false
    $camp_old_found = $false
    if ($old_json_available){
        for ($camp_old_sel = 0; $camp_old_sel -lt $camps_old.count; $camp_old_sel++){
            if ($camps[$camp_sel].link -eq $camps_old[$camp_old_sel].link){
                $camp_old_found = $true
                if ($camps[$camp_sel].availability -gt $camps_old[$camp_old_sel].availability){
                    $new_availability = $true
                }
                break
            }
        }
    }
    if (-not $camp_old_found){$new_availability = $true}

    if ($new_availability){$new_aval_col = '"background-color:rgb(0, 255, 0);"'} else {$new_aval_col = '"background-color:rgb(255, 255, 255);"'}

    $newstr += '<tr>' +
               '<td><a href="' + $camps[$camp_sel].link + '">' + $camps[$camp_sel].spec+'</a></td>' +
               '<td>' + $camps[$camp_sel].location + '</td>' +
               '<td>' + $camps[$camp_sel].province + '</td>' +
               '<td>' + $camps[$camp_sel].datestart + '</td>' +
               '<td>' + $camps[$camp_sel].dateend + '</td>' +
               '<td>' + $camps[$camp_sel].fee + '</td>' +
               '<td style='+$new_aval_col+'>' + $camps[$camp_sel].taken + '</td>' +
               '<td style='+$new_aval_col+'>' + $camps[$camp_sel].seatsmax + '</td>' +
               '<td style="background-color:rgb(' + $rgb[0] + ', ' + $rgb[1] + ', ' + $rgb[2] + ');">' + $camps[$camp_sel].availability + '</td>' +
               '</tr>' + "`n"

    if ($new_availability){
        $new_availability_string += $newstr 
    }
}

$content_output = $template[0] + $newstr + $template[1]

$mail_content = $template[0] + $new_availability_string + $template[1]


# $ErrorActionPreference = "SilentlyContinue"
$camps_json = $camps | ConvertTo-Json -Depth 4 -ErrorAction SilentlyContinue
$camps_json | Out-File -FilePath .\buonacaccia_spec.json -Encoding UTF8


$content_output | Out-File -FilePath .\buonacaccia_spec_mod.html -Encoding UTF8
Remove-Item -Path "$PSScriptRoot\buonacaccia_spec.html"




