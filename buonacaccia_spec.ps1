
function GetColorCode {

    param (
        $color1,
        $color2,
        [float]$weight
    )

    $w1 = $weight;
    $w2 = 1 - $w1;
    $rgb_status = 0,0,0
    $rgb_status[0] = [int32]($color1[0] * $w1 + $color2[0] * $w2)
    $rgb_status[1] = [int32]($color1[1] * $w1 + $color2[1] * $w2)
    $rgb_status[2] = [int32]($color1[2] * $w1 + $color2[2] * $w2)

    return $rgb_status
}

Add-Type -Assembly System.Windows.Forms
Remove-Variable * -ErrorAction SilentlyContinue
$ProgressPreference = "SilentlyContinue"
Set-Location -Path "D:\GoogleDrive\Scout\RisorseVarie\buonacaccia_eventi\"

Write-Output "Retrieving web page"
Invoke-WebRequest "https://buonacaccia.net/Events.aspx?RID=F&CID=21" -OutFile ".\buonacaccia_spec.html"
$camps_path = ".\buonacaccia_spec.json"
$trigger_popup_newstatus = $true


$content = Get-Content -Path .\buonacaccia_spec.html -Encoding UTF8 -Raw
$template = Get-Content -Path "$PSScriptRoot\output_template_spec.html" -Encoding UTF8 -Raw


$template = $template -split 'PLACEHOLDER'

$content = $content -replace '(<span id="MainContent_EventsGridView_Taken_\d+">\s*\d+\s*</span>).*?(<span id="MainContent_EventsGridView_SeatsMax_)', '$1 $2'
$content = $content -split '<img src="Images/light_.*?.png" /></td> </tr>'

Write-Output "Creating results output"
$camps_unsorted = [System.Collections.ArrayList]@()
$camp = @{}
for ($cont_sel = 0; $cont_sel -lt $content.count; $cont_sel++)
{
    $match = $content[$cont_sel] -match '<a href="(?<link>event\.aspx\?e=\d+)">(?<spec>.*?)\s*\|\s*(?<province2>.*?)</a></td>.*?(?<datestart>\d+/\d+/\d+).*?(?<dateend>\d+/\d+/\d+).*?MainContent_EventsGridView_Fee_\d+">(?<fee>.*?) \u20ac.*?MainContent_EventsGridView_Location_\d+".*?>(?<location>.*?) \((?<province>\w+)\).*?MainContent_EventsGridView_Taken_\d+">(?<taken>\d+).*\n*.*?MainContent_EventsGridView_SeatsMax_\d+">(?<seatsmax>\d+)'
    if ($match){
        $camp = @{}
        $camp.seatsmax = [int32]$Matches.seatsmax
        $camp.taken = [int32]$Matches.taken
        $camp.spec = $Matches.spec
        $camp.sex = $Matches.sex
        $camp.location = $Matches.location
        $camp.province = $Matches.province
        $camp.base = $Matches.base
        $camp.datestart = $Matches.datestart
        $camp.dateend = $Matches.dateend
        $camp.fee = $Matches.fee
        $camp.link = 'https://buonacaccia.net/' + $Matches.link

        # auxiliary fields
        $camp.availability = $camp.taken/$camp.seatsmax

        if ($camp.seatsmax -gt $camp.taken){
            $camp.status = 'libero'
        }
        elseif ($camp.seatsmax+5 -gt $camp.taken)
        {
            $camp.status = 'coda'
        }
        else{
            $camp.status = 'pieno'
        }


        [void]$camps_unsorted.Add($camp);
    }

}

Write-Output "Sorting by availability"
$camps = [System.Collections.ArrayList]@()
$camps.AddRange($camps_unsorted.where({ $_.status -like 'libero'}))
$camps.AddRange($camps_unsorted.where({ $_.status -like 'coda'}))
$camps.AddRange($camps_unsorted.where({ $_.status -like 'pieno'}))
# $camps = $camps | Sort-Object {$_.availability}

# check if an old version of json is available
$old_json_available = $false
if (Test-Path $camps_path){
    Write-Output "Comparing with previous run"
    $old_json_available = $true
    $camps_old_string = Get-Content -Path $camps_path -Encoding UTF8 -Raw
    # $camps_old_string = [System.IO.File]::ReadAllText($camps_path)
    $camps_old = ConvertFrom-Json $camps_old_string
}

Write-Output "Generating info"
$camps_newstatus_str = ""
for ($camp_sel = 0; $camp_sel -lt $camps.count; $camp_sel++)
{

    if ($camps[$camp_sel].status -like "libero"){
        $rgb_status = (0,255,0)
    }
    elseif ($camps[$camp_sel].status -like "coda"){
        $rgb_status = (255,255,0)
    }
    elseif ($camps[$camp_sel].status -like "pieno"){
        $rgb_status = (255,0,0)
    }

    $new_availability = $false
    $new_status = $false
    $camp_old_found = $false
    if ($old_json_available){
        for ($camp_old_sel = 0; $camp_old_sel -lt $camps_old.count; $camp_old_sel++){
            if ($camps[$camp_sel].link -eq $camps_old[$camp_old_sel].link){
                $camp_old_found = $true
                if ($camps[$camp_sel].availability -gt $camps_old[$camp_old_sel].availability*1.01){
                    $new_availability = $true
                }
                if (($camps[$camp_sel].status -like "libero" -and $camps_old[$camp_old_sel].status -like "coda") -or ((-not $camps[$camp_sel].status -like "pieno") -and $camps_old[$camp_old_sel].status -like "pieno")){
                    $new_status = $true
                    $camps_newstatus_str += $camps[$camp_sel].spec + " (" + $camps[$camp_sel].province + "): " + $camps_old[$camp_old_sel].status.ToUpper() + "-->" + $camps[$camp_sel].status.ToUpper() + "`n"
                }
                break
            }
        }
    }
    
    if (-not $camp_old_found){
        $new_availability = $true
        $new_status = $true
        $camps_newstatus_str += $camps[$camp_sel].spec + " (" + $camps[$camp_sel].province + "): " + "NUOVO`n"
    }

    if ($new_availability){$new_aval_col = 'background-color:rgb(0, 255, 0)'} else {$new_aval_col = 'background-color:rgb(255, 255, 255)'}
    if ($new_status){$new_status_col = 'background-color:rgb(0, 255, 0)'} else {$new_status_col = 'background-color:rgb(255, 255, 255)'}


    $newstr += '<tr>' +
               '<td style="text-align:left;'+$new_status_col+';"><a href="' + $camps[$camp_sel].link + '">' + $camps[$camp_sel].spec+'</a></td>' +
               '<td style="text-align:center;background-color:rgb(' + $rgb_status[0] + ', ' + $rgb_status[1] + ', ' + $rgb_status[2] + ');">' + $camps[$camp_sel].status + '</td>' +
               '<td style="text-align:left;">' + $camps[$camp_sel].location + '</td>' +
               '<td style="text-align:center;">' + $camps[$camp_sel].province + '</td>' +
               '<td style="text-align:center;">' + $camps[$camp_sel].datestart + '</td>' +
               '<td style="text-align:center;">' + $camps[$camp_sel].dateend + '</td>' +
               '<td style="text-align:center;">' + $camps[$camp_sel].fee + '</td>' +
               '<td style="text-align:center;'+$new_aval_col+';">' + $camps[$camp_sel].taken + '</td>' +
               '<td style="text-align:center;'+$new_aval_col+';">' + $camps[$camp_sel].seatsmax + '</td>' +
               '</tr>' + "`n"
}

$content_output = $template[0] + $newstr + $template[1]

Write-Output "Writing JSON archive file: .\buonacaccia_spec.json"
# $ErrorActionPreference = "SilentlyContinue"
$camps_json = $camps | ConvertTo-Json -Depth 4 -ErrorAction SilentlyContinue
$camps_json | Out-File -FilePath .\buonacaccia_spec.json -Encoding UTF8

Write-Output "Writing HTML result file: .\buonacaccia_spec_mod.html"
$content_output | Out-File -FilePath .\buonacaccia_spec_mod.html -Encoding UTF8
Remove-Item -Path ".\buonacaccia_spec.html"

if ($trigger_popup_newstatus -and $old_json_available -and $camps_newstatus_str -ne ""){
    $Result = [System.Windows.Forms.MessageBox]::Show(
        "Nuova disponibilità di campetti di specialità per:`n" + $camps_newstatus_str,
        "Campetti Specialità - Nuova disponibilità",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}
