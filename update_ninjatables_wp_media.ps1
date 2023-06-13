#this script takes a ninjatable export, extracts the media list from wordpress,
# finds all href and tries to find it in the media list, corrects it, and spits
# out an importable csv for ninjatables

#TODO some files show as missing when there is a match, unsure why

#========== Options ==========
$wp_root = "https://WEBSITE"
$wp_user = "USERNAME"
$wp_pass = ConvertTo-SecureString -String "PASSWORD" -AsPlainText -Force #token/api key
$csv_file = "input.csv" #ninjatable export
$out_file = "output.csv" #to import to ninjatable VERIFY BEFORE REPLACING!
$dupe_file = "dupe.csv" #list of ambiguous file matches
$missing_file = "missing.csv" #list of missing files
$mediaList_file = "MediaList.json" #exported list of wordpress list
$use_mediaListCache = $true #use previously cached list to not query again
$extra_cleaning= $true #TODO replace sided quotes with regular quotes, for now use notepad++
#========== END Options ==========

$wp_credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $wp_user, $wp_pass
$wp_api_test = "/wp-json"
$wp_api_media = "/wp-json/wp/v2/media"
$wp_perPage = 100 #number of media files to get at a time, 1-100
#=====Get Media List=====
$wp_mediaList = @()
$response_data = $null

if ($use_mediaListCache) {
    if (Test-Path -Path $mediaList_file) {
        "Cached media list found"
    }
    else {
        "Cached media list does not exist!"
        $use_mediaListCache = $false
    }
}
if (!$use_mediaListCache) {
    $response = Invoke-WebRequest -Uri "$wp_root$wp_api_test" -Authentication Basic -Credential $wp_credential
    if ($response.StatusCode -eq 200) {
        Write-Host "Authenticated!"
        $moreMedia = $true
        for ($currentPage = 1; $moreMedia; $currentPage++) {
        #for ($currentPage = 1; $currentPage -eq 1; $currentPage++) { #DEBUG
            Write-Host -NoNewline ("Page:{0}..." -f $currentPage)
            $response = Invoke-WebRequest -Uri "$($wp_root)$($wp_api_media)?page=$($currentPage)&per_page=$($wp_perPage)" -Authentication Basic -Credential $wp_credential
            #FIXME test if 404?, what happens if you go to a page beyond how many you have
            $response_data = $response.Content | ConvertFrom-Json
            if ($response_data.Count -lt $wp_perPage) {
                $moreMedia = $false
            }
            $wp_mediaList += $response_data
            Write-Host ("Files so far:{0}" -f ($wp_mediaList.Count))
        }
    }
    else {
        "code:".$response.StatusCode
    }
}
else {
    "Reading cache media list"
    $wp_mediaList = Get-Content -Path $mediaList_file | ConvertFrom-Json
}
"number of items:{0}" -f $wp_mediaList.Count
$wp_mediaList | ConvertTo-Json -depth 100 | Out-File -FilePath $mediaList_file -Encoding unicode

#=====Get CSV and fix=====
$numChanged = 0
$numUnchanged = 0
$numDupes = 0
$numMissing = 0
$csv_data = Import-Csv -Path $csv_file
$splitstr = 'href="'
$dupeArr = @()
$missingArr = @()
#$csv_data | Format-Table

foreach ($row in $csv_data)
{
    $row.PSObject.Properties | ForEach-Object {
        #===================== PER CELL=======================
        Write-Host -NoNewline "."
        #$_ | Out-File -Path "test.txt" -Append
        #split LHS of tokens
        $splitarr = $_.Value -split $splitstr
        $splitArr2 = @()
        
        $splitArr2 += $splitarr[0]
        #split RHS of tokens
        for ($i = 1; $i -lt $splitarr.Count; $i++) {
            $loc = $splitarr[$i].IndexOf('"');
            $splitArr2 += $splitarr[$i].SubString(0,$loc).Trim()
            $splitArr2 += $splitarr[$i].SubString($loc)
        }
        Remove-Variable -name splitarr #reclaim memory
        #put back splitstr
        for ($i = 0; $i -lt ($splitarr2.Count - 1); $i += 2) {
            $splitArr2[$i] = $splitArr2[$i] + $splitstr
        }
        #compare / update path
        for ($i = 1; $i -lt ($splitarr2.Count - 1); $i += 2) {
            #Leaf gets extension, LeafBase does not
            $leaf = Split-Path $splitArr2[$i] -LeafBase
            #"Leaf:{0}" -f $leaf  | Out-File -Path "test.txt" -Append #DEBUG
            #TODO detect the -### at the end that shows up with duplicate file
            $url = ($wp_medialist | Where-Object { $_.slug -like $leaf.Substring(0,$leaf.Length)+"*" })
            
            #$url  | Out-File -Path "test.txt" -Append #DEBUG
            if ($url.Length -eq 1) {
                #"only one" #DEBUG
                if ($url[0].source_url -ne $splitArr2[$i]) {
                    $splitArr2[$i] = $url[0].source_url
                    $numChanged++;
                }
                else {
                    $numUnchanged++;
                }
                
            }
            elseif ($url.length -gt 1) {
                "Dupe potential"
                $dupeArr += $url.source_url
                $numDupes++;
            }
            else {
                "missing"
                $missingArr += $_.Value
                $missingArr += $leaf
                $numMissing++;
            }
        }
        #reconstruct
        $_.Value = ""
        foreach ($line in $splitarr2) {
            $_.Value += $line
        }
        #===================== END PER CELL===================
    }#foreach / percell
}#foreach / row
""
$csv_data | Export-Csv -Path $out_file
$dupeArr | Out-File -FilePath $dupe_file
$missingArr | Out-File -FilePath $missing_file
"numChanged:   {0}" -f $numChanged
"numUnchanged: {0}" -f $numUnchanged
"numDupes:     {0}" -f $numDupes
"numMissing:   {0}" -f $numMissing