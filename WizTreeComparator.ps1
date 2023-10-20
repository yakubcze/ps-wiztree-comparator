# --------------------------------
# Jakub Mlynek
#
# WizTree + WizTreeCompare skript
# --------------------------------

$wtPath = "C:\Program Files\WizTree\WizTree64.exe" #cesta k WizTree programu
$wtcPath = "C:\wiztree\WizTreeCompare-v0.3.0-RC.2\WizTreeCompare.exe" #cesta k WizTreeCompare programu
$scanPath = "\\f-server-v\E" #cesta ke skenovanemu adresari

$csvFilePath_wt = "C:\wiztree\export\" #cesta kam se budou ukladat exportovane csv soubory pro WizTree
$csvFileName_wt = "WizTree_$(Get-Date -Format "dd_MM_yyyy_HHmm").csv" #format vystupniho csv souboru pro WizTree

$csvFilePath_wtc = "C:\wiztree\compare\" #cesta kam se budou ukladat exportovoane csv soubor pro WizTreeCompare
#csvFileName_wtc = nize v kodu

try {
    #----------------------------------------Skenovani adresare----------------------------------------

    $process = Start-Process $wtPath -ArgumentList "$scanPath /export=$csvFilePath_wt$csvFileName_wt /admin 1 /sortby=2 /exportfolders=1" -PassThru #spusteni WizTree = spusteni skenu adresare

    <#
    nekonecny "progress bar" - jenom vypisuje ze se neco deje, nepise procenta ani zbyvajici cas
    aby toto fungovalo je nutno pouzit -PassThru pri spousteni procesu

    pokud by toto nebylo provedeno, skript by treba na hodinu zamrzl a nic by se nedelo
    #>
    for($i = 0; $i -le 100; $i = ($i + 1) % 100)
    {
        Write-Progress -Activity "WizTree" -PercentComplete $i -Status "Probiha sken adresare..."
        Start-Sleep -Milliseconds 100
        if ($process.HasExited) {
            <#
            #pokud proces skonci, znovu se zkontroluje jestli program WizTree nebezi 
            (protoze wiztree spusti jednu instanci, kterou po chvili zavre a nasledne otevre dalsi s jinym PID = proto je nutne to takhle debilne dvojite kontrolovat)
            #>
            $isRunning = (Get-Process | Where-Object { $_.Name -eq "WizTree64" }).Count -gt 0
            if ($isRunning){ 
                $process = Get-Process "WizTree64"
                continue #pokud program bezi, prepise se objekt $process a vyskoci se na zacatek for smycky
            }

            else {
                Write-Progress -Activity "Wiztree" -Completed
                break #pokud program opravdu skoncil, ukonci se progress bar a pokracuje se dal v programu
            }
        }
    }

    #----------------------------------------Porovnavani souboru----------------------------------------
    <#
    projde soubory ve specifikovanem adresari, vezme vsechny .csv soubory, seradi podle data vytvoreni, selectne dva nejnovejsi soubory a ulozi je do pole na indexy:
    nejnovejsi: [0]
    2. nejnovejsi: [1]
    #>
    $items = (Get-ChildItem -Path $csvFilePath_wt -Filter *.csv | Sort-Object -Descending -Property CreationTime | Select-Object -First 2)
    
    #zformatuje nazvy souboru do pozadovaneho formatu pro WizTreeCompare - "<starsi.csv>" "<novejsi.csv>" "<vystupni.csv>"
    $csvFileName_wtc = "WizTreeCompare$(($items[1].Name | Out-String) -replace 'WizTree|\.csv|\r?\n\z', '')_vs$(($items[0].Name | Out-String) -replace 'WizTree|\.csv|\r?\n\z', '').csv"
    #'WizTree|\.csv|\r?\n\z', '' => nahrazeni WizTree NEBO .csv NEBO newline charakteru za prazdny retezec pomoci regexu (=formatovani nazvu vystupni souboru)

    $csvs = $csvFilePath_wt + $items[1].Name + " " + $csvFilePath_wt + $items[0].Name + " " + $csvFilePath_wtc + $csvFileName_wtc
    #        \---------starsi.csv----------/         \--------novejsi.csv-----------/         \-----------vystupni.csv----------/         
    Start-Process $wtcPath -ArgumentList "-D $csvs" #spusti WizTreeCompare s predanymi argumenty
}

finally {
    <#
    na konci behu skriptu:

    zkontroluje se jestli bezi WizTree proces
    pokud ano -> ukonci se
    pokud ne -> jenom vypise hlasku

    toto umoznuje kdydokliv ukoncit skript pomoci Ctrl+C, cimz se zaroven ukonci WizTree
    bez tohoto by WizTree proces zustal bezet v pozadi
    #>
    $isRunning = (Get-Process | Where-Object { $_.Name -eq "WizTree64" }).Count -gt 0
    if($isRunning) {
        $processToKill = Get-Process | Where-Object { $_.Name -eq "WizTree64" }
        taskkill /PID $processToKill.Id /F
    }
    else {
        Write-Host "Beh skriptu probehl uspesne"
    }
}


<#
dopsat nekde nejaky readme soubor:
- wiztree musi byt v anglictine
- zdokumentovat vsechny cesty (path) co tu jsou napsane
- co program dela (proskenuje adresar, porovna ho s predchozi verzi, ...)
- pokud je ve slozce exports jenom jeden soubor, program spadne ale nic nenapise
- ctrl+c pro zruseni
#>