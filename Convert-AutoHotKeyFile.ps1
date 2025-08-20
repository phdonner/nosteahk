# Convert-AutoHotKeyFile.pst - Ver. 1.1 19.08.2025

# Copyright (c) 2025, Ph. Donner

# Tässä versiossa oletamme edelleen että AutoHotKey skripti on sijoitettu työpöydälle

Function Get-BillingPeriodDlg
<#
.SYNOPSIS
Short description

.DESCRIPTION
Osuuskunta huolehtii myyntilaskutuksesta manuaalisesti OP-pankin yrityksen verkkopankin WWW-käyttöliittymän kautta.
Pankin matkapuhelinsovelluksessa voi ilmeisesti kopioida laskun tiedot, mutta se ei ole mahdollista tietokoneella.

Siksi käytämme apuna AutoHotKey-makroja, jotka täyttävät laskun tiedot automaattisesti.
F8-näppäimellä käynnistettävä makro täyttää laskun tiedot, kuten laskutuskauden, eräpäivän ja maksuerän.
F9-näppäin luo myös tietueen paperilaskusta aiheutuvan lisäkulun kattamiseksi.

HUOM: 
OP-palvelun toiminta on epätasaista, joten makro ei aina toimi moitteettomasti.
Jos täyttö epäonnistuu, on helpointa käynnistää makro uudelleen.

.PARAMETER Year
Parameter description

.PARAMETER Set
Parameter description

.PARAMETER Caption
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
    {
    param (
        [Parameter(Mandatory=$false)]
        [int]$Year=0, 

        [Parameter(Mandatory=$false)]
        [int]$Period=0, 

        [Parameter(Mandatory=$false)]
        [bool]$Set=$false,

        [Parameter(Mandatory=$false)]
        [string]$Caption='Valitse vuosi ja laskutuskausi'
        ) 

    $CurrentDate = Get-Date

    If (($Year -eq 0) -or ($Set -eq $true))
        {
        $Year= $CurrentDate.Year
        $Set = [Math]::Truncate(($CurrentDate.Month + 1)/2)
        }

    If (($Year -lt ($CurrentDate.Year - 5)) -or ($Year -gt ($CurrentDate.Year + 3)))
        {
        # Failed, reset to current year
        $Year= $CurrentDate.Year
        }

    If (($Period -lt 1) -or ($Period -gt 6))
        {
        # Failed, reset to set derived from current month
        $Period = [Math]::Truncate(($CurrentDate.Month + 1)/2)
        }

    ## Main Window and Caption
    Add-Type -assembly System.Windows.Forms
    $YearSetDlg = New-Object System.Windows.Forms.Form

    $YearSetDlg.Width           = 350
    $YearSetDlg.Height          = 122
    $YearSetDlg.AutoSize        = $false
    $YearSetDlg.Text            = $Caption
    $YearSetDlg.MinimizeBox     = $true
    $YearSetDlg.MaximizeBox     = $false
    $YearSetDlg.SizeGripStyle   = 'Hide'
    $YearSetDlg.FormBorderStyle = 'Fixed3D'
    $YearSetDlg.StartPosition   = 'CenterScreen'
    $YearSetDlg.Topmost         = $true

    #### Create 'Year - Period' Label
    $YearPeriodLabel            = New-Object System.Windows.Forms.Label
    $YearPeriodLabel.Text       = 'Vuosi - Kausi:'                           
    $YearPeriodLabel.Location   = New-Object System.Drawing.Point(13,20)
    $YearPeriodLabel.AutoSize   = $true
    $YearSetDlg.Controls.Add($YearPeriodLabel)

    #### Create Year ComboBox
    $ComboBoxYear = New-Object System.Windows.Forms.ComboBox
    $ComboBoxYear.Width = 50

    #### Add Year Combobox items
    $ItemYear = $Year - 4
    
    While ($ItemYear -lt ($Year + 4))
        {
        $ComboBoxYear.Items.Add($ItemYear.ToString()) | Out-Null

        $ItemYear++
        }

    # Select default Year item
    $ComboBoxYear.SelectedIndex = 4
            
    $ComboBoxYear.Location  = New-Object System.Drawing.Point(93,18)
    $YearSetDlg.Controls.Add($ComboBoxYear)

    #### Hyphen Label
    $LabelHyphen            = New-Object System.Windows.Forms.Label
    $LabelHyphen.Text       = '-'
    $LabelHyphen.Location   = New-Object System.Drawing.Point(144,20)
    $LabelHyphen.AutoSize   = $true
    $YearSetDlg.Controls.Add($LabelHyphen)

    #### Create Batch ComboBox
    $ComboBoxSet            = New-Object System.Windows.Forms.ComboBox
    $ComboBoxSet.Width      = 40

    #### Add Period Combobox items
    $n = 1

    While ($n -le 6)
        {
        $strPeriod = '{0:d}' -f $n
        $null = $ComboBoxSet.Items.Add($strPeriod)
        $n++
        }

    # Select default Period item
    $null = $ComboBoxSet.SelectedIndex = $Period - 1

    $ComboBoxSet.Location = New-Object System.Drawing.Point(154,18)
    $YearSetDlg.Controls.Add($ComboBoxSet)

    #### Create OK Button
    $ButtonOk           = New-Object System.Windows.Forms.Button
    $ButtonOk.Location  = New-Object System.Drawing.Size(153,50)
    $ButtonOk.Size      = New-Object System.Drawing.Size(60,23)
    $ButtonOk.Text      = 'OK'
    $ButtonOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $YearSetDlg.Controls.Add($ButtonOk)

    #### Search for OK Button Function
    $ButtonOk.Add_Click(
            {
            $Year   = $ComboBoxYear.text
            $Period = $ComboBoxSet.text
            $YearSetDlg.Close()
            }
        )

    #### Cancel Button
    $ButtonCancel = New-Object System.Windows.Forms.Button
    $ButtonCancel.Location = New-Object System.Drawing.Size(223,50)
    $ButtonCancel.Size = New-Object System.Drawing.Size(60,23)
    $ButtonCancel.Text = 'Peruuta'
    $ButtonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel # 'Cancel'

    $YearSetDlg.Controls.Add($ButtonCancel)

    #### Set Cancel Button Function
    $ButtonCancel.Add_Click(
            {
            $Year = $null
            $Set = $null
            $YearSetDlg.Close()
            }
        )

    #### Set keyboard feedback
    $YearSetDlg.AcceptButton = $ButtonOk
    $YearSetDlg.CancelButton = $ButtonCancel

    $BillingPeriod = $null

    # Show the select billing dialog topmost and with focus
    $YearSetDlg.Add_Shown({$YearSetDlg.Activate()})

    If ($YearSetDlg.ShowDialog($this) -eq [System.Windows.Forms.DialogResult]::OK)
        {
        $BillingPeriod = "$($ComboBoxYear.Text)-$($ComboBoxSet.Text)"
        }

    #### Clean up
    $YearSetDlg.Dispose()

    Return $BillingPeriod
    }

Function Get-BillingPeriodDescription
    {
    <#
    .SYNOPSIS
    Prepares a string which describes the start date and and the end date from the descriptive string of the billing period
    
    .DESCRIPTION
    Long description
    
    .PARAMETER Period
    The period string is submitted in the form <year>-<month> (e.g. 2025-02 and 2026-12)
    
    .EXAMPLE
    Get-BillingPeriodDescription -Period '2025-02'
    
    .NOTES
    Wish list: Create alternative parameterset which accepts period as an object (year, month)
    #>

    Param
        (
        [Parameter(Mandatory=$true)]
        [string]$Period
        )

    # Convert period string to number parameters

    $Year   = $Period.Substring(0, 4)
    $Number = $Period.Substring(5, 1) -as [int]
    
    # Prepare month name descriptions from period info

    $fi = New-Object system.globalization.cultureinfo('fi-FI')

    $firstmonth = $fi.DateTimeFormat.GetMonthName((2*$Number)-1)
    $secondmonth = $fi.DateTimeFormat.GetMonthName(2*$Number)

    $BillingPeriodDescription = "$($firstmonth)-$($secondmonth) $Year"

    Return $BillingPeriodDescription
    }

$Desktop = [Environment]::GetFolderPath('Desktop')
$Script = "$Desktop\nostelasku.ahk"

If (-Not (Test-Path $Script -PathType Leaf))
    {
    $Message = 'Työpöydällä ei ole AutoHotKey-tiedostoa: nostelasku.ahk'
    Write-Warning $Message
    Exit
    }

$Period = Get-BillingPeriodDlg

If ($Null -ne $Period)
    {
    $Year   = $Period.Substring(0, 4)
    $Batch  = $Period.Substring(5)
    $n      = 2 * $Batch

#   $AlkuPvm = [DateTime]"01.$($n).$Year"
    $AlkuPvm    = Get-Date -Year $Year -Month $n -Day 1
    $LoppuPvm   = $AlkuPvm.AddMonths(1)

    $AlkuPvm    = $LoppuPvm.AddMonths(-2)
    $LoppuPvm   = $LoppuPvm.AddDays(-1)

    $AlkuPvm    = $AlkuPvm.ToString('dd.MM.yyyy')
    $LoppuPvm   = $LoppuPvm.ToString('dd.MM.yyyy')

    [array]$arrayScript = @(Get-Content -Path $Script -Encoding utf8)

#    Write-Debug $arrayScript[0]
#    Write-Debug $arrayScript[1]
#    Write-Debug $arrayScript[2]

    $Count = $arrayScript.Count
    $n = 0

    While ($n -lt $Count)
        {
        $Line = $arrayScript[$n]
        # Write-Host $Line

        Switch ($Line)
            {
            '; AlkuPvm' {$arrayScript[$n+1] = "Send `"$AlkuPvm`""}
            '; LoppuPvm'{$arrayScript[$n+1] = "Send `"$LoppuPvm`""}
            '; MaksuErä' {$arrayScript[$n+1] = "Send `"Kuukausimaksuerä $($Year)-$Batch`""}
            }

        $n++
        }

    Set-Content -Path $Script -Value $arrayScript -Encoding utf8
    $command = "`"C:\Program Files\AutoHotkey\v2\autohotkey.exe`""
    $argument = $Script

    Write-Debug $argument

    Start-Process -FilePath $command -ArgumentList $argument -NoNewWindow

    # Write-Host 'END'
    }     
    