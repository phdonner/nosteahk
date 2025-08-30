g; AutoHotkey script, v. 2.4

; - The NOSTE invoice macro can be autoupdated with this PowerShell script: 
;   Convert-AutoHotKeyFile.ps1 in PSExperiment\AutoHotkey folder
; - The nostelasku.ahk script can be nullified by executing the null_f8_f9.ahk script.

; Old, perhaps outdated commands (v.1.0?)
; #NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
; SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
; SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; nostelasku.ahk ver. 2.4 20.6.2025 Finetune because OP changed the layout of their form
; nostelasku.ahk ver. 2.3 Replace all v. 1 Send strings with v. 2 formatted ones
; nostelasku.ahk ver. 2.2 Disable #NoEnv command
; nostelasku.ahk ver. 2.1, 12:00 19.10.2024

; Manual: https://www.autohotkey.com/docs/v2/

; OP verkkopalvelun toiminta on epätasaista, erityisesti silloin kun palveluun
; kohdistuu ruuhkaa. Tätä varten makroja joutuu joskus toistamaan.
; Jos näyttää siltä että hidaste on eri päivinä pysyvää, täytyy makron viiveitä hienosäätää.
; Samanlainen tarve syntyy kun pankin lomakkeiden ulkoasua muutetaan.

; NOSTE OP-laskujen täyttöä palvelevan autohotkey-skriptin päivitysohje:

; 1. Tee varmuuskopio
; 2. Jaa toimitettava skripti osiin merkinnöillä:
;         <alkuosa>
;         }
;
;         F12:: 
;         {
;         <loppuosa>
; 3. Siirrä jako loppua kohti sitä mukaan kun skripti valmistuu
; 4. Poista lopulta skriptiä jakavat merkinnät
;
; Valmista tarvittaessa numeroituja välikopioita.
; Varmista että kohdistusliput säilyvät paikoillaan

; Wish list:
; - Vaihda kohdistusliput AHK label -merkinnöillä

F9::
{
; ^+p::

; Tällä makrolla voit kirjoittaa Finvoice-laskun, asiakkaille jotka  
; tarvitsevat osuuspankin toimittaman paperilaskun.

; Kirjoittaa Finvoice-laskun 'Lisää töitä tai tuotteita laskulle' 
; -dialogin KKM ja KKMPA kuukausimaksua koskevat arvot 

; Aloita 'Lisää töitä tai tuotteita laskulle' -dialogista
; Valitse laskutettava asiakas ja aloita makro 'Lisää työ tai tuote' -näkymästä

; Tällä makrolla voit kirjoittaa Finvoice-laskun, asiakkaille joille 
; riittää pelkkä E-lasku (ilman postissa lähetettyä paperiversiota).
; Makro valitsee E-laskulomakkeen ja lisää kuukausimaksua koskevat arvot.

; Aloita makron suorittaminen Yrityksen OP verkkopankin laskutuksesta.
; Hyväksy osuuskunta laskuttajaksi ja valitse asiakas.
; Siirry sen jälkeen 'Lisää töitä tai tuotteita laskulle' -välilehdelle.

; Makro käynnistyy näppäinyhdistelmällä: Ctrl+Shift+p

; Varoitus: Älä tee mitään tuolla 'Lisää töitä...' välilehdellä! 
; Focus ei saa olla valittuna johonkin dialogin osioon, esim. painikkeeseen. 
; Jos vahingossa siirryt johonkin välilehden kontrolliin 
; voit poistaa valinnan osoittamalla välilehden tyhjää pintaa.
; Osoita vasemmalla hiiripainikkeella dialogin asiakasaluetta.

; Toinen varoitus: Irrota otteesi painikkeista heti sen jälkeen
; kun makro on käynnistynyt. Muuten makroa suorittava ohjelma
; huomioi myös painettuja näppäimiä.

; Aloita makron suorittaminen Yrityksen OP verkkopankin laskutuksesta.
; Hyväksy osuuskunta laskuttajaksi ja valitse asiakas.
; Siirry sen jälkeen 'Lisää töitä tai tuotteita laskulle' -välilehdelle.

; Varoitus 1: Älä tee mitään tuolla välilehdellä! 
; Focus ei saa olla valittuna johonkin dialogin osioon, esim. painikkeeseen. 
; Jos vahingossa siirryt johonkin välilehden kontrolliin 
; voit poistaa valinnan osoittamalla välilehden tyhjää pintaa.
; Osoita vasemmalla hiiripainikkeella dialogin asiakasaluetta.

; Varoitus 2: Irrota otteesi painikkeista heti sen jälkeen
; kun makro on käynnistynyt

; Varoitus 3: Aloita aina niin että www-sivu on maksimoituna

Sleep 900
Send "{Tab down}{Tab up}"
Send "{Space down}{Space up}"
Sleep 900

; Nyt 'Lisää töitä tai tuotteita laskulle' -dialogi pitäisi olla näkyvillä
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Right}"
Sleep 400
Send "{Tab down}{Tab up}"
Sleep 400
Send "KKMEL"
Sleep 400
Send "{Enter}"

Sleep 1100
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Sleep 400
Send "{Enter}"
Sleep 400
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "2"
Sleep 400
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Space down}{Space up}"
Send "{Tab down}{Tab up}"
Sleep 500

; AlkuPvm
Send "01.05.2025"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"

Sleep 500

; LoppuPvm
Send "30.06.2025"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"

Sleep 500

; MaksuErä
Send "Kuukausimaksuerä 2025-3"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Sleep 700
Send "{Space down}{Space up}"
Sleep 700

Send "{Space down}{Space up}"
Sleep 700

; 'Lisää töitä tai tuotteita laskulle' -dialogin pitäisi olla taas näkyvillä
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Right}"
Sleep 400
Send "{Tab down}{Tab up}"
Sleep 500
Send "KKMPL"
Sleep 500

Send "{Enter}"
Sleep 500
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Enter}"
Sleep 500
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Space down}{Space up}"
Sleep 500
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
}

F8::
{
; ::^+e
; Tällä makrolla voit kirjoittaa Finvoice-laskun, asiakkaille joille 
; riittää pelkkä E-lasku.
; Makro valitsee E-lasku lomakkeen ja lisää kuukausimaksua koskevat arvot.

; Aloita makron suorittaminen Yrityksen OP verkkopankin laskutuksesta.
; Hyväksy osuuskunta laskuttajaksi ja valitse asiakas.
; Siirry sen jälkeen 'Lisää töitä tai tuotteita laskulle' -välilehdelle.

; Varoitus: Älä tee mitään tuolla välilehdellä! 
; Focus ei saa olla valittuna johonkin dialogin osioon, esim. painikkeeseen. 
; Jos vahingossa siirryt johonkin välilehden kontrolliin 
; voit poistaa valinnan osoittamalla välilehden tyhjää pintaa.
; Osoita vasemmalla hiiripainikkeella dialogin asiakasaluetta.

; Toinen varoitus: Irrota otteesi painikkeista heti sen jälkeen
; kun makro on käynnistynyt

Sleep 900
Send "{Tab down}{Tab up}"
Send "{Space down}{Space up}"
Sleep 900

; Nyt 'Lisää töitä tai tuotteita laskulle' -dialogi pitäisi olla näkyvillä
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Right}"
Sleep 400
Send "{Tab down}{Tab up}"
Sleep 400
Send "KKMEL"
Sleep 400
Send "{Enter}"

Sleep 1100
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Sleep 400
Send "{Enter}"
Sleep 400
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "2"
Sleep 400
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Space down}{Space up}"
Send "{Tab down}{Tab up}"
Sleep 500

; AlkuPvm
Send "01.05.2025"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"

Sleep 500

; LoppuPvm
Send "30.06.2025"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"

Sleep 500

; MaksuErä
Send "Kuukausimaksuerä 2025-3"
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"
Send "{Space down}{Space up}"
Sleep 3000
Send "{Tab down}{Tab up}"
Send "{Tab down}{Tab up}"

Return
}
