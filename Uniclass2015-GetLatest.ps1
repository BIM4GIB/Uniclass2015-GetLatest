#########################################################################################
#   ________  _______   ___      ___ ___  _________  _______   _______   ________       #  
#  |\   __  \|\  ___ \ |\  \    /  /|\  \|\___   ___\\  ___ \ |\  ___ \ |\   __  \      #  
#  \ \  \|\  \ \   __/|\ \  \  /  / | \  \|___ \  \_\ \   __/|\ \   __/|\ \  \|\  \     #  
#   \ \   _  _\ \  \_|/_\ \  \/  / / \ \  \   \ \  \ \ \  \_|/_\ \  \_|/_\ \   _  _\    #  
#    \ \  \\  \\ \  \_|\ \ \    / /   \ \  \   \ \  \ \ \  \_|\ \ \  \_|\ \ \  \\  \|   #  
#     \ \__\\ _\\ \_______\ \__/ /     \ \__\   \ \__\ \ \_______\ \_______\ \__\\ _\   #  
#      \|__|\|__|\|_______|\|__|/       \|__|    \|__|  \|_______|\|_______|\|__|\|__|  # 
#########################################################################################
# 
#######################
# Uniclass2015 Tables #
#ClassificationManager#
# GetLatest&Merge     #
# v0.11               #
# 13/10/21            #
# by RPG @BIM4GIB     #
# reviteer@hotmail.com#
#######################


# Get Start Time
$startDTM = (Get-Date)

#Add Windows forms, d'uh...
Add-Type -AssemblyName System.Windows.Forms

#Get current location and date
$currentPath = [environment]::CurrentDirectory #[string](Get-Location)
$DateStamp = Get-Date -Format yyyy-MM

#Set TLS1.2 or the web-requests will fail
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Assign NBS uri
$uriNBS = "https://www.thenbs.com"
$uriUniclass = "https://www.thenbs.com/our-tools/uniclass-2015"
$invokeURI = Invoke-WebRequest -uri $uriUniclass -UseBasicParsing

#Assign paths                     #Below the HTML paths to xlsx's, maybe will use in future...
$Co = $currentPath + "\Co.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[2]/td[1]/a                        
$En = $currentPath + "\En.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[3]/td[1]/a                    
$Ac = $currentPath + "\Ac.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[4]/td[1]/a 
$SL = $currentPath + "\SL.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[5]/td[1]/a
$EF = $currentPath + "\EF.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[6]/td[1]/a
$Ss = $currentPath + "\Ss.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[7]/td[1]/a
$Pr = $currentPath + "\Pr.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[8]/td[1]/a
$TE = $currentPath + "\TE.xlsx"   #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[9]/td[1]/a
$PM = $currentPath + "\PM.xlsx"  #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[10]/td[1]/a
$Zz = $currentPath + "\Zz.xlsx"  #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[11]/td[1]/a
$FI = $currentPath + "\FI.xlsx"  #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[12]/td[1]/a
$Ro = $currentPath + "\Ro.xlsx"  #//*[@id="modal-download-tables"]/article/div/table/tbody/tr[13]/td[1]/a
$CMDBpath = $currentPath + '\ClassificationManagerDatabase-Uniclass2015.xlsx'
$CoPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "Co.pdf"                           
$EnPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "En.pdf"                       
$AcPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "Ac.pdf"
$SLPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "SL.pdf"
$EFPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "EF.pdf"
$SsPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "Ss.pdf"
$PrPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "Pr.pdf"
$TEPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "TE.pdf"
$PMPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "PM.pdf"
$ZzPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "Zz.pdf"
$FIPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "FI.pdf"
$RoPDF = $currentPath + "\UpdatesNotesPDFs\" + $DateStamp + "Ro.pdf"

#Check if Uniclass2015-AllTables.xlsx exists, and if so, deletes it
if (Test-Path ($currentPath +'\Uniclass2015-AllTables.xlsx'))
{Remove-Item ($currentPath + '\Uniclass2015-AllTables.xlsx')}

#Get latest tables links from NBS and temporarily save to current folder
$tableCo = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*co*xlsx*"}).href)
Invoke-WebRequest -Uri $tableCo -OutFile $Co
$tableEn = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*en*xlsx*"}).href)
Invoke-WebRequest -Uri $tableEn -OutFile $En
$tableAc = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ac*xlsx*"}).href)
Invoke-WebRequest -Uri $tableAc -OutFile $Ac
$tableSL = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*sl*xlsx*"}).href)
Invoke-WebRequest -Uri $tableSL -OutFile $SL
$tableEF = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ef*xlsx*"}).href)
Invoke-WebRequest -Uri $tableEF -OutFile $EF
$tableSs = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ss*xlsx*"}).href)
Invoke-WebRequest -Uri $tableSs -OutFile $Ss
$tablePr = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*pr*xlsx*"}).href)
Invoke-WebRequest -Uri $tablePr -OutFile $Pr
$tableTE = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*te*xlsx*"}).href)
Invoke-WebRequest -Uri $tableTE -OutFile $TE
$tablePM = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*pm*xlsx*"}).href)
Invoke-WebRequest -Uri $tablePM -OutFile $PM
$tableZz = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*zz*xlsx*"}).href)
Invoke-WebRequest -Uri $tableZz -OutFile $Zz
$tableFI = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*fi*xlsx*"}).href)
Invoke-WebRequest -Uri $tableFI -OutFile $FI
$tableRo = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ro*xlsx*"}).href)
Invoke-WebRequest -Uri $tableRo -OutFile $Ro

#PDF updates...
#Temporarily removed question askign user y/n to downalod pdfs.
#$getPDFs = ((Read-Host -Prompt "Download Status and Revision Information PDFs? (y/n)").ToLower()) #Check if PDF need downloading
#if ($getPDFs -eq 'y'){if(!(Test-Path -Path ($currentPath + "\UpdatesNotesPDFs\"))) { }              #Check if pdf destination path exists
	Test-Path -Path ($currentPath + "\UpdatesNotesPDFs\")
    New-Item -ItemType directory -Path ($currentPath + "\UpdatesNotesPDFs\") -Force 

    $tableCo = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*co*pdf*"}).href)
    Invoke-WebRequest -Uri $tableCo -OutFile $CoPDF
    $tableEn = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*en*pdf*"}).href)
    Invoke-WebRequest -Uri $tableEn -OutFile $EnPDF
    $tableAc = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ac*pdf*"}).href)
    Invoke-WebRequest -Uri $tableAc -OutFile $AcPDF
    $tableSL = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*sl*pdf*"}).href)
    Invoke-WebRequest -Uri $tableSL -OutFile $SLPDF
    $tableEF = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ef*pdf*"}).href)
    Invoke-WebRequest -Uri $tableEF -OutFile $EFPDF
    $tableSs = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ss*pdf*"}).href)
    Invoke-WebRequest -Uri $tableSs -OutFile $SsPDF
    $tablePr = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*pr*pdf*"}).href)
    Invoke-WebRequest -Uri $tablePr -OutFile $PrPDF
    $tableTE = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*te*pdf*"}).href)
    Invoke-WebRequest -Uri $tableTE -OutFile $TEPDF
    $tablePM = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*pm*pdf*"}).href)
    Invoke-WebRequest -Uri $tablePM -OutFile $PMPDF
    #disabling Zz until update exists, to avoid invalid PDF
    $tableZz = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*zz*pdf*"}).href)
    #Invoke-WebRequest -Uri $tableZz -OutFile $ZzPDF
    $tableFI = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*fi*pdf*"}).href)
    Invoke-WebRequest -Uri $tableFI -OutFile $FIPDF
    $tableRo = ($uriNBS + ($invokeURI.Links | ? {$_.href -like "*uniclass2015*ro*pdf*"}).href)
    Invoke-WebRequest -Uri $tableRo -OutFile $RoPDF
#} End of if statement above, temporarily removed because reasons.
#Begin merging tables into one file

#Initialise excel
$xl = New-Object -c excel.application
$xl.Visible = $false #was true tho!
$xl.displayAlerts = $false

# Open destination workbooks                      
$CMDB = $xl.workbooks.open($CMDBpath)#Open Classification Manager Database Uniclass 2015
$wb12 = $xl.workbooks.open($Ro) #Open table Ro so other tables are copied into it below

#Start copying tables >>https://theolddogscriptingblog.wordpress.com/2010/06/01/powershell-excel-cookbook-ver-2/
#FI
$wb11 = $xl.workbooks.open($FI, $null, $false)          
$sh1_wb11 = $wb12.sheets.item(1)    
$sheetToCopy = $wb11.sheets.item("FI")   
$sheetToCopy.copy($sh1_wb11) 
$wb11.close($false)               
#Zz
$wb10 = $xl.workbooks.open($Zz, $null, $false)          
$sh1_wb10 = $wb12.sheets.item(1)    
$sheetToCopy = $wb10.sheets.item("Zz")   
$sheetToCopy.copy($sh1_wb10) 
$wb10.close($false)               
#PM
$wb9 = $xl.workbooks.open($PM, $null, $false)          
$sh1_wb9 = $wb12.sheets.item(1)    
$sheetToCopy = $wb9.sheets.item("PM")   
$sheetToCopy.copy($sh1_wb9) 
    #update Classification Manager>make funcation
$shEF = $CMDB.sheets.item("Uniclass Table PM")#destination
$used = $sheetToCopy.usedRange
$lastCell = $used.SpecialCells(11)
$lastRow = $lastCell.row
$lastRow2 = $lastRow + 5
[void]$sheetToCopy.Range("A4:A$lastRow").Copy()
[void]$shEF.Range("A9").PasteSpecial(-4163) 
[void]$sheetToCopy.Range("F4:F$lastRow").Copy()
[void]$shEF.Range("B9").PasteSpecial(-4163)
[void]$sheetToCopy.Range("A1").Copy()
[void]$shEF.Range("E3").PasteSpecial(-4163)
[void]$shEF.Range("F9:F$lastRow2").Copy()
[void]$shEF.Range("C9").PasteSpecial(-4163)
$wb9.close($false) 
#TE
$wb8 = $xl.workbooks.open($TE, $null, $false)          
$sh1_wb8 = $wb12.sheets.item(1)    
$sheetToCopy = $wb8.sheets.item("TE")   
$sheetToCopy.copy($sh1_wb8) 
$wb8.close($false)
#Pr
$wb7 = $xl.workbooks.open($Pr, $null, $false)          
$sh1_wb7 = $wb12.sheets.item(1)    
$sheetToCopy = $wb7.sheets.item("Pr")   
$sheetToCopy.copy($sh1_wb7) 
    #update Classification Manager
$shEF = $CMDB.sheets.item("Uniclass Table Pr")#destination
$used = $sheetToCopy.usedRange
$lastCell = $used.SpecialCells(11)
$lastRow = $lastCell.row
$lastRow2 = $lastRow + 5
[void]$sheetToCopy.Range("A4:A$lastRow").Copy()
[void]$shEF.Range("A9").PasteSpecial(-4163) 
[void]$sheetToCopy.Range("F4:F$lastRow").Copy()
[void]$shEF.Range("B9").PasteSpecial(-4163)
[void]$sheetToCopy.Range("A1").Copy()
[void]$shEF.Range("E3").PasteSpecial(-4163)
[void]$shEF.Range("F9:F$lastRow2").Copy()
[void]$shEF.Range("C9").PasteSpecial(-4163)
$wb7.close($false)
#Ss
$wb6 = $xl.workbooks.open($Ss, $null, $false)
$sh1_wb6 = $wb12.sheets.item(1)    
$sheetToCopy = $wb6.sheets.item("Ss")   
$sheetToCopy.copy($sh1_wb6) 
    #update Classification Manager
$shEF = $CMDB.sheets.item("Uniclass Table Ss")#destination
$used = $sheetToCopy.usedRange
$lastCell = $used.SpecialCells(11)
$lastRow = $lastCell.row
$lastRow2 = $lastRow + 5
[void]$sheetToCopy.Range("A4:A$lastRow").Copy()
[void]$shEF.Range("A9").PasteSpecial(-4163) 
[void]$sheetToCopy.Range("F4:F$lastRow").Copy()
[void]$shEF.Range("B9").PasteSpecial(-4163)
[void]$sheetToCopy.Range("A1").Copy()
[void]$shEF.Range("E3").PasteSpecial(-4163)
[void]$shEF.Range("F9:F$lastRow2").Copy()
[void]$shEF.Range("C9").PasteSpecial(-4163)
$wb6.close($false)
#EF
$wb5 = $xl.workbooks.open($EF, $null, $false) #source         
$sh1_wb5 = $wb12.sheets.item(1)    
$sheetToCopy = $wb5.sheets.item("EF")   
$sheetToCopy.copy($sh1_wb5)
    #update Classification Manager
$shEF = $CMDB.sheets.item("Uniclass Table EF")#destination
$used = $sheetToCopy.usedRange
$lastCell = $used.SpecialCells(11)
$lastRow = $lastCell.row
$lastRow2 = $lastRow + 5
[void]$sheetToCopy.Range("A4:A$lastRow").Copy()
[void]$shEF.Range("A9").PasteSpecial(-4163) 
[void]$sheetToCopy.Range("F4:F$lastRow").Copy()
[void]$shEF.Range("B9").PasteSpecial(-4163)
[void]$sheetToCopy.Range("A1").Copy()
[void]$shEF.Range("E3").PasteSpecial(-4163)
[void]$shEF.Range("F9:F$lastRow2").Copy()
[void]$shEF.Range("C9").PasteSpecial(-4163)
$wb5.close($false)
#SL
$wb4 = $xl.workbooks.open($SL, $null, $false)          
$sh1_wb4 = $wb12.sheets.item(1)    
$sheetToCopy = $wb4.sheets.item("SL")   
$sheetToCopy.copy($sh1_wb4) 
    #update Classification Manager
$shEF = $CMDB.sheets.item("Uniclass Table SL")#destination
$used = $sheetToCopy.usedRange
$lastCell = $used.SpecialCells(11)
$lastRow = $lastCell.row
$lastRow2 = $lastRow + 5
[void]$sheetToCopy.Range("A4:A$lastRow").Copy()
[void]$shEF.Range("A9").PasteSpecial(-4163) 
[void]$sheetToCopy.Range("F4:F$lastRow").Copy()
[void]$shEF.Range("B9").PasteSpecial(-4163)
[void]$sheetToCopy.Range("A1").Copy()
[void]$shEF.Range("E3").PasteSpecial(-4163)
[void]$shEF.Range("F9:F$lastRow2").Copy()
[void]$shEF.Range("C9").PasteSpecial(-4163)
$wb4.close($false)
#Ac
$wb3 = $xl.workbooks.open($Ac, $null, $false)          
$sh1_wb3 = $wb12.sheets.item(1)    
$sheetToCopy = $wb3.sheets.item("Ac")   
$sheetToCopy.copy($sh1_wb3) 
$wb3.close($false)
#En
$wb2 = $xl.workbooks.open($En, $null, $false)          
$sh1_wb2 = $wb12.sheets.item(1)    
$sheetToCopy = $wb2.sheets.item("En")   
$sheetToCopy.copy($sh1_wb2)
    #update Classification Manager
$shEF = $CMDB.sheets.item("Uniclass Table En")#destination
$used = $sheetToCopy.usedRange
$lastCell = $used.SpecialCells(11)
$lastRow = $lastCell.row
$lastRow2 = $lastRow + 5
[void]$sheetToCopy.Range("A4:A$lastRow").Copy()
[void]$shEF.Range("A9").PasteSpecial(-4163) 
[void]$sheetToCopy.Range("F4:F$lastRow").Copy()
[void]$shEF.Range("B9").PasteSpecial(-4163)
[void]$sheetToCopy.Range("A1").Copy()
[void]$shEF.Range("E3").PasteSpecial(-4163)
[void]$shEF.Range("F9:F$lastRow2").Copy()
[void]$shEF.Range("C9").PasteSpecial(-4163)
$wb2.close($false)
#Co
$wb1 = $xl.workbooks.open($Co, $null, $false)          
$sh1_wb1 = $wb12.sheets.item(1)    
$sheetToCopy = $wb1.sheets.item("Co")   
$sheetToCopy.copy($sh1_wb1) 
$wb1.close($false)

#Finished merging, copying and updating

# Close and save destination workbook
$wb12.close($true)
$CMDB.Save()
$CMDB.Close()
$xl.Quit()
spps -n excel

#Rename destination workbook and delete sources
Rename-Item $Ro "Uniclass2015-AllTables.xlsx"
Get-ChildItem -Path  $currentPath -Recurse | Where{$_.Name -like "??.xlsx"} | Remove-Item

#Tidy up
Remove-Variable -name Co,En,Ac,SL,EF,Ss,Pr,TE,PM,Zz,FI,Ro,tableCo,tableEn,tableAc,tableSL,tableEF,tableSs,tablePr,tableTE,tablePM,tableZz,tableFI,tableRo,currentPath,sheetToCopy,CMDB,CMDBpath,shEF,wb1,wb2,wb3,wb4,wb5,wb6,wb7,wb8,wb9,wb10,wb11,wb12,xl,sh1_wb1,sh1_wb2,sh1_wb3,sh1_wb4,sh1_wb5,sh1_wb6,sh1_wb7,sh1_wb8,sh1_wb9,sh1_wb10,sh1_wb11
[gc]::collect() 
[gc]::WaitForPendingFinalizers()

# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed. Now it comes to focus
[System.Windows.Forms.MessageBox]::Show($this, "We are done here. Great Success!`n`nThat took about: $(($endDTM-$startDTM).totalseconds) seconds`n(give or take a millisecond)")
