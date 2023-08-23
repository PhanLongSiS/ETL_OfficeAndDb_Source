#---------------------------------------------------------------
#  Global Configuration
#---------------------------------------------------------------
$WordExtractor = ".\ExtractWord.exe"
$WordImportor = ".\ModifyWord.exe"

# Input file to export cell values to json
#$InputExcelFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.xlsx"
# Get filename from the commandline parameter
$InputWordFile = $Args[0]
$OutputWordFile = $Args[1]

#---------------------------------------------------------------
#  Main
#---------------------------------------------------------------

# Input json to get cell positions and update cell values
#$InOutJsonFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.json"
$InOutJsonFile = $InputWordFile.Substring(0, $InputWordFile.Length - 4) + "json"

Write-Output "--------------------------------Excel --> Json ----------------------------"
Write-Output "  from: $ExcelImportor"
Write-Output "    to: $InOutJsonFile"
Write-Output "  Bat dau extract..."
Write-Output "$WordExtractor -f ""$InputWordFile""  -j "".\$InOutJsonFile"""
Invoke-Expression "$WordExtractor -f ""$InputWordFile""  -j "".\$InOutJsonFile"""
Write-Output "    Xong 1/2"


# $TemplateExcelFile = $ExcelExtractor
# $GeneratedExcelFile = "A5300B_Lab_1.B.1_v1.0(Biochemlab)done.xlsx"

Write-Output "--------------------------------Json --> Excel ----------------------------"
Write-Output "    to: $OutputWordFile"
Invoke-Expression "$WordImportor -f ""$OutputWordFile""  -j "".\$InOutJsonFile"""