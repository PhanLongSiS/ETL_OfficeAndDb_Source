#---------------------------------------------------------------
#  Global Configuration
#---------------------------------------------------------------
$ExcelExtractor = ".\ExtractExcel.exe"
$ExcelImportor = ".\ModifyExcel.exe"

# Input file to export cell values to json
#$InputExcelFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.xlsx"
# Get filename from the commandline parameter
$InputExcelFile = $Args[0]
$OutputExcelFile = $Args[1]

#---------------------------------------------------------------
#  Main
#---------------------------------------------------------------

# Input json to get cell positions and update cell values
#$InOutJsonFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.json"
$InOutJsonFile = $InputExcelFile.Substring(0, $InputExcelFile.Length - 4) + "json"

Write-Output "--------------------------------Excel --> Json ----------------------------"
Write-Output "  from: $ExcelImportor"
Write-Output "    to: $InOutJsonFile"
Write-Output "  Bat dau extract..."
Write-Output "$ExcelExtractor -f ""$InputExcelFile""  -j "".\$InOutJsonFile"""
Invoke-Expression "$ExcelExtractor -f ""$InputExcelFile""  -j "".\$InOutJsonFile"""
Write-Output "    Xong 1/2"


# $TemplateExcelFile = $ExcelExtractor
# $GeneratedExcelFile = "A5300B_Lab_1.B.1_v1.0(Biochemlab)done.xlsx"

Write-Output "--------------------------------Json --> Excel ----------------------------"
Write-Output "  from: $TemplateExcelFile  and  $InOutJsonFile"
Write-Output "    to: $OutputExcelFile"
Invoke-Expression "$ExcelImportor -f ""$OutputExcelFile""  -j "".\$InOutJsonFile"""