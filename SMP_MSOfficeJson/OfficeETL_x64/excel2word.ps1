#---------------------------------------------------------------
#  Global Configuration
#---------------------------------------------------------------
$ExcelExtractor = ".\ExtractExcel.exe"
$WordImportor = ".\ModifyWord.exe"
$TransformJsonExcelToWord=".\TransformExJsToWordJs.exe"

# Input file to export cell values to json
#$InputExcelFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.docx"
# Get filename from the commandline parameter
$InputExcelFile = $Args[0]
$OutputWordFile = $Args[1]

#---------------------------------------------------------------
#  Main
#---------------------------------------------------------------

# Input json to get cell positions and update cell values
#$InOutJsonFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.json"
$InOutJsonFile = $InputExcelFile.Substring(0, $InputExcelFile.Length - 4) + "json"

Write-Output "--------------------------------Exel --> Json ----------------------------"
Write-Output "  Bat dau extract..."
Invoke-Expression "$ExcelExtractor -f ""$InputExcelFile""  -j "".\$InOutJsonFile"""
Write-Output "    Xong 1/3"

Write-Output "--------------------------------Exel Json -->Word Json ----------------------------"
Write-Output "  Bat dau Transform"
Invoke-Expression "$TransformJsonExcelToWord -e "".\$InOutJsonFile"""
Write-Output "    Xong 1/3"

Write-Output "--------------------------------Json --> Word ----------------------------"
Write-Output "  Bat dau Import"
Invoke-Expression "$WordImportor -f ""$OutputWordFile""  -j "".\$InOutJsonFile"""