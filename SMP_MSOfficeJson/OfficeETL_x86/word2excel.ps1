#---------------------------------------------------------------
#  Global Configuration
#---------------------------------------------------------------
$WordExtractor = ".\ExtractWord.exe"
$ExcelImportor = ".\ModifyExcel.exe"
$TransformJsonWordToExcel=".\TransformWordJsToExJs.exe"

# Input file to export cell values to json
#$InputExcelFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.docx"
# Get filename from the commandline parameter
$InputWordFile = $Args[0]
$OutputExcelFile = $Args[1]

#---------------------------------------------------------------
#  Main
#---------------------------------------------------------------

# Input json to get cell positions and update cell values
#$InOutJsonFile = "A5300B_Lab_1.A.1_v1.0(MicrolabBM)done.json"
$InOutJsonFile = $InputWordFile.Substring(0, $InputWordFile.Length - 4) + "json"

Write-Output "--------------------------------Word --> Json ----------------------------"
Write-Output "  Bat dau extract..."
Invoke-Expression "$WordExtractor -f ""$InputWordFile""  -j "".\$InOutJsonFile"""
Write-Output "    Xong 1/3"

Write-Output "--------------------------------Word Json -->Excel Json ----------------------------"
Write-Output "  Bat dau Transform"
Invoke-Expression "$TransformJsonWordToExcel -w "".\$InOutJsonFile"""
Write-Output "    Xong 2/3"

Write-Output "--------------------------------Json --> Excel ----------------------------"
Write-Output "  Bat dau Import"
Invoke-Expression "$ExcelImportor -f ""$OutputExcelFile""  -j "".\$InOutJsonFile"""