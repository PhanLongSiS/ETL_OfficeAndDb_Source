

Copy-Item -Path ".\ExtractExcel" -Destination ".\..\..\ETL_Office\ExtractExcel" -Recurse
Copy-Item -Path ".\ExtractWord" -Destination ".\..\..\ETL_Office\ExtractWord" -Recurse
Copy-Item -Path ".\ModifyExcel" -Destination ".\..\..\ETL_Office\ModifyExcel" -Recurse
Copy-Item -Path ".\ModifyWord" -Destination ".\..\..\ETL_Office\ModifyWord" -Recurse
Copy-Item -Path ".\TransformExJsToWordJs" -Destination ".\..\..\ETL_Office\TransformExJsToWordJs" -Recurse
Copy-Item -Path ".\TransformWordJsToExJs" -Destination ".\..\..\ETL_Office\TransformWordJsToExJs" -Recurse

