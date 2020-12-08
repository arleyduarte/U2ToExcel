# U2ToExcel


#Para ejecutar
U2ToExcel.exe RUTA_ORIGIEN RUTA_DESTINO

Ejemplo script power shell
Write-Host "U2ToExcel ...."
$execFile = "C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\bin\Debug\net5.0\U2ToExcel.exe"
$params = "C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\Resources\REP-ORIGINAL.csv C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\bin\Debug\net5.0\Destion.xlsx"

# Wait until the started process has finished
&$execFile $params
