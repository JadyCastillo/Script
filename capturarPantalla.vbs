Set obj = CreateObject("wscript.shell")

Function CloseAplication()
   Dim objWMIService : Set objWMIService = GetObject("winmgmts:")
   Dim ProcessList, process
   Set ProcessList= objWMIService.ExecQuery _
   ("select * from Win32_Process")

   For Each process In ProcessList
      if process.Name="mspaint.exe" Then
         process.terminate
         Exit For
      End If
   Next
End Function

'Cierra la aplicación paint si encuentra un proceso ya activo
CloseAplication()

'Presionamos la tecla Impr pant para realizar la captura de pantalla
WScript.Sleep 2000
obj.SendKeys "{PRTSC}"
WScript.Sleep 2000

obj.Run "mspaint.exe"
WScript.Sleep 1000


'Activar la aplicación de Paint
obj.AppActivate "Sin título - Paint"
WScript.Sleep 1000
 
'Pegar la la captura de pantalla en paint
obj.SendKeys "^v"
WScript.Sleep 500
 
'Guardar la imagen
obj.SendKeys "^g"
WScript.Sleep 500

'Colocamos la ruta y su nombre para guardar imagen
obj.SendKeys "C:\captura.png"
WScript.Sleep 500
obj.SendKeys "{ENTER}"

'Cierra la aplicación paint
CloseAplication()
