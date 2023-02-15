If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

exportFile = "C:\Users\vboxuser\Documents\SAP\SAP GUI\tmp.txt"

'Ingresamos a la tx
session.StartTransaction "co03"

'Ingresamos a la orden
session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = "1000000"
session.findById("wnd[0]").sendVKey 0

'Guardamos texto corto
shortText = session.findById("wnd[0]/usr/txtCAUFVD-MATXT").text

'Hacemos foco sobre el texto e ingresamos al texto extendido
session.findById("wnd[0]/usr/txtCAUFVD-MATXT").setFocus
session.findById("wnd[0]").sendVKey 2

'Descargamos archivo
session.findById("wnd[0]/mbar/menu[0]/menu[4]").select
session.findById("wnd[1]/usr/radITCTK-TDASCII").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[2]/usr/ctxtITCTK-TDFILENAME").text = exportFile
session.findById("wnd[2]/usr/ctxtITCTK-TDCODEPAGE").text = "1152"
session.findById("wnd[2]/tbar[0]/btn[0]").press
'Reemplazamos si ya existe el archivo
If Not session.findById("wnd[2]", False) Is Nothing Then session.findById("wnd[2]").sendVKey 5

'Back
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

'Creamos objeto para abrir archivo
Set fso = CreateObject("Scripting.FileSystemObject")
'Abrimos archivo
Set targetFile = fso.OpenTextFile(exportFile, 1)
'Almacenamos texto
extendedText = Replace(targetFile.ReadAll, "##", vbNullString)
'Cerrar archivo y liberar objetos
targetFile.Close
Set fso = Nothing
Set targetFile = Nothing

MsgBox "Texto corto:" & vbnewline & shortText & vbnewline & vbnewline & _
"Texto extendido:" & vbnewline & extendedText