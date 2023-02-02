'
' Alejandro Franzoni Gimenez
' https://www.youtube.com/@alefranzoni
'
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

order = "4000020"
text = "Este es un texto largo de ejemplo. Múltiples líneas"

'Ingresamos a la tx
session.StartTransaction "iw32"
'Ingresamos a la orden
session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = order
session.findById("wnd[0]").sendVKey 17
'Ingresamos texto
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7,1]").text = Mid(textToUnicode(text),1,40)
session.findById("wnd[0]").sendVKey 0

If Len(text) > 40 Then
   'Presionamos boton texto extendido
   session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").getAbsoluteRow(1).selected = true
   session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8,1]").press
   'Borrar contenido
   session.findById("wnd[0]").sendVKey 14
   session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
   'Pegar contenido del clipboard
   'VBS
   putTextInClipboard text
   session.findById("wnd[0]").sendVKey 9
   'Back
   session.findById("wnd[0]").sendVKey 3
   'Deseleccion
   session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").getAbsoluteRow(1).selected = false
End If

Sub putTextInClipboard(text)
   WScript.CreateObject("WScript.Shell").Run "cmd.exe /c echo " & textToUnicode(text) & " | clip", 0, True 
End Sub

Function textToUnicode(textInput)
   With CreateObject("ADODB.Stream")
      .Open
      .Charset = "Windows-1252"
      .WriteText textInput
      .Position = 0
      .Type = 2
      .Charset = "utf-8"
      textToUnicode = .ReadText(-1)
      .Close
   End With
End Function