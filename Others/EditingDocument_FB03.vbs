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

' Ingremos a la transaccion
session.StartTransaction ("fb03")

' Ingresamos al documento a editar
session.findById("wnd[0]/usr/txtRF05L-BELNR").text = "100000000"
session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = "0001"
session.findById("wnd[0]/usr/txtRF05L-GJAHR").text = "2000"
session.findById("wnd[0]").sendVKey 0

' Activamos modo edición
session.findById("wnd[0]").sendVKey 25

' Seteamos la tabla contenedora
Set tbl = session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell")

' Recorrer el listado
accountSearched = "3"
text = "Edited by scripting :)"
For i = 0 To (tbl.rowCount - 1)
   ' Obtener valor de account y comprobar
   If tbl.getCellValue(CInt(i), "KTONR") = accountSearched Then
      ' Ingresamos a la row
      tbl.currentCellRow = CInt(i)
      tbl.selectedRows = CStr(i)
      tbl.doubleClickCurrentCell

      ' Editar el texto
      session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = text

      ' Back
      session.findById("wnd[0]").sendVKey 3
      Exit For
   End If
Next

' Save document
session.findById("wnd[0]").sendVKey 11