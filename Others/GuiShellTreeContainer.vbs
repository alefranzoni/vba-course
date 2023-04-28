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

' Get container type
Set tbl = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell")
' MsgBox tbl.type
' MsgBox tbl.subtype

' Get nodes count
Set nodes = tbl.GetAllNodeKeys
' MsgBox nodes.Count

' Get columns counts
' MsgBox "Cantidad de columnas: " & tbl.GetColumnHeaders.Count

' Node key and column name
' MsgBox "Nodo <" & tbl.SelectedItemNode & ">"
' MsgBox "Columna <" & tbl.SelectedItemColumn & ">"

' Get nodes keys
' nodesKeys = vbnullstring
' For i = 0 To (nodes.Count - 1)
'    If (nodesKeys = vbnullstring) Then
'       nodesKeys = "[" & nodes.Item(CInt(i)) & "]" 
'    Else
'       nodesKeys = nodesKeys & vbnewline & "[" & nodes.Item(CInt(i)) & "]" 
'    End If
' Next
' MsgBox nodesKeys

' Get text by node key
' MsgBox tbl.GetNodeTextByKey("          1")

' Get text by key and column
' MsgBox tbl.GetItemText("          1", "2")

' Folder expandable
' MsgBox tbl.IsFolderExpandable("          5")
' tbl.ExpandNode("          5")
' ' tbl.CollapseNode("          5")
' MsgBox tbl.IsFolderExpanded("          5")

' Search for a value in list
' Set column = tbl.GetColumnCol(CInt(1))
' searchValue = "SAPS-ALL-AA-01"
' For i = 0 To (column.Count - 1)
'    If column.Item(CInt(i)) = searchValue Then
'       MsgBox "Descripción del nodo " & column.Item(CInt(i)) & vbnewline & tbl.GetColumnCol(2).Item(CInt(i))
'       Exit For
'    End If 
' Next