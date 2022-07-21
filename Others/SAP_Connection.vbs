' --------------------------------------------------------------------
' Title         SAP Scripting: Establishing connection by SSO & Auth
' Author        Alejandro Franzoni Gimenez
' 
' Contact       contacto@alejandrofranzoni.com.ar
'               www.alejandrofranzoni.com.ar 
' --------------------------------------------------------------------

' SAP
SAP_BIN = "saplogon.exe"
SAP_GUI_PATH = "C:\Program Files (x86)\SAP\FrontEnd\SapGui\" & SAP_BIN

' Connections
lk0_sso = "- LK0 [SSO - SIC Domestico - TESTING PROYECTOS EHP6 SP7]"
lk0_auth = "LK0 [EHP6 SP7 SIC Domestico - TESTING]"
' Auth
user = "username"
password = "pw"

Main

Sub Main()
    If Not FileExists(SAP_GUI_PATH) Then
        MsgBox "El archivo no existe en la ruta especificada.", vbExclamation, "Archivo no encontrado"
        Exit Sub
    End If

    ExecuteAndWaitForSAP

    ' SapGuiApplication
    Set root = GetObject("SAPGUI")
    Set Application = root.GetScriptingEngine

    ' Open sync connection
    '
    ' Without SSO
    ' Set Connection = Application.OpenConnection(lk0_auth, True)
    ' Set Session = Connection.Children(0)
    ' Session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
    ' Session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    ' Session.findById("wnd[0]/tbar[0]/btn[0]").press
    '    
    ' SSO
    Set Connection = Application.OpenConnection(lk0_sso, True)
    Set Session = Connection.Children(0)

    ' Open any tx
    Session.StartTransaction("IW32")
    MsgBox "Ok!"

    ' Shutdown the connection
    Set Session = Nothing
    Connection.CloseSession("ses[0]")
    Set Connection = Nothing
    Set Application = Nothing
End Sub

Sub ExecuteAndWaitForSAP()
    ' Run saplogon bin
    WScript.CreateObject("WScript.Shell").Run Chr(34) & SAP_GUI_PATH & Chr(34), 2

    ' Wait to be initialized
    isSapInitialized = False
    Do While Not isSapInitialized
        isSapInitialized = IsProcessRunning(SAP_BIN)
    Loop
    
    WScript.Sleep 3000
End Sub

Function IsProcessRunning(targetProcess)
    Set WMIService = GetObject("winmgmts:\\.\root\cimv2")
    query = "SELECT * FROM Win32_Process"
    Set items = WMIService.ExecQuery(query)

    For Each item In items
        If item.Name = targetProcess Then
            IsProcessRunning = True
            Exit Function
        End If
    Next

    IsProcessRunning = False
End Function

Function FileExists(filePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then 
        FileExists = True 
    Else
        FileExists = False
    End If
End Function