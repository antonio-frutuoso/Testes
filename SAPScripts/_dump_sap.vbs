Dim SapGuiAuto, application, connection, session
On Error Resume Next
Set SapGuiAuto  = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection  = application.Children(0)
Set session     = connection.Children(0)
If Not IsObject(session) Then
   WScript.Echo "ERRO: sem sessao SAP scriptavel."
   WScript.Quit 1
End If

Sub Dump(obj, prefixo)
   On Error Resume Next
   Dim t, nm, tx
   t = obj.Type
   nm = obj.Name
   tx = obj.Text
   WScript.Echo prefixo & obj.Id & "  [" & t & "]  name=" & nm & "  text=" & Left(tx,50)
   Dim c
   For Each c In obj.Children
      Dump c, prefixo
   Next
   Err.Clear
End Sub

Sub DumpTela(tcode)
   On Error Resume Next
   session.findById("wnd[0]/tbar[0]/okcd").text = "/n" & tcode
   session.findById("wnd[0]").sendVKey 0
   WScript.Sleep 1500
   WScript.Echo "================ " & tcode & " ================"
   Dim w
   For Each w In session.Children
      WScript.Echo "----- janela " & w.Id & " (titulo: " & w.Text & ") -----"
      Dump w, "  "
   Next
   Err.Clear
End Sub

DumpTela "CG3Z"
DumpTela "ZUP_REQUEST"
WScript.Echo "FIM"
WScript.Quit 0
