Dim objIE, objShell
Dim strDX

Set objIE = CreateObject("InternetExplorer.Application")
Set objShell = CreateObject("WScript.Shell")

strDX = "AT-0125B"

objIE.Navigate "about:blank"

objIE.Document.Title = "Covered Diagnosis"
objIE.ToolBar = False
objIE.Resizable = False
objIE.StatusBar = False
objIE.Width = 350
objIE.Height = 200
'objIE1250T.Scrollbars="no"

' Center the Window on the screen
With objIE.Document.ParentWindow.Screen
    objIE.Left = (.AvailWidth - objIE.Width ) \ 2
    objIE.Top = (.Availheight - objIE.Height) \ 2
End With

objIE.document.body.innerHTML = "<b>" & strDX & " is a covered diagnosis code.</b><p>&nbsp;</p>" & _
"<center><input type='submit' value='OK' onclick='VBScript:ClickedOk()'></center>" & _
"<input type='hidden' id='OK' name='OK' value='0'>"

objIE.Visible = True
'objShell.AppActivate "Covered Diagnosis"
'MsgBox objIE.Document.All.OK.Value
Function ClickedOk
'If objIE.Document.All.OK.Value = 1 Then
    'objIE.Document.All.OK.Value = 0
    'objShell.AppActivate "Covered Diagnosis"
    'ouit
    Window.Close()
'End If
End Function  
