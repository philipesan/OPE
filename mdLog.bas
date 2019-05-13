Attribute VB_Name = "mdLog"
Public Sub Log(ByVal text As String)

frmLog.tbLog.text = frmLog.tbLog.text & Time & vbTab & Date & " - " & text & vbCrLf

Dim flog As String

' Get a free file number
flog = FreeFile

' Create Test.txt
Open sLogPath For Output As flog

' Write the contents of TextBox1 to Test.txt
Print #flog, frmLog.tbLog.text

' Close the file
Close flog
End Sub
