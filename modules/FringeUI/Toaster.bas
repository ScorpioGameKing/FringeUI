Attribute VB_Name = "Toaster"
'----- Timed Message Module
Sub Toast(msg As String, duration As Integer, title As String, button_type As Integer)
    CreateObject("WScript.Shell").PopUp msg, duration, title, button_type
End Sub
