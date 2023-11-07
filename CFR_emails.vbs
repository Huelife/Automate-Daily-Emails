Option Explicit

Dim xlApp, xlBook, path
path = "C:\Users\To\File\Location\CFR_emails_macro.xlsm"

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open(path, 0, True)

xlApp.Run "'" & path & "'!CFR_emails"

xlBook.Close
xlApp.Quit

Set xlApp = Nothing
Set xlBook = Nothing

WScript.Quit
