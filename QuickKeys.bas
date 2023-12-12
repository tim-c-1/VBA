Attribute VB_Name = "QuickKeys"
Option Explicit

Sub InsertCurrentTime()
Attribute InsertCurrentTime.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' InsertCurrentTime Macro
' Inserts exact current time and formats to time w/ seconds. Converts to static text.
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    With Selection
        .Formula = "=NOW()"
        .NumberFormat = "h:mm:ss"
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End With
    Application.CutCopyMode = False
End Sub


Sub MarkRowComplete()
Attribute MarkRowComplete.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' MarkRowComplete Macro
' mark row complete, drop to next, copy cell
'
' Keyboard Shortcut: Ctrl+Shift+C
'

'designate active row to avoid coloring entire row.
Dim colcount As Long
colcount = ActiveSheet.UsedRange.Columns.Count

'uses "good" style as "completed".
Range(Cells(Selection.Row, 1), Cells(Selection.Row, colcount)).Style = "Good"

'drops down to next cell in list and copies the contents
    ActiveCell.Offset(1, 0).Select
    Selection.Copy
    
End Sub

Sub PasteValues()
Attribute PasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' PasteAsValues Macro
' Pastes contents of clipboard as values in active cell.
'
' Keyboard Shortcut: Ctrl+Shift+V
'
    Selection.PasteSpecial Paste:=xlPasteValues
End Sub

Sub CompRateEmail()

Dim objOL As Outlook.Application
Dim objns As Object
Dim objfolder As Object

Set objOL = New Outlook.Application
Set objns = objOL.GetNamespace("MAPI")
Set objfolder = objns.GetDefaultFolder(olFolderInbox)

Dim NewEmail As Outlook.MailItem
Set NewEmail = Outlook.CreateItemFromTemplate("C:\Users\tcooney\AppData\Roaming\Microsoft\Templates\Comp Rate Query.oft")

With NewEmail
    .Subject = "Comp Rate Query | " & Date
    .SendUsingAccount
    .Display
    
End With

End Sub

Sub NewCheckPullSheet()
'
' NewCheckPullSheet Macro
'

'
    Dim ns As Worksheet
    Dim nsName As String
    nsName = Format(Date, "mm.dd.yy")
    
    ActiveSheet.Copy after:=Sheets(Sheets.Count)
    Set ns = ActiveSheet
    ns.name = nsName
    Range("B22:K50").Select
    Selection.ClearContents
    Range("B22:K50").Select

End Sub

Sub emailedFOAcomment()
Attribute emailedFOAcomment.VB_ProcData.VB_Invoke_Func = "A\n14"
'Insert new threaded comment that reads "Emailed FOA"
'Keyboard Shortcut: ctrl+shift+a

Selection.AddCommentThreaded "Emailed FOA"

End Sub

Sub replytocomment()
Attribute replytocomment.VB_ProcData.VB_Invoke_Func = "R\n14"
'reply to existing comment on currently selected cell
'Keyboard Shortcut: ctrl+shift+r

Dim tdate As String
tdate = Format(Date, "mm/dd/yyyy")
Dim userreply As String
userreply = InputBox("What would you like to say?")

If userreply = "" Then
 Exit Sub
 
Else: Selection.CommentThreaded.AddReply userreply

End If

End Sub


Function getcommentdate(xCell As Range) As String
'custom function to return date of most recent comment in a threaded comment.
'use in worksheets to easily track which items need follow-up.
'

Dim mostrecent As Long
mostrecent = xCell.CommentThreaded.Replies.Count

On Error Resume Next
If getcommentdate = Format(xCell.CommentThreaded.Replies.Item(mostrecent).Date, "mm/dd/yyyy") = "" Then
    getcommentdate = Format(xCell.CommentThreaded.Date, "mm/dd/yyyy")
Else
    getcommentdate = Format(xCell.CommentThreaded.Replies.Item(mostrecent).Date, "mm/dd/yyyy")
End If

End Function

