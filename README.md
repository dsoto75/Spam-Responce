# Spam-Responce
Outlook VBA code to reply to spam email without personal signature
Sub ReplytoSpam()
Dim Item As Outlook.MailItem
Set Item = Application.ActiveExplorer.Selection.Item(1)

Dim olInspector As Outlook.Inspector
Dim olDocument As Word.Document
Dim olSelection As Word.Selection

'Begin reply to active email routein

Set myReply = Item.Reply
myReply.Display

If Item.Subject <> "Unsubscribe" Then

Item.Subject = "Unsubscribe"

End If

Set olInspector = myReply.GetInspector
Set olDocument = olInspector.WordEditor
Set olSelection = olDocument.Application.Selection

olSelection.InsertBefore "Unsubscribe, Opt-Out, Leave Out, Stop, Can Spam"
                                              
' Begin signature removal routein

Set oBookmark = olDocument.Bookmarks("_MailAutoSig")

If Not oBookmark Is Nothing Then
   oBookmark.Select
   olDocument.Windows(1).Selection.Delete
End If

'Uncomment to send mail automatically
'myReply.Send

End Sub
