Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New AttachmentContextMenuItem()
    End Function


    Private Sub Application_ItemSend(ByVal Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend
        If TypeOf Item Is Outlook.MailItem Then
            'Cancel = Not SaveSentMail(Item)
            SaveSentMail(Item)
        End If
        Cancel = False
    End Sub

    Private Sub SaveSentMail(Item As Outlook.MailItem)
        Dim f As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
        If Item.DeleteAfterSubmit = False Then
            System.Diagnostics.Debug.Print(Application.ActiveExplorer.CurrentFolder.FolderPath)
            F = Application.Session.PickFolder
        End If

        If Not F Is Nothing Then
            Item.SaveSentMessageFolder = Application.Session.GetFolderFromID(F.EntryID)
        End If
    End Sub

End Class
