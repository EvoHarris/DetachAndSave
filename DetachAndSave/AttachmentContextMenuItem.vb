Imports System.Windows.Forms

<Runtime.InteropServices.ComVisible(True)> _
Public Class AttachmentContextMenuItem
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("DetachAndSave.AttachmentContextMenuItem.xml")
    End Function

#Region "Menu call back"
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Function getDetachAndSaveLabel(ByVal control As Microsoft.Office.Core.IRibbonControl) As String
        Dim singular As String = "Anhang speichern und entfernen"
        Dim plural As String = "Anhänge speichern und entfernen"

        Dim context As Object = control.Context
        getDetachAndSaveLabel = "???"

        If (TypeOf context Is Outlook.AttachmentSelection) Then
            Dim selection As Outlook.AttachmentSelection = CType(context, Outlook.AttachmentSelection)
            If selection.Count > 1 Then
                getDetachAndSaveLabel = plural
            ElseIf selection.Count = 1 Then
                getDetachAndSaveLabel = singular
            Else
                getDetachAndSaveLabel = "???"
            End If
        End If
    End Function

    Public Sub OnDetachAndSave(ByVal control As Office.IRibbonControl)
        Dim context As Object = control.Context
        If (TypeOf context Is Outlook.AttachmentSelection) Then
            Dim selection As Outlook.AttachmentSelection = CType(context, Outlook.AttachmentSelection)
            Dim attachmentFiles(selection.Count) As String

            If selection.Count = 1 Then
                Dim attachment As Outlook.Attachment = selection.Item(1)
                Dim saveFileDialog1 As New SaveFileDialog()

                saveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments
                saveFileDialog1.FileName = attachment.FileName
                System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment)
                If saveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    attachmentFiles(1) = saveFileDialog1.FileName
                Else
                    Exit Sub
                End If
            ElseIf selection.Count > 1 Then
                Dim gotAttDir As Boolean
                Do
                    gotAttDir = True
                    Dim dirDlg As New Windows.Forms.FolderBrowserDialog
                    dirDlg.RootFolder = Environment.SpecialFolder.MyDocuments
                    If dirDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        For i As Integer = 1 To selection.Count
                            Dim attachment As Outlook.Attachment = selection.Item(i)
                            attachmentFiles(i) = System.IO.Path.Combine(dirDlg.SelectedPath, attachment.FileName)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment)
                            If System.IO.File.Exists(attachmentFiles(i)) Then
                                Dim result = MessageBox.Show(System.IO.Path.GetFileName(attachmentFiles(i)) + _
                                                             " ist bereits vorhanden." + vbCrLf + _
                                                             "Möchten Sie sie ersetzen?",
                                                             "Speichern unter bestätigen",
                                                             MessageBoxButtons.YesNo)
                                If result = DialogResult.No Then
                                    gotAttDir = False
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        Exit Sub
                    End If
                Loop Until gotAttDir
            Else
                Exit Sub
            End If

            Try
                For i As Integer = 1 To selection.Count
                    Dim attachment As Outlook.Attachment = selection.Item(i)
                    attachment.SaveAsFile(attachmentFiles(i))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment)
                Next
            Catch e As Exception
                MsgBox("Fehler beim Speichern: " + e.Message)
                Exit Sub
            End Try

            For i As Integer = 1 To selection.Count
                Dim attachment As Outlook.Attachment = selection.Item(i)
                Dim current As Outlook.MailItem = CType(attachment.Parent, Outlook.MailItem)
                current.Body = "Anhang " + attachment.DisplayName + " wurde als " +
                    "<file://" + attachmentFiles(i) +
                    "> gespeichert und entfernt." + vbCrLf + current.Body
                'System.Diagnostics.Debug.Print(current.Body)
                attachment.Delete()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(current)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment)
            Next
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(context)
    End Sub


#End Region

#Region "helper programs"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
