Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Runtime.InteropServices

Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.AttachRemove = Me.Factory.CreateRibbonButton
        Me.AttachExport = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Label = "TabVOLKAN"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.AttachRemove)
        Me.Group1.Items.Add(Me.AttachExport)
        Me.Group1.Label = "Email"
        Me.Group1.Name = "Group1"
        '
        'AttachRemove
        '
        Me.AttachRemove.Label = "Attachments Remove"
        Me.AttachRemove.Name = "AttachRemove"
        Me.AttachRemove.ShowImage = True
        '
        'AttachExport
        '
        Me.AttachExport.Label = "Attachments Export"
        Me.AttachExport.Name = "AttachExport"
        Me.AttachExport.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Label = "Calendar"
        Me.Group2.Name = "Group2"
        '
        'Button3
        '
        Me.Button3.Label = "Button3"
        Me.Button3.Name = "Button3"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Label = "Tasks"
        Me.Group3.Name = "Group3"
        '
        'Button4
        '
        Me.Button4.Label = "Button4"
        Me.Button4.Name = "Button4"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AttachRemove As Microsoft.Office.Tools.Ribbon.RibbonButton
    Private Sub ReplaceCharsForFileName3(sName As String, sChr As String)
        sName = Replace(sName, "RE: ", "")
        sName = Replace(sName, "FW: ", "")
        sName = Replace(sName, "/", sChr)
        sName = Replace(sName, "\", sChr)
        sName = Replace(sName, ":", sChr)
        sName = Replace(sName, "?", sChr)
        sName = Replace(sName, Chr(34), sChr)
        sName = Replace(sName, "<", sChr)
        sName = Replace(sName, ">", sChr)
        sName = Replace(sName, "|", sChr)
        sName = Replace(sName, "&", sChr)
        sName = Replace(sName, "%", sChr)
        sName = Replace(sName, "*", sChr)
        sName = Replace(sName, " ", sChr)
        sName = Replace(sName, "{", sChr)
        sName = Replace(sName, "[", sChr)
        sName = Replace(sName, "]", sChr)
        sName = Replace(sName, "}", sChr)
    End Sub
    Function SaveSelectedAsDoc(saveFolder As String, Selection As Outlook.Selection)

        Dim currentExplorer As Outlook.Explorer
        Dim aItem As Object
        Dim newItem, newItemTemp As Outlook.MailItem = Nothing
        Dim dtDate As Date
        Dim sName As String
        Dim newFName As String
        Dim saveFileName As String

        Dim overwrite As Boolean

        'currentExplorer =
        'Selection = currentExplorer.Selection
        newItem = Application 'Outlook.Application 'Outlook.MailItem '.CreateItem(olMailItem)

        For Each aItem In Selection
            newItemTemp = aItem
            newItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText 'OlFormatText 'olFormatRichText

            sName = aItem.Subject
            ReplaceCharsForFileName3(sName, " ")
            dtDate = aItem.ReceivedTime
            sName = String.Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & String.Format(dtDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem) & "-" & sName

            saveFileName = saveFolder & "\" & sName & ".doc"
            overwrite = False
            '***********Display input box when file exists
            If Dir(saveFileName) <> vbNullString Then
                newFName = InputBox("The file already exists. Enter a new file name or OK overwrite.", "Confirm File Name", sName)
            Else : newFName = InputBox("Confirm File name:", "Confirm File Name", sName)
            End If
            If newFName = vbNullString Then GoTo skipfile
            If newFName = sName Then overwrite = True Else : sName = newFName
            saveFileName = saveFolder & "\" & sName & ".doc"
            '***********

            aItem.SaveAs(saveFileName, Outlook.OlBodyFormat.olFormatRichText) 'olRTF)
            '   aItem.SaveAs "C:\Mail\" & sName & ".doc", olRTF
skipfile:

        Next aItem

        currentExplorer = Nothing
        Selection = Nothing
        aItem = Nothing
        newItem = Nothing
        newItemTemp = Nothing

#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths
    Function BrowseForFolder(Optional Prompt As String = "C:\", Optional OpenAt As Object = "C:\") As String
        If OpenAt Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(OpenAt))
        End If

        Dim ShellApp As Object
        ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, Prompt, &H4000, 16)

        On Error Resume Next
        BrowseForFolder = ShellApp.self.Path
        On Error GoTo 0
        ShellApp = Nothing

        'Check for invalid or non-entries and send to the Invalid error handler if found
        'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename.  All others are invalid
        Select Case Mid(BrowseForFolder, 2, 1)
            Case Is = ":" : If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
            Case Is = "\" : If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
            Case Else : GoTo Invalid
        End Select

        Exit Function
Invalid:
        'If it was determined that the selection was invalid, set to False
        BrowseForFolder = vbNullString
    End Function


    Private Sub Button1_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles AttachRemove.Click
        'Remove Attachment
        Dim objOL As Outlook.Application
        Dim objMsg As Object
        'Dim olMail As Outlook.MailItem
        Dim objAttachments As Outlook.Attachments
        Dim objSelection As Outlook.Selection
        Dim i As Long, lngCount As Long
        Dim filesRemoved As String, fName As String, fSize As Long, fType As String, saveFolder As String, savePath As String
        Dim alterEmails As Boolean, overwrite As Boolean
        Dim result

        saveFolder = "C:\Users\evolozd\Desktop\00 DELETE"
        'If saveFolder = vbNullString Then Exit Sub

        'result = MsgBox("Do you want to remove attachments from selected file(s)? " & vbNewLine, vbYesNo + vbQuestion)
        result = vbYes
        alterEmails = (result = vbYes)

        objOL = CreateObject("Outlook.Application")
        objSelection = objOL.ActiveExplorer.Selection

        For Each objMsg In objSelection
            If (TypeOf objMsg Is Outlook.MailItem) Then 'objMsg.Class = Outlook.MailItem Then
                objAttachments = objMsg.Attachments
                lngCount = objAttachments.Count
                If lngCount > 0 Then
                    filesRemoved = ""
                    For i = lngCount To 1 Step -1
                        fName = objAttachments.Item(i).FileName
                        fType = Right(objAttachments.Item(i).FileName, 4)
                        fSize = objAttachments.Item(i).Size
                        If fSize < 100000 And (fType = ".jpg" Or fType = ".gif" Or fType = ".png") Then GoTo skipfile
                        'MsgBox objAttachments.Item(i).Type
                        savePath = saveFolder & "\" & fName
                        overwrite = True

                        objAttachments.Item(i).SaveAsFile(savePath)

                        If alterEmails Then
                            filesRemoved = filesRemoved & "<br>""" & objAttachments.Item(i).FileName & """ (" & objAttachments.Item(i).Size & ") " & "<a href=""" & savePath & """>[Link]</a>"
                            objAttachments.Item(i).Delete()
                        End If
skipfile:
                    Next i
                    If alterEmails Then
                        filesRemoved = "<font size=""2"" color=""red""" & "<b><u>Attachments removed: </u>" & filesRemoved & "</font></b><br><br>"

                        Dim objDoc As Object
                        Dim objInsp As Outlook.Inspector
                        objInsp = objMsg.GetInspector
                        objDoc = objInsp.WordEditor

                        objMsg.HTMLBody = filesRemoved + objMsg.HTMLBody
                        objMsg.Save
                    End If
                End If
            End If
        Next

ExitSub:
        objAttachments = Nothing
        objMsg = Nothing
        objSelection = Nothing
        objOL = Nothing
    End Sub
    Function formatSize(size As Long) As String
        Dim val As Double, newVal As Double
        Dim unit As String

        val = size
        unit = "bytes"

        newVal = Math.Round(val / 1024, 1)
        If newVal > 0 Then
            val = newVal
            unit = "KB"
        End If
        newVal = Math.Round(val / 1024, 1)
        If newVal > 0 Then
            val = newVal
            unit = "MB"
        End If
        newVal = Math.Round(val / 1024, 1)
        If newVal > 0 Then
            val = newVal
            unit = "GB"
        End If

        formatSize = val & " " & unit
    End Function
    Friend WithEvents AttachExport As Microsoft.Office.Tools.Ribbon.RibbonButton

    Private Sub Button2_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles AttachExport.Click
        'Export Attachments
        Dim objOL As Outlook.Application
        Dim objMsg As Outlook._MailItem
        Dim objAttachments As Outlook.Attachments
        Dim objSelection As Outlook.Selection
        Dim i As Long, lngCount As Long
        Dim filesRemoved As String, fName As String, fSize As Long, fType As String
        Dim saveFolder As String
        Dim savePath As String
        Dim alterEmails As Boolean, overwrite As Boolean
        Dim result

        saveFolder = BrowseForFolder("Select the folder to save attachments to.", "C:\")
        If saveFolder = vbNullString Then Exit Sub

        result = MsgBox("Do you want to remove attachments from selected file(s)? " & vbNewLine, vbYesNo + vbQuestion)
        alterEmails = (result = vbYes)

        objOL = New Outlook.Application 'CreateObject("Outlook.Application")
        objSelection = objOL.ActiveExplorer.Selection

        For Each objMsg In objSelection
            If (TypeOf objMsg Is Outlook.MailItem) Then
                objAttachments = objMsg.Attachments
                lngCount = objAttachments.Count
                If lngCount > 0 Then
                    filesRemoved = ""
                    For i = lngCount To 1 Step -1
                        If objAttachments.Item(i).DisplayName = "Picture (Device Independent Bitmap)" Then GoTo skipfile
                        fName = objAttachments.Item(i).FileName
                        fType = Right(objAttachments.Item(i).FileName, 4)
                        fSize = objAttachments.Item(i).Size
                        If fSize < 100000 And (fType = ".jpg" Or fType = ".gif" Or fType = ".png") Then GoTo skipfile
                        'If Right(objAttachments.Item(i).FileName, 4) = ".jpg" Or Right(objAttachments.Item(i).FileName, 4) = ".gif" Then GoTo skipfile
                        'MsgBox objAttachments.Item(i).Type
                        savePath = saveFolder & "\" & fName
                        overwrite = False
                        While Dir(savePath) <> vbNullString And Not overwrite
                            Dim newFName As String
                            newFName = InputBox("The file '" & fName &
                                "' already exists. Please enter a new file name, or just hit OK overwrite.", "Confirm File Name", fName)
                            If newFName = vbNullString Then GoTo skipfile
                            If newFName = fName Then overwrite = True Else fName = newFName
                            savePath = saveFolder & "\" & fName
                        End While


                        objAttachments.Item(i).SaveAsFile(savePath)

                        If alterEmails Then
                            'folderName = "<br>""" & "<a href=""" & saveFolder2 & """>[Folder]</a>"
                            filesRemoved = filesRemoved & "<br>""" & objAttachments.Item(i).FileName & """ (" &
                                                                    formatSize(objAttachments.Item(i).Size) & ") " &
                                                                    """" & "<a href=""" & saveFolder & """>[Folder]</a>" & " " &
                                "<a href=""" & savePath & """>[File]</a>"
                            objAttachments.Item(i).Delete()
                        End If
skipfile:
                    Next i

                    If alterEmails Then
                        'folderName = "<font size=""2"" color=""red""" & "<b><u>Folder: </u>" & folderName & "</font></b><br><br>"
                        filesRemoved = "<font size=""2"" color=""red""" & "<b><u>Attachments removed: </u>" & filesRemoved & "</font></b><br><br>"

                        Dim objDoc As Object
                        Dim objInsp As Outlook.Inspector
                        objInsp = objMsg.GetInspector
                        objDoc = objInsp.WordEditor

                        objMsg.HTMLBody = filesRemoved + objMsg.HTMLBody
                        objMsg.Save
                    End If
                End If
            End If
        Next
        Call SaveSelectedAsDoc(saveFolder, objSelection)
ExitSub:
        objAttachments = Nothing
        objMsg = Nothing
        objSelection = Nothing
        objOL = Nothing
    End Sub

    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class

