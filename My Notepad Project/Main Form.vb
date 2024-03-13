Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Public Class frmMain
    Private strFilePath As String = ""
    Private blnChanged As Boolean = False

    Private Sub FileToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.DropDownOpening
        If Me.blnChanged Then
            Me.SaveToolStripMenuItem.Enabled = True
            Me.NewToolStripMenuItem.Enabled = True
        Else
            Me.SaveToolStripMenuItem.Enabled = False
            Me.NewToolStripMenuItem.Enabled = False
        End If
        If Me.strFilePath = "" Then
            Me.SaveAsToolStripMenuItem.Enabled = False
        Else
            Me.SaveAsToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub DocumentTextbox_TextChanged(sender As Object, e As EventArgs) Handles DocumentTextbox.TextChanged
        Me.blnChanged = True
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        ''If the file has not been saved yet, then prompt the user to save the file
        If Me.strFilePath = "" Then
            With SaveFileDialog1
                .Title = "Save as Text File"
                .Filter = "Text Files|*.txt"
                .InitialDirectory = SpecialDirectories.MyDocuments
                If .ShowDialog() = DialogResult.OK Then
                    Me.strFilePath = .FileName
                    File.WriteAllText(Me.strFilePath, Me.DocumentTextbox.Text)
                End If
            End With
            ''If the file has been saved, then save the file
        Else
            File.WriteAllText(Me.strFilePath, Me.DocumentTextbox.Text)
        End If
        Me.blnChanged = False
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ''If the file has been changed, then prompt the user to save the file
        If Me.blnChanged Then
            SaveToolStripMenuItem_Click(sender, e)
        End If
        With OpenFileDialog1
            .Title = "Open Text File"
            .Filter = "Text Files|*.txt"
            .InitialDirectory = SpecialDirectories.MyDocuments
            .FileName = ""
            ''If the user selects a file, then open the file
            If .ShowDialog() = DialogResult.OK Then
                Me.strFilePath = .FileName
                ''If the file exists, then open the file
                Me.DocumentTextbox.Text = File.ReadAllText(Me.strFilePath)
                Me.blnChanged = False
            End If
        End With
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        If Me.blnChanged Then
            SaveToolStripMenuItem_Click(sender, e)
            DocumentTextbox.Text = ""
            Me.strFilePath = ""
            Me.blnChanged = False
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Me.strFilePath = ""
        SaveToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class
