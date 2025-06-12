Imports System.IO

Public Class ApiKeySettingForm
    Inherits Form

    Private txtApiKey As TextBox
    Private btnSave As Button
    Private lblInfo As Label
    Private apiKeyFilePath As String

    Private Sub InitializeComponents()
        Me.Text = "Pengaturan API Key"
        Me.Size = New Size(600, 180)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblInfo = New Label() With {
            .Text = "Masukkan API Key untuk Sistem Sigaluh:",
            .Location = New Point(15, 15),
            .AutoSize = True
        }

        txtApiKey = New TextBox() With {
            .Location = New Point(15, 40),
            .Size = New Size(550, 30),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        btnSave = New Button() With {
            .Text = "Simpan & Tutup",
            .Location = New Point(435, 85),
            .Size = New Size(130, 35),
            .BackColor = Color.FromArgb(40, 167, 69),
            .ForeColor = Color.White,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        AddHandler btnSave.Click, AddressOf OnSaveClick

        Me.Controls.AddRange({lblInfo, txtApiKey, btnSave})
    End Sub

    Public Sub New(ByVal filePath As String)
        Me.apiKeyFilePath = filePath
        InitializeComponents()
        LoadApiKey()
    End Sub

    Private Sub LoadApiKey()
        If File.Exists(apiKeyFilePath) Then
            txtApiKey.Text = File.ReadAllText(apiKeyFilePath).Trim()
        End If
    End Sub

    Private Sub OnSaveClick(sender As Object, e As EventArgs)
        Try
            File.WriteAllText(apiKeyFilePath, txtApiKey.Text.Trim())
            MessageBox.Show("API Key berhasil disimpan.", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Catch ex As Exception
            MessageBox.Show($"Gagal menyimpan API Key: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class