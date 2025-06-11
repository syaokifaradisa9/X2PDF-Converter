Imports System.IO

Public Class SheetSettingForm
    Inherits Form

    Private sheetFilePath As String
    Private txtContent As TextBox
    Private btnSave As Button

    Public Sub New(filePath As String)
        ' Pastikan memanggil konstruktor Form
        MyBase.New()

        Me.sheetFilePath = filePath
        SetupFormUI()
    End Sub

    Private Sub SetupFormUI()
        Me.Text = "Setting Sheet"
        Me.Size = New Size(400, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        txtContent = New TextBox() With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Location = New Point(20, 20),
            .Size = New Size(340, 270)
        }
        Me.Controls.Add(txtContent)

        btnSave = New Button() With {
            .Text = "Simpan",
            .Location = New Point(260, 310),
            .Size = New Size(100, 30)
        }
        AddHandler btnSave.Click, AddressOf OnSaveClick
        Me.Controls.Add(btnSave)

        If File.Exists(sheetFilePath) Then
            txtContent.Text = File.ReadAllText(sheetFilePath)
        End If
    End Sub

    Private Sub OnSaveClick(sender As Object, e As EventArgs)
        Try
            File.WriteAllText(sheetFilePath, txtContent.Text)
            MessageBox.Show("Sheet berhasil disimpan!", "Berhasil", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Gagal menyimpan sheet: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
