Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form1
    ' Deklarasi UI
    Private btnInfo As Button
    Private labelExcel As Label
    Private labelPdf As Label
    Private txtExcelFolder As TextBox
    Private txtPdfFolder As TextBox
    Private btnBrowseExcel As Button
    Private btnBrowsePdf As Button
    Private btnConvert As Button
    Private progressBar As ProgressBar

    Public Sub New()
        Me.InitializeComponent()

        ' Form settings
        Me.Text = "Sigaluh X2PDF Converter"
        Me.Size = New Size(700, 350)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(245, 245, 245)

        Dim titleLabel As New Label With {
            .Text = "SIGALUH X2PDF CONVERTER",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(33, 33, 33),
            .Location = New Point(200, 10),
            .AutoSize = True
        }
        Me.Controls.Add(titleLabel)

        Dim yOffset = 60
        Dim labelWidth = 150
        Dim textboxWidth = 400
        Dim controlHeight = 28

        ' Folder Excel
        labelExcel = New Label() With {
            .Text = "Folder File Excel:",
            .Location = New Point(30, yOffset),
            .Size = New Size(labelWidth, controlHeight),
            .Font = New Font("Segoe UI", 10)
        }
        txtExcelFolder = New TextBox() With {
            .Location = New Point(190, yOffset),
            .Size = New Size(textboxWidth, controlHeight),
            .Font = New Font("Segoe UI", 10)
        }
        btnBrowseExcel = New Button() With {
            .Text = "Browse",
            .Location = New Point(600, yOffset),
            .Size = New Size(75, controlHeight),
            .BackColor = Color.LightSteelBlue,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        AddHandler btnBrowseExcel.Click, AddressOf OnBrowseExcel

        yOffset += 40

        ' Folder PDF
        labelPdf = New Label() With {
            .Text = "Folder Output PDF:",
            .Location = New Point(30, yOffset),
            .Size = New Size(labelWidth, controlHeight),
            .Font = New Font("Segoe UI", 10)
        }
        txtPdfFolder = New TextBox() With {
            .Location = New Point(190, yOffset),
            .Size = New Size(textboxWidth, controlHeight),
            .Font = New Font("Segoe UI", 10)
        }
        btnBrowsePdf = New Button() With {
            .Text = "Browse",
            .Location = New Point(600, yOffset),
            .Size = New Size(75, controlHeight),
            .BackColor = Color.LightSteelBlue,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        AddHandler btnBrowsePdf.Click, AddressOf OnBrowsePdf

        yOffset += 50

        ' Tombol Konversi
        btnConvert = New Button() With {
            .Text = "Konversi",
            .Location = New Point(190, yOffset),
            .Size = New Size(100, controlHeight + 5),
            .BackColor = Color.MediumSeaGreen,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        AddHandler btnConvert.Click, AddressOf OnConvertClick

        ' Tombol Setting Sheet
        Dim btnSettingSheet As New Button With {
            .Text = "Setting Sheet",
            .Location = New Point(300, yOffset),
            .Size = New Size(120, controlHeight + 5),
            .BackColor = Color.CornflowerBlue,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        AddHandler btnSettingSheet.Click, AddressOf OnSettingSheetClick

        ' Tombol Informasi
        btnInfo = New Button() With {
            .Text = "Informasi",
            .Location = New Point(430, yOffset),
            .Size = New Size(120, controlHeight + 5),
            .BackColor = Color.DarkSlateGray,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        AddHandler btnInfo.Click, AddressOf OnInfoClick

        yOffset += 50

        ' Progress Bar
        progressBar = New ProgressBar() With {
            .Location = New Point(190, yOffset),
            .Size = New Size(textboxWidth, 20)
        }

        ' Tambahkan ke Form
        Me.Controls.AddRange({labelExcel, txtExcelFolder, btnBrowseExcel,
                              labelPdf, txtPdfFolder, btnBrowsePdf,
                              btnConvert, btnSettingSheet, btnInfo,
                              progressBar})
    End Sub

    Private Sub OnSettingSheetClick(sender As Object, e As EventArgs)
        Dim sheetListPath = Path.Combine(Application.StartupPath, "SheetList.txt")

        ' Buat default jika belum ada
        If Not File.Exists(sheetListPath) Then
            Dim defaultSheets = New List(Of String) From {
            "ID", "UB", "UB RPM", "UB TIMER", "UB TACHO", "UB BPM", "UB SUHU", "PENYELIA", "LH"
        }
            File.WriteAllLines(sheetListPath, defaultSheets)
        End If

        ' Tampilkan form editor
        Dim editor As New SheetSettingForm(sheetListPath)
        editor.ShowDialog()
    End Sub


    Private Sub OnInfoClick(sender As Object, e As EventArgs)
        Dim infoText As String = "Aplikasi Sigaluh X2PDF Converter" & Environment.NewLine &
                                 "Versi     : 1.0.0" & Environment.NewLine &
                                 "Developer : Muhammad Syaoki Faradisa" & Environment.NewLine & Environment.NewLine &
                                 "Deskripsi :" & Environment.NewLine &
                                 "Aplikasi ini membantu petugas untuk mengonversi file Excel menjadi PDF secara otomatis." &
                                 "Dikembangkan untuk mempercepat proses penyeliaan dan pengunggahan dokumen ke sistem SIGALUH."

        MessageBox.Show(infoText, "Tentang Aplikasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub OnBrowseExcel(sender As Object, e As EventArgs)
        Using dialog As New FolderBrowserDialog()
            If dialog.ShowDialog() = DialogResult.OK Then
                txtExcelFolder.Text = dialog.SelectedPath
            End If
        End Using
    End Sub

    Private Sub OnBrowsePdf(sender As Object, e As EventArgs)
        Using dialog As New FolderBrowserDialog()
            If dialog.ShowDialog() = DialogResult.OK Then
                txtPdfFolder.Text = dialog.SelectedPath
            End If
        End Using
    End Sub

    Private Function SheetExists(workbook As Excel.Workbook, sheetName As String) As Boolean
        For Each sheet As Excel.Worksheet In workbook.Sheets
            If sheet.Name.Trim().ToLower() = sheetName.Trim().ToLower() Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function SafeReadCell(sheet As Excel.Worksheet, row As Integer, col As Integer, fileName As String) As String
        Try
            Dim value = CStr(sheet.Cells(row, col).Value)
            Return If(value IsNot Nothing, value.Trim(), "")
        Catch ex As Exception
            Throw New Exception($"Error membaca cell ({row}, {col}) di sheet '{sheet.Name}' dalam file '{fileName}': {ex.Message}")
        End Try
    End Function

    Private Function LoadTargetSheetsFromFile(filePath As String) As List(Of String)
        If Not File.Exists(filePath) Then
            Dim defaultSheets = New List(Of String) From {
                "ID", "UB", "UB RPM", "UB TIMER", "UB TACHO", "UB BPM", "UB SUHU", "PENYELIA", "LH"
            }
            File.WriteAllLines(filePath, defaultSheets)
            Return defaultSheets
        Else
            Return File.ReadAllLines(filePath).
                Where(Function(line) Not String.IsNullOrWhiteSpace(line)).
                Select(Function(line) line.Trim()).ToList()
        End If
    End Function

    Private Sub OnConvertClick(sender As Object, e As EventArgs)
        Dim excelFolder = txtExcelFolder.Text.Trim()
        Dim pdfFolder = txtPdfFolder.Text.Trim()
        Dim sheetListPath = Path.Combine(Application.StartupPath, "SheetList.txt")
        Dim targetSheets = LoadTargetSheetsFromFile(sheetListPath).ToArray()

        If Not Directory.Exists(excelFolder) OrElse Not Directory.Exists(pdfFolder) Then
            MessageBox.Show("Folder tidak valid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim files = Directory.GetFiles(excelFolder).Where(Function(f)
                                                              Dim ext = Path.GetExtension(f).ToLower()
                                                              Dim fileName = Path.GetFileName(f)
                                                              Return (ext = ".xls" OrElse ext = ".xlsx") AndAlso Not fileName.StartsWith("~$")
                                                          End Function).ToArray()

        If files.Length = 0 Then
            MessageBox.Show("Tidak ada file Excel ditemukan.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        progressBar.Minimum = 0
        progressBar.Maximum = files.Length
        progressBar.Value = 0

        Dim xlApp As New Excel.Application
        xlApp.Visible = False

        Try
            Dim logPath = Path.Combine(pdfFolder, "ExtractedData.txt")
            Dim errorLogPath = Path.Combine(pdfFolder, "ErrorLog.txt")
            If File.Exists(logPath) Then File.Delete(logPath)
            If File.Exists(errorLogPath) Then File.Delete(errorLogPath)

            For Each filename In files
                Dim xlWorkbook = xlApp.Workbooks.Open(filename)
                Dim tempWorkbook = xlApp.Workbooks.Add()

                Dim existingSheetsUpper = xlWorkbook.Sheets.Cast(Of Excel.Worksheet)().Select(Function(s) s.Name.ToUpper()).ToList()

                For Each sheetName In targetSheets
                    Dim sheetNameUpper = sheetName.ToUpper()
                    If existingSheetsUpper.Contains(sheetNameUpper) Then
                        For Each ws As Excel.Worksheet In xlWorkbook.Sheets
                            If ws.Name.ToUpper() = sheetNameUpper Then
                                ws.Copy(After:=tempWorkbook.Sheets(tempWorkbook.Sheets.Count))
                                Exit For
                            End If
                        Next
                    End If
                Next

                If tempWorkbook.Sheets.Count > 1 Then
                    Try
                        tempWorkbook.Sheets(1).Delete()
                    Catch
                    End Try
                    Dim outputFile = Path.Combine(pdfFolder, Path.GetFileNameWithoutExtension(filename) & ".pdf")
                    tempWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile)
                End If

                If SheetExists(xlWorkbook, "ID Sigaluh") Then
                    Try
                        Dim idSigaluhSheet As Excel.Worksheet = CType(xlWorkbook.Sheets("ID Sigaluh"), Excel.Worksheet)

                        Dim rawCertificate As String = SafeReadCell(idSigaluhSheet, 1, 3, filename)
                        Dim rawHasil = SafeReadCell(idSigaluhSheet, 9, 3, filename)

                        Dim startPos As Integer = InStr(rawHasil, "Alat yang dikalibrasi dalam batas toleransi dan dinyatakan")
                        startPos += Len("Alat yang dikalibrasi dalam batas toleransi dan dinyatakan")
                        Dim endPos As Integer = InStr(startPos, rawHasil, ",")
                        Dim result As String = If(endPos > 0, Mid(rawHasil, startPos + 1, endPos - startPos - 1), Mid(rawHasil, startPos + 1))
                        result = Trim(result)

                        Dim certificatePart As String = rawCertificate
                        If result = "TIDAK LAIK PAKAI" Then
                            If certificatePart.ToLower().StartsWith("nomor sertifikat") Then
                                certificatePart = certificatePart.Substring(certificatePart.IndexOf(":") + 1).Trim()
                            End If
                        Else
                            If certificatePart.ToLower().StartsWith("nomor surat keterangan") Then
                                certificatePart = certificatePart.Substring(certificatePart.IndexOf(":") + 1).Trim()
                            End If
                        End If

                        Dim parts = certificatePart.Split("/"c).Select(Function(p) p.Trim()).ToArray()
                        Dim alkesNumber As String = If(parts.Length > 0, parts(0), "")
                        Dim alkesOrderNumber As String = If(parts.Length > 1, parts(1), "")
                        Dim monthYearRaw As String = If(parts.Length > 2, parts(2), "")
                        Dim orderNumber As String = If(parts.Length > 3, certificatePart.Substring(certificatePart.IndexOf(parts(3))).Trim(), "")

                        Dim romawiToAngka As New Dictionary(Of String, String) From {
                            {"I", "01"}, {"II", "02"}, {"III", "03"}, {"IV", "04"},
                            {"V", "05"}, {"VI", "06"}, {"VII", "07"}, {"VIII", "08"},
                            {"IX", "09"}, {"X", "10"}, {"XI", "11"}, {"XII", "12"}
                        }

                        Dim monthPart As String = "", yearPart As String = ""
                        Dim monthYearParts = monthYearRaw.Split("-"c).Select(Function(p) p.Trim()).ToArray()
                        If monthYearParts.Length = 2 Then
                            Dim romawi = monthYearParts(0)
                            Dim tahun = "20" & monthYearParts(1)
                            If romawiToAngka.ContainsKey(romawi) Then
                                monthPart = romawiToAngka(romawi)
                                yearPart = tahun
                            End If
                        End If

                        Dim formattedMonth = If(monthPart <> "" AndAlso yearPart <> "", $"{yearPart}-{monthPart}", "")

                        Dim merek = SafeReadCell(idSigaluhSheet, 2, 3, filename)
                        Dim modelTipe = SafeReadCell(idSigaluhSheet, 3, 3, filename)
                        Dim noSeri = SafeReadCell(idSigaluhSheet, 4, 3, filename)
                        Dim tanggalKalibrasi = SafeReadCell(idSigaluhSheet, 5, 3, filename)
                        Dim tempatPengerjaan = SafeReadCell(idSigaluhSheet, 6, 3, filename)
                        Dim namaRuang = SafeReadCell(idSigaluhSheet, 7, 3, filename)
                        Dim tipe = SafeReadCell(idSigaluhSheet, 8, 3, filename)
                        Dim petugas = SafeReadCell(idSigaluhSheet, 10, 3, filename)

                        Dim posisiHasil As Integer = InStr(tipe, "HASIL")
                        Dim startPosType As Integer = posisiHasil + Len("HASIL")
                        Dim sisaString As String = Mid(tipe, startPosType + 1)
                        Dim arr() As String = Split(Trim(sisaString), " ")
                        If UBound(arr) >= 0 Then
                            tipe = arr(0)
                        End If

                        Dim dataResult As New List(Of String) From {
                            "Nomor Alkes       : " & alkesNumber,
                            "Nomor Pengerjaan  : " & alkesOrderNumber,
                            "Bulan Pengerjaan  : " & formattedMonth,
                            "Nomor Order       : " & orderNumber,
                            "Merek             : " & merek,
                            "Model/Tipe        : " & modelTipe,
                            "No. Seri          : " & noSeri,
                            "Tanggal Kalibrasi : " & tanggalKalibrasi,
                            "Tempat Pengerjaan : " & tempatPengerjaan,
                            "Nama Ruang        : " & namaRuang,
                            "Tipe              : " & tipe,
                            "Hasil             : " & result,
                            "Petugas           : " & petugas
                        }

                        File.AppendAllText(logPath, "File: " & Path.GetFileName(filename) & Environment.NewLine)
                        File.AppendAllLines(logPath, dataResult)
                        File.AppendAllText(logPath, Environment.NewLine & Environment.NewLine)

                    Catch ex As Exception
                        File.AppendAllText(errorLogPath, $"[ERROR] {Path.GetFileName(filename)} - {ex.Message}{Environment.NewLine}")
                    End Try
                End If

                tempWorkbook.Close(False)
                xlWorkbook.Close(False)

                progressBar.Value += 1
                Application.DoEvents()
            Next

            MessageBox.Show("Konversi selesai!", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Kesalahan umum: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp.Quit()
            ReleaseObject(xlApp)
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Catch
        Finally
            obj = Nothing
            GC.Collect()
        End Try
    End Sub
End Class
