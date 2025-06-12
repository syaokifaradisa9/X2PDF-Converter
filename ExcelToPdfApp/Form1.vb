Imports System.IO
Imports System.Net.Http
Imports Microsoft.Office.Interop

Public Class CalibrationData
    Public Property alkes_code As String
    Public Property calibration_order_number As String
    Public Property calibration_month As String
    Public Property order_number As String
    Public Property merk As String
    Public Property model As String
    Public Property serial_number As String
    Public Property calibration_date As String
    Public Property calibration_place As String
    Public Property room_name As String
    Public Property order_type As String
    Public Property laik_status As String
    Public Property officer As String
    Public Property source_file As String
    Public Property pdf_file As String
    Public Property processed_date As String

    Public Sub New()
        ' Constructor kosong
    End Sub
End Class

Public Class Form1
    ' Deklarasi UI (sama seperti sebelumnya)
    Private btnInfo As Button
    Private labelExcel As Label
    Private labelPdf As Label
    Private txtExcelFolder As TextBox
    Private txtPdfFolder As TextBox
    Private btnBrowseExcel As Button
    Private btnBrowsePdf As Button
    Private btnConvert As Button
    Private progressBar As ProgressBar

    ' Tambahkan HttpClient untuk API calls
    Private Shared ReadOnly httpClient As New HttpClient()

    Public Sub New()
        ' Inisialisasi komponen dasar (wajib ada)
        Me.InitializeComponent()

        ' =================================================================
        ' PENGATURAN FORM UTAMA
        ' =================================================================
        Me.Text = "X2PDF Converter v2.0"
        Me.Size = New Size(700, 420) ' Sedikit lebih tinggi untuk layout baru
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(245, 245, 245)
        ' Mengatur font default untuk semua kontrol di form ini
        Me.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))

        ' =================================================================
        ' JUDUL APLIKASI
        ' =================================================================
        Dim titleLabel As New Label With {
        .Text = "X2PDF CONVERTER",
        .Font = New Font("Segoe UI", 14.25F, FontStyle.Bold),
        .ForeColor = Color.FromArgb(64, 64, 64),
        .AutoSize = True
    }
        ' Posisi X dihitung agar selalu di tengah form
        titleLabel.Location = New Point((Me.ClientSize.Width - titleLabel.Width) / 2, 15)
        Me.Controls.Add(titleLabel)

        ' =================================================================
        ' GROUPBOX UNTUK PENGATURAN FOLDER
        ' =================================================================
        Dim grpFolders As New GroupBox() With {
        .Text = "Pengaturan Folder",
        .Location = New Point(15, 60),
        .Size = New Size(Me.ClientSize.Width - 30, 115),
        .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    }
        Me.Controls.Add(grpFolders)

        ' --- Kontrol di dalam GroupBox ---
        Dim padding As Integer = 15
        Dim controlHeight As Integer = 28
        Dim labelWidth As Integer = 150
        Dim buttonWidth As Integer = 80

        ' Folder Excel
        labelExcel = New Label() With {
        .Text = "Folder File Excel:",
        .Location = New Point(padding, 35),
        .Size = New Size(labelWidth, controlHeight),
        .TextAlign = ContentAlignment.MiddleLeft
    }
        txtExcelFolder = New TextBox() With {
        .Location = New Point(padding + labelWidth, 35),
        .Size = New Size(grpFolders.ClientSize.Width - (padding * 3) - labelWidth - buttonWidth, controlHeight),
        .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    }
        btnBrowseExcel = New Button() With {
        .Text = "Browse...",
        .Location = New Point(txtExcelFolder.Right + padding, 35),
        .Size = New Size(buttonWidth, controlHeight),
        .BackColor = Color.Gainsboro,
        .Anchor = AnchorStyles.Top Or AnchorStyles.Right
    }
        AddHandler btnBrowseExcel.Click, AddressOf OnBrowseExcel

        ' Folder PDF
        labelPdf = New Label() With {
        .Text = "Folder Output PDF:",
        .Location = New Point(padding, 75),
        .Size = New Size(labelWidth, controlHeight),
        .TextAlign = ContentAlignment.MiddleLeft
    }
        txtPdfFolder = New TextBox() With {
        .Location = New Point(padding + labelWidth, 75),
        .Size = New Size(grpFolders.ClientSize.Width - (padding * 3) - labelWidth - buttonWidth, controlHeight),
        .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    }
        btnBrowsePdf = New Button() With {
        .Text = "Browse...",
        .Location = New Point(txtPdfFolder.Right + padding, 75),
        .Size = New Size(buttonWidth, controlHeight),
        .BackColor = Color.Gainsboro,
        .Anchor = AnchorStyles.Top Or AnchorStyles.Right
    }
        AddHandler btnBrowsePdf.Click, AddressOf OnBrowsePdf

        ' Menambahkan semua kontrol folder ke dalam GroupBox
        grpFolders.Controls.AddRange({labelExcel, txtExcelFolder, btnBrowseExcel, labelPdf, txtPdfFolder, btnBrowsePdf})


        ' =================================================================
        ' TOMBOL AKSI UTAMA & PROGRESS BAR
        ' =================================================================
        ' Tombol Konversi (Aksi Utama)
        btnConvert = New Button() With {
        .Text = "Mulai Konversi & Kirim ke Sigaluh",
        .Location = New Point(15, grpFolders.Bottom + 15),
        .Size = New Size(Me.ClientSize.Width - 30, 45),
        .BackColor = Color.FromArgb(40, 167, 69), ' Warna hijau yang lebih modern
        .ForeColor = Color.White,
        .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
        .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    }
        AddHandler btnConvert.Click, AddressOf OnConvertClick
        Me.Controls.Add(btnConvert)


        ' Progress Bar
        progressBar = New ProgressBar() With {
        .Location = New Point(15, btnConvert.Bottom + 10),
        .Size = New Size(Me.ClientSize.Width - 30, 23),
        .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    }
        Me.Controls.Add(progressBar)


        ' =================================================================
        ' TOMBOL AKSI SEKUNDER
        ' =================================================================
        ' Tombol Informasi
        btnInfo = New Button() With {
        .Text = "Informasi",
        .Size = New Size(120, 33),
        .BackColor = Color.DarkSlateGray,
        .ForeColor = Color.White,
        .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    }
        btnInfo.Location = New Point(Me.ClientSize.Width - btnInfo.Width - 15, Me.ClientSize.Height - btnInfo.Height - 15)
        AddHandler btnInfo.Click, AddressOf OnInfoClick
        Me.Controls.Add(btnInfo)

        ' Tombol Setting Sheet
        Dim btnSettingSheet As New Button With {
        .Text = "Setting Sheet",
        .Size = New Size(120, 33),
        .BackColor = Color.CornflowerBlue,
        .ForeColor = Color.White,
        .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    }
        btnSettingSheet.Location = New Point(btnInfo.Left - btnSettingSheet.Width - 10, btnInfo.Top)
        AddHandler btnSettingSheet.Click, AddressOf OnSettingSheetClick
        Me.Controls.Add(btnSettingSheet)


        ' =================================================================
        ' SETUP AWAL & KONFIGURASI API
        ' =================================================================
        ' Dapatkan path absolut
        Dim sourcePath As String = Path.GetFullPath(Path.Combine(Application.StartupPath, "..", "Source"))
        Dim targetPath As String = Path.GetFullPath(Path.Combine(Application.StartupPath, "..", "Target"))

        ' Buat folder jika belum ada
        Directory.CreateDirectory(sourcePath)
        Directory.CreateDirectory(targetPath)

        ' Set nilai textbox
        txtExcelFolder.Text = sourcePath
        txtPdfFolder.Text = targetPath

        ' Konfigurasi API Client
        httpClient.DefaultRequestHeaders.Add("X-CONVERTER-API-KEY", "xgr8xX2Ee0ZDUAWOmQULQhgprd9udQvrHFtbfn0Ep7kif8HYtcOXDgvmcve6bDma")
    End Sub

    ' Method alternatif jika ingin mengirim sebagai form data biasa
    Private Async Function SendDataToAPI(dataObj As CalibrationData, pdfFilePath As String) As Task(Of Boolean)
        Try
            Dim apiUrl As String = "http://127.0.0.1:8000/api/v1/revision/send"

            Using content As New MultipartFormDataContent()
                content.Add(New StringContent(If(dataObj.alkes_code, "")), "alkes_code")
                content.Add(New StringContent(If(dataObj.calibration_order_number, "")), "calibration_order_number")
                content.Add(New StringContent(If(dataObj.calibration_month, "")), "calibration_month")
                content.Add(New StringContent(If(dataObj.order_number, "")), "order_number")
                content.Add(New StringContent(If(dataObj.merk, "")), "merk")
                content.Add(New StringContent(If(dataObj.model, "")), "model")
                content.Add(New StringContent(If(dataObj.serial_number, "")), "serial_number")
                content.Add(New StringContent(If(dataObj.calibration_date, "")), "calibration_date")
                content.Add(New StringContent(If(dataObj.calibration_place, "")), "calibration_place")
                content.Add(New StringContent(If(dataObj.room_name, "")), "room_name")
                content.Add(New StringContent(If(dataObj.order_type, "")), "order_type")
                content.Add(New StringContent(If(dataObj.laik_status, "")), "laik_status")
                content.Add(New StringContent(If(dataObj.officer, "")), "officer")
                content.Add(New StringContent(If(dataObj.source_file, "")), "source_file")
                content.Add(New StringContent(If(dataObj.pdf_file, "")), "pdf_file")
                content.Add(New StringContent(If(dataObj.processed_date, "")), "processed_date")

                If File.Exists(pdfFilePath) Then
                    Dim fileBytes As Byte() = File.ReadAllBytes(pdfFilePath)
                    Dim fileContent As New ByteArrayContent(fileBytes)
                    fileContent.Headers.ContentType = New System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf")
                    content.Add(fileContent, "pdf_file_upload", Path.GetFileName(pdfFilePath))
                End If


                ' Kirim request
                Dim response As HttpResponseMessage = Await httpClient.PostAsync(apiUrl, content)

                If response.IsSuccessStatusCode Then
                    Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                    Console.WriteLine($"API Response: {responseContent}")
                    Return True
                Else
                    Console.WriteLine($"API Error: {response.StatusCode} - {response.ReasonPhrase}")
                    Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                    Console.WriteLine($"Error Content: {errorContent}")
                    Return False
                End If
            End Using

        Catch ex As Exception
            Console.WriteLine($"Error sending to API: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ParseIndonesianDate(ByVal indonesianDate As String) As String
        ' Jika input kosong, kembalikan string kosong
        If String.IsNullOrWhiteSpace(indonesianDate) Then
            Return ""
        End If

        Try
            ' Buat objek CultureInfo untuk Bahasa Indonesia (id-ID).
            ' Ini yang akan membuat .NET mengenali "Agustus", "Januari", dll.
            Dim culture As New System.Globalization.CultureInfo("id-ID")

            ' 1. Ubah string menjadi objek DateTime yang sebenarnya.
            Dim dt As DateTime = DateTime.Parse(indonesianDate, culture)

            ' 2. Format objek DateTime tersebut ke format "yyyy-MM-dd".
            Return dt.ToString("yyyy-MM-dd")

        Catch ex As Exception
            ' Jika terjadi error (misal format tidak dikenali), catat error dan kembalikan string kosong.
            Console.WriteLine($"Gagal mengurai tanggal '{indonesianDate}': {ex.Message}")
            Return ""
        End Try
    End Function

    ' Modifikasi method OnConvertClick untuk menambahkan pengiriman ke API
    Private Async Sub OnConvertClick(sender As Object, e As EventArgs)
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

        ' Disable tombol convert selama proses
        btnConvert.Enabled = False
        btnConvert.Text = "Proses Konversi"

        Dim xlApp As New Excel.Application
        xlApp.Visible = False

        Try
            Dim logPath = Path.Combine(pdfFolder, "ExtractedData.txt")
            Dim errorLogPath = Path.Combine(pdfFolder, "ErrorLog.txt")
            Dim apiLogPath = Path.Combine(pdfFolder, "APILog.txt")

            If File.Exists(logPath) Then File.Delete(logPath)
            If File.Exists(errorLogPath) Then File.Delete(errorLogPath)
            If File.Exists(apiLogPath) Then File.Delete(apiLogPath)

            Dim successCount As Integer = 0
            Dim apiSuccessCount As Integer = 0

            For Each filename In files
                Dim xlWorkbook = xlApp.Workbooks.Open(filename)
                Dim tempWorkbook = xlApp.Workbooks.Add()
                Dim outputFile As String = ""
                Dim dataObj As CalibrationData = Nothing

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
                    outputFile = Path.Combine(pdfFolder, Path.GetFileNameWithoutExtension(filename) & ".pdf")
                    tempWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile)
                End If

                If SheetExists(xlWorkbook, "ID Sigaluh") Then
                    Try
                        Dim idSigaluhSheet As Excel.Worksheet = CType(xlWorkbook.Sheets("ID Sigaluh"), Excel.Worksheet)

                        ' ... (kode ekstraksi data sama seperti sebelumnya)
                        Dim rawCertificate As String = SafeReadCell(idSigaluhSheet, 1, 3, filename)
                        Dim rawHasil = SafeReadCell(idSigaluhSheet, 9, 3, filename)

                        Dim startPos As Integer = InStr(rawHasil, "Alat yang dikalibrasi dalam batas toleransi dan dinyatakan")
                        startPos += Len("Alat yang dikalibrasi dalam batas toleransi dan dinyatakan")
                        Dim endPos As Integer = InStr(startPos, rawHasil, ",")
                        Dim result As String = If(endPos > 0, Mid(rawHasil, startPos + 1, endPos - startPos - 1), Mid(rawHasil, startPos + 1))
                        result = Trim(result)

                        result = "LAIK PAKAI"
                        If (rawHasil.Contains("TIDAK LAIK PAKAI")) Then
                            result = "TIDAK LAIK PAKAI"
                        End If


                        Dim certificatePart As String = rawCertificate
                        If result = "TIDAK LAIK PAKAI" Then
                            If certificatePart.ToLower().StartsWith("nomor surat keterangan") Then
                                certificatePart = certificatePart.Substring(certificatePart.IndexOf(":") + 1).Trim()
                            End If
                        Else
                            If certificatePart.ToLower().StartsWith("nomor sertifikat") Then
                                certificatePart = certificatePart.Substring(certificatePart.IndexOf(":") + 1).Trim()
                            End If
                        End If

                        Dim parts = certificatePart.Split("/"c).Select(Function(p) p.Trim()).ToArray()
                        Dim alkesNumber As String = If(parts.Length > 0, parts(0), "")
                        Dim alkesOrderNumber As String = If(parts.Length > 1, String.Concat(parts(1).Where(AddressOf Char.IsDigit)), "")
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
                        tanggalKalibrasi = ParseIndonesianDate(tanggalKalibrasi)

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

                        ' Buat dataObj untuk dikirim ke API
                        dataObj = New CalibrationData() With {
                            .alkes_code = If(alkesNumber, ""),
                            .calibration_order_number = If(alkesOrderNumber, ""),
                            .calibration_month = If(formattedMonth, ""),
                            .order_number = If(orderNumber, ""),
                            .merk = If(merek, ""),
                            .model = If(modelTipe, ""),
                            .serial_number = If(noSeri, ""),
                            .calibration_date = If(tanggalKalibrasi, ""),
                            .calibration_place = If(tempatPengerjaan, ""),
                            .room_name = If(namaRuang, ""),
                            .order_type = If(tipe, ""),
                            .laik_status = If(result, ""),
                            .officer = If(petugas, ""),
                            .source_file = If(Path.GetFileName(filename), ""),
                            .pdf_file = If(File.Exists(outputFile), Path.GetFileName(outputFile), ""),
                            .processed_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        }

                        successCount += 1

                    Catch ex As Exception
                        File.AppendAllText(errorLogPath, $"[ERROR] {Path.GetFileName(filename)} - {ex.Message}{Environment.NewLine}")
                    End Try
                End If

                ' Kirim data ke API jika ada dataObj dan file PDF
                Try
                    ' Update status di UI
                    btnConvert.Text = $"Upload Ke Sigaluh ({progressBar.Value + 1}/{files.Length})"
                    Application.DoEvents()

                    ' Kirim ke API (gunakan await)
                    Dim apiSuccess As Boolean = Await SendDataToAPI(dataObj, outputFile)

                    If apiSuccess Then
                        apiSuccessCount += 1
                        File.AppendAllText(apiLogPath, $"[SUCCESS] {Path.GetFileName(filename)} - Data dan PDF berhasil dikirim ke API{Environment.NewLine}")
                    Else
                        File.AppendAllText(apiLogPath, $"[FAILED] {Path.GetFileName(filename)} - Gagal mengirim ke API{Environment.NewLine}")
                    End If

                Catch apiEx As Exception
                    File.AppendAllText(apiLogPath, $"[ERROR] {Path.GetFileName(filename)} - API Error: {apiEx.Message}{Environment.NewLine}")
                End Try

                ' Clean up workbooks
                tempWorkbook.Close(False)
                xlWorkbook.Close(False)

                progressBar.Value += 1
                Application.DoEvents()
            Next

            ' Tampilkan hasil akhir
            Dim resultMessage As String = $"Konversi selesai!" & Environment.NewLine & Environment.NewLine &
                                        $"Total file diproses        : {files.Length}" & Environment.NewLine &
                                        $"Berhasil ekstrak data PDf  : {successCount}" & Environment.NewLine &
                                        $"Berhasil kirim ke Sigaluh  : {apiSuccessCount}"

            MessageBox.Show(resultMessage, "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Kesalahan umum: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Reset UI
            btnConvert.Enabled = True
            btnConvert.Text = "Konversi"

            ' Cleanup Excel
            xlApp.Quit()
            ReleaseObject(xlApp)
        End Try
    End Sub

    ' Tambahkan method untuk validasi data sebelum kirim (opsional)
    Private Function ValidateDataBeforeSending(dataObj As Object) As Boolean
        Try
            ' Validasi field-field penting
            Dim objType = dataObj.GetType()
            Dim alkesCode = objType.GetProperty("alkes_code")?.GetValue(dataObj)?.ToString()
            Dim serialNumber = objType.GetProperty("serial_number")?.GetValue(dataObj)?.ToString()

            ' Minimal harus ada alkes_code
            If String.IsNullOrWhiteSpace(alkesCode) Then
                Return False
            End If

            Return True
        Catch
            Return False
        End Try
    End Function

    ' Method untuk log detailed API response (opsional)
    Private Sub LogAPIResponse(fileName As String, response As String, isSuccess As Boolean)
        Try
            Dim logPath = Path.Combine(txtPdfFolder.Text.Trim(), "DetailedAPILog.txt")
            Dim logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {fileName} - {If(isSuccess, "SUCCESS", "FAILED")}" & Environment.NewLine &
                          $"Response: {response}" & Environment.NewLine &
                          "==========================================" & Environment.NewLine

            File.AppendAllText(logPath, logEntry)
        Catch
            ' Silent fail untuk logging
        End Try
    End Sub

    ' Sisanya method tetap sama seperti sebelumnya
    Private Sub OnSettingSheetClick(sender As Object, e As EventArgs)
        Dim sheetListPath = Path.Combine(Application.StartupPath, "SheetList.txt")

        ' Buat default jika belum ada
        If Not File.Exists(sheetListPath) Then
            Dim defaultSheets = New List(Of String) From {
                "Sheet1"
            }
            File.WriteAllLines(sheetListPath, defaultSheets)
        End If

        ' Tampilkan form editor
        Dim editor As New SheetSettingForm(sheetListPath)
        editor.ShowDialog()
    End Sub

    Private Sub OnInfoClick(sender As Object, e As EventArgs)
        Dim infoText As String = "Aplikasi Sigaluh X2PDF Converter" & Environment.NewLine &
                                 "Versi     : 1.1.0" & Environment.NewLine &
                                 "Developer : Muhammad Syaoki Faradisa" & Environment.NewLine & Environment.NewLine &
                                 "Deskripsi :" & Environment.NewLine &
                                 "Aplikasi ini membantu untuk mengonversi file Excel menjadi PDF secara otomatis dan mengirim data hasil ekstraksi ke Sistem Sigaluh."

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
                "Sheet1"
            }
            File.WriteAllLines(filePath, defaultSheets)
            Return defaultSheets
        Else
            Return File.ReadAllLines(filePath).
                Where(Function(line) Not String.IsNullOrWhiteSpace(line)).
                Select(Function(line) line.Trim()).ToList()
        End If
    End Function

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    ' Destructor untuk cleanup HttpClient
    Protected Overrides Sub Finalize()
        httpClient?.Dispose()
        MyBase.Finalize()
    End Sub

End Class
