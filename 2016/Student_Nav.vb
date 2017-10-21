Imports System.IO
Imports System.Data.SqlClient
Public Class Student_Nav
    Dim db As New Database
    Dim dbcomm As New SqlCommand
    Dim sql As String
    Dim dbread As SqlDataReader
    Dim fileLocDownload As String
    Dim filelocUpload As String
    Dim jamSisa As Double
    Dim ms As New MemoryStream
    Dim fs As FileStream
    Dim deadline As String
    Public idsiswa As String
    Private Sub Student_Nav_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed

        db.conn()
        jadwalUjian()
        If lblTimer.Text = "0" Then
            doSubmit()
        End If
    End Sub

    Private Sub jadwalUjian()
        sql = "select * from Siswa as s join AktivitasUjian as au on s.kode_pendaftaran=au.kode_pendaftaran join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian where au.kode_pendaftaran = '" & idsiswa & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                cbUjian.Items.Add(dbread("kode_ujian") & " - " & dbread("tanggal_ujian") & " - " & dbread("durasi_ujian") & " Menit")
                jamSisa = dbread("durasi_ujian") * 60
                deadline = dbread("tanggal_ujian").toshortdatestring & " " & DateTime.Now.AddMinutes(dbread("durasi_ujian")).ToString("hh:mm:ss")
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub downloaddSoal()
        sql = "select * from Siswa as s join AktivitasUjian as au on s.kode_pendaftaran=au.kode_pendaftaran join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian where au.kode_pendaftaran = '" & idsiswa & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            If dbread.HasRows Then
                Dim x As Byte() = CType(dbread("soal_ujian"), Byte()).ToArray()
                Dim fs As New FileStream(fileLocDownload, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                fs.Write(x, 0, x.Length)
                fs.Close()
                MsgBox("Download success", MsgBoxStyle.Information)
                Process.Start(fileLocDownload)
                Timer1.Start()
            Else
                MsgBox("Soal belum di Upload", MsgBoxStyle.Exclamation)
            End If
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub lblDownloadSoal_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblDownloadSoal.LinkClicked
        If cbUjian.Text = "Pilih" Then
            MsgBox("Select exam schedule first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            fileLocDownload = SaveFileDialog1.FileName
            downloaddSoal()
        End If

    End Sub

    Private Sub btnUploadJawaban_Click(sender As Object, e As EventArgs) Handles btnUploadJawaban.Click
        If DataGridViewJawaban.Rows.Count = 3 Then
            MsgBox("Max 3 times to upload answer!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        OpenFileDialog1.Reset()
        OpenFileDialog1.Filter = "DOCX Files(*.docx)|*.docx|PDF Files(*.pdf)|*.pdf"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            txtUploadSoal.Text = OpenFileDialog1.FileName
            filelocUpload = OpenFileDialog1.FileName
            DataGridViewJawaban.Rows.Add(filelocUpload.Substring(filelocUpload.LastIndexOf("\") + 1), DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"))
            fs = New FileStream(filelocUpload, FileMode.Open, FileAccess.Read)
            fs.CopyTo(ms)
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If jamSisa > 0 Then
            jamSisa = jamSisa - 1
            Dim waktu As TimeSpan = TimeSpan.FromSeconds(jamSisa)
            lblTimer.Text = waktu.Hours.ToString & ":" & waktu.Minutes.ToString & ":" & waktu.Seconds.ToString
        End If
    End Sub

    Private Sub btnFInalize_Click(sender As Object, e As EventArgs) Handles btnFInalize.Click
        MsgBox(DataGridViewJawaban.Item(1, 0).Value)
        doSubmit()
    End Sub

    Private Sub doSubmit()
        If DataGridViewJawaban.Rows.Count = 1 Then
            db.uploadJawaban1(DataGridViewJawaban.Item(1, 0).Value, deadline, ms.ToArray, idsiswa)
            MsgBox("Answer Uploaded", MsgBoxStyle.Information)
            resetAll()
        ElseIf DataGridViewJawaban.Rows.Count = 2 Then
            db.uploadJawaban2(DataGridViewJawaban.Item(1, 0).Value, DataGridViewJawaban.Item(1, 1).Value, deadline, ms.ToArray, idsiswa)
            MsgBox("Answer Uploaded", MsgBoxStyle.Information)
            resetAll()
        ElseIf DataGridViewJawaban.Rows.Count = 3 Then
            db.uploadJawaban3(DataGridViewJawaban.Item(1, 0).Value, DataGridViewJawaban.Item(1, 1).Value, DataGridViewJawaban.Item(1, 2).Value, deadline, ms.ToArray, idsiswa)
            MsgBox("Answer Uploaded", MsgBoxStyle.Information)
            resetAll()
        ElseIf DataGridViewJawaban.Rows.Count = 0
            MsgBox("Loser!", MsgBoxStyle.Exclamation)
            resetAll()
        End If
    End Sub

    Private Sub resetAll()
        cbUjian.SelectedIndex = -1
        cbUjian.Text = "Pilih"
        txtUploadSoal.Text = ""
        DataGridViewJawaban.Rows.Clear()
        lblTimer.Text = ""
        ms.Close()
        fs.Close()
    End Sub

    Private Sub hasilUjian()
        Chart1.Series.Clear()
        Chart1.Series.Add("Nilai")
        sql = "select * from Siswa as s join AktivitasUjian as au on s.kode_pendaftaran=au.kode_pendaftaran join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian where au.kode_pendaftaran = 'S0001' and ju.kode_ujian='" & txtKode.Text & "' and ju.tanggal_ujian='" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                Chart1.Series("Nilai").Points.AddXY("Nilai PG", dbread("nilai_pg").ToString)
                Chart1.Series("Nilai").Points.AddXY("Nilai Essay", dbread("nilai_essay").ToString)
                Chart1.Series("Nilai").Points.AddXY("Nilai Kasus", dbread("nilai_kasus").ToString)

                Dim totalScore As String = (30% * dbread("nilai_pg")) + (30% * dbread("nilai_essay")) + (40% * dbread("nilai_kasus"))
                If totalScore.Substring(0, 2) >= 65 Then
                    lblTotal.Text = totalScore.Substring(0, 2)
                    lblHasil.Text = "Lulus"
                ElseIf totalScore.Substring(0, 2) < 65 Then
                    lblTotal.Text = totalScore.Substring(0, 2)
                    lblHasil.Text = "Tidak Lulus"
                End If
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        hasilUjian()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectTab(0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TabControl1.SelectTab(1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Main_Frm.Show()
        Me.Close()
    End Sub
End Class