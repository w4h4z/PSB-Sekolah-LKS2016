Imports System.IO
Imports System.Data.SqlClient
Public Class Staff_Nav
    Dim db As New Database
    Dim locJurusan As String
    Dim dbcomm As New SqlCommand
    Dim sql As String
    Dim dbread As SqlDataReader
    Public idstaff As String
    Private Sub Staff_Nav_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the '_2016DataSet1.Karyawan' table. You can move, or remove it, as needed.
        Me.KaryawanTableAdapter.Fill(Me._2016DataSet1.Karyawan)
        'TODO: This line of code loads data into the '_2016DataSet1.Kelas' table. You can move, or remove it, as needed.
        Me.KelasTableAdapter.Fill(Me._2016DataSet1.Kelas)
        'TODO: This line of code loads data into the '_2016DataSet.UserAccount' table. You can move, or remove it, as needed.
        Me.UserAccountTableAdapter.Fill(Me._2016DataSet.UserAccount)
        'TODO: This line of code loads data into the '_2016DataSet.Jurusan' table. You can move, or remove it, as needed.
        Me.JurusanTableAdapter.Fill(Me._2016DataSet.Jurusan)

        TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed

        db.conn()
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Function incrementKode(kode)
        Dim kodeID As String
        Dim huruf As String = kode.Substring(0, 1)
        Dim id As String = kode.Substring(1)
        Dim angka As Integer = Integer.Parse(id)
        angka += 1
        kode = huruf & angka.ToString("D" & id.Length)
        kodeID = huruf & angka.ToString("D" & id.Length)
        Return kodeID
    End Function

#Region "Side Bar"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectTab(0)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TabControl1.SelectTab(1)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TabControl1.SelectTab(2)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TabControl1.SelectTab(3)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TabControl1.SelectTab(4)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TabControl1.SelectTab(5)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TabControl1.SelectTab(6)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TabControl1.SelectTab(7)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TabControl1.SelectTab(8)
        kodeJurusan()
        kodeGuru()
        kodeJadwal()
        dataGuru()
        dataSiswa()
        dataAkun()
        dataJadwalUjian()
        jadwalUjian()
        siswaNull()
        siswaNotNull()
        dataPassed()
        kodeLaporan()
        cariGuru()
        cariSiswa()
        cariJadwal()
        cariKoreksi()
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Login_Frm.Show()
        Me.Close()
    End Sub
#End Region

#Region "Jurusan"
    Private Function kodeJurusan()
        Dim lastkode As String = db.maxKode("kode_jurusan", "Jurusan")
        Dim kd As String = incrementKode(lastkode)
        txtKodeJurusan.Text = kd
        Return kd
    End Function

    Private Sub btnImageJurusan_Click(sender As Object, e As EventArgs) Handles btnImageJurusan.Click
        OpenFileDialog1.Filter = "JPEG Files(*.jpeg)|*.jpeg|PNG Files(*.png)|*.png|JPG Files(*.jpg)|*.jpg"
        OpenFileDialog1.Reset()
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            locJurusan = OpenFileDialog1.FileName
            PictureBoxJurusan.ImageLocation = locJurusan
        End If
    End Sub

    Private Sub btnResetJurusan_Click(sender As Object, e As EventArgs) Handles btnResetJurusan.Click
        resetJurusan()
    End Sub

    Private Sub resetJurusan()
        txtNamaJurusan.Text = ""
        txtDescJurusan.Text = ""
        PictureBoxJurusan.Image = My.Resources.blank
        kodeJurusan()
    End Sub

    Private Sub btnAddJurusan_Click(sender As Object, e As EventArgs) Handles btnAddJurusan.Click
        If txtNamaJurusan.Text = "" Or txtDescJurusan.Text = "" Or locJurusan = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim ms As MemoryStream = New MemoryStream()
        Dim fs As FileStream = New FileStream(locJurusan, FileMode.Open, FileAccess.Read)
        fs.CopyTo(ms)

        db.insertJurusan(txtKodeJurusan.Text, txtNamaJurusan.Text, txtDescJurusan.Text, ms.ToArray)
        MsgBox("Insert data success", MsgBoxStyle.Information)

        fs.Close()
        ms.Close()
        resetJurusan()
        Me.JurusanTableAdapter.Fill(Me._2016DataSet.Jurusan)
    End Sub

    Private Sub DataGridViewJurusan_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewJurusan.CellMouseClick
        Dim i As Integer = DataGridViewJurusan.CurrentRow.Index

        txtKodeJurusan.Text = DataGridViewJurusan.Item(0, i).Value
        txtNamaJurusan.Text = DataGridViewJurusan.Item(1, i).Value
        txtDescJurusan.Text = DataGridViewJurusan.Item(2, i).Value

        Try
            Dim gambar As Byte() = DataGridViewJurusan.Item(3, i).Value
            Dim ms As New MemoryStream(gambar)
            PictureBoxJurusan.Image = Image.FromStream(ms)
            ms.Close()
        Catch ex As Exception
            PictureBoxJurusan.Image = My.Resources.blank
        End Try
    End Sub

    Private Sub btnEditJurusan_Click(sender As Object, e As EventArgs) Handles btnEditJurusan.Click
        btnSaveJurusan.Enabled = True
    End Sub

    Private Sub btnDeleteJurusan_Click(sender As Object, e As EventArgs) Handles btnDeleteJurusan.Click
        db.deleteJurusan(txtKodeJurusan.Text)
        MsgBox("Delete data success", MsgBoxStyle.Information)
        Me.JurusanTableAdapter.Fill(Me._2016DataSet.Jurusan)
        resetJurusan()
    End Sub

    Private Sub btnSaveJurusan_Click(sender As Object, e As EventArgs) Handles btnSaveJurusan.Click
        If txtNamaJurusan.Text = "" Or txtDescJurusan.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If locJurusan = "" Then
            db.updateJurusan(txtKodeJurusan.Text, txtNamaJurusan.Text, txtDescJurusan.Text)
        ElseIf Not locJurusan = ""
            Dim ms As MemoryStream = New MemoryStream()
            Dim fs As FileStream = New FileStream(locJurusan, FileMode.Open, FileAccess.Read)
            fs.CopyTo(ms)

            db.updateJurusanImage(txtKodeJurusan.Text, txtNamaJurusan.Text, txtDescJurusan.Text, ms.ToArray)

            ms.Close()
            fs.Close()
        End If

        MsgBox("Update data success", MsgBoxStyle.Information)


        resetJurusan()
        Me.JurusanTableAdapter.Fill(Me._2016DataSet.Jurusan)
        btnSaveJurusan.Enabled = False
    End Sub

#End Region

#Region "Karyawan dan Guru"
    Private Function kodeGuru()
        Dim lastkode As String = db.maxKode("kode_karyawan", "Karyawan")
        Dim kd As String = incrementKode(lastkode)
        txtKodeGuru.Text = kd
        Return kd
    End Function

    Private Sub cariGuru()
        cbCariGuru.DisplayMember = "Text"
        cbCariGuru.ValueMember = "Key"
        cbCariGuru.Items.Add(New With {Key .Text = "Kode Karyawan", Key .Value = "kode_karyawan"})
        cbCariGuru.Items.Add(New With {Key .Text = "NIP", Key .Value = "NIP"})
        cbCariGuru.Items.Add(New With {Key .Text = "Nama Karyawan", Key .Value = "nama_karyawan"})
        cbCariGuru.Items.Add(New With {Key .Text = "Tanggal Lahir Karyawan", Key .Value = "tgl_lhr_karyawan"})
        cbCariGuru.Items.Add(New With {Key .Text = "Alamat Karyawan", Key .Value = "alamat_karyawan"})
        cbCariGuru.Items.Add(New With {Key .Text = "Jenis Kelamin Karyawan", Key .Value = "jk_karyawan"})
        cbCariGuru.Items.Add(New With {Key .Text = "No Telp Karyawan", Key .Value = "no_telp_karyawan"})

        cbSortGuru.DisplayMember = "Text"
        cbSortGuru.ValueMember = "Key"
        cbSortGuru.Items.Add(New With {Key .Text = "Kode Karyawan", Key .Value = "kode_karyawan"})
        cbSortGuru.Items.Add(New With {Key .Text = "NIP", Key .Value = "NIP"})
        cbSortGuru.Items.Add(New With {Key .Text = "Nama Karyawan", Key .Value = "nama_karyawan"})
        cbSortGuru.Items.Add(New With {Key .Text = "Tanggal Lahir Karyawan", Key .Value = "tgl_lhr_karyawan"})
        cbSortGuru.Items.Add(New With {Key .Text = "Alamat Karyawan", Key .Value = "alamat_karyawan"})
        cbSortGuru.Items.Add(New With {Key .Text = "Jenis Kelamin Karyawan", Key .Value = "jk_karyawan"})
        cbSortGuru.Items.Add(New With {Key .Text = "No Telp Karyawan", Key .Value = "no_telp_karyawan"})
    End Sub

    Private Sub searchGuru(kolom, rows)
        DataGridViewGuru.Rows.Clear()
        sql = "select * from Karyawan where " & kolom & " like '%" & rows & "%'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewGuru.Rows.Add(dbread("kode_karyawan"), dbread("NIP"), dbread("nama_karyawan"), dbread("tgl_lhr_karyawan").toshortdatestring, dbread("alamat_karyawan"), dbread("jk_karyawan"), dbread("no_telp_karyawan"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub sortGuru(kolom, by)
        DataGridViewGuru.Rows.Clear()
        sql = "select * from Karyawan order by " & kolom & " " & by & ""

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewGuru.Rows.Add(dbread("kode_karyawan"), dbread("NIP"), dbread("nama_karyawan"), dbread("tgl_lhr_karyawan").toshortdatestring, dbread("alamat_karyawan"), dbread("jk_karyawan"), dbread("no_telp_karyawan"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnSortGuru_Click(sender As Object, e As EventArgs) Handles btnSortGuru.Click
        If cbSortGuru.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim by As String
        If rbAscGuru.Checked = True Then
            by = "ASC"
        ElseIf rbDescGuru.Checked = True Then
            by = "DESC"
        End If
        Try
            sortGuru(cbSortGuru.SelectedItem.Value, by)
        Catch ex As Exception
            MsgBox("Tipe data ini tidak dapat di sort!", MsgBoxStyle.Exclamation)
        End Try

    End Sub

    Private Sub btnCariGuru_Click(sender As Object, e As EventArgs) Handles btnCariGuru.Click
        If cbCariGuru.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        searchGuru(cbCariGuru.SelectedItem.Value, txtCariGuru.Text)
    End Sub

    Private Sub dataGuru()
        DataGridViewGuru.Rows.Clear()
        sql = "select * from karyawan"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewGuru.Rows.Add(dbread("kode_karyawan"), dbread("NIP"), dbread("nama_karyawan"), dbread("tgl_lhr_karyawan").toshortdatestring, dbread("alamat_karyawan"), dbread("jk_karyawan"), dbread("no_telp_karyawan"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub resetGuru()
        txtNipGuru.Text = ""
        txtNamaGuru.Text = ""
        DateTimePickerGuru.Value = Date.Now
        txtAlamatGuru.Text = ""
        rbLkGuru.Checked = False
        rbPrGuru.Checked = False
        txtPhoneGuru.Text = ""
        kodeGuru()
    End Sub

    Private Sub btnAddGuru_Click(sender As Object, e As EventArgs) Handles btnAddGuru.Click
        Dim jk As String
        Dim b As Boolean
        If rbLkGuru.Checked = True Then
            jk = "Laki-Laki"
            b = True
        ElseIf rbPrGuru.Checked = True
            jk = "Perempuan"
            b = True
        End If

        If rbLkGuru.Checked = False And rbPrGuru.Checked = False Then
            b = False
        End If

        If txtNipGuru.Text = "" Or txtNamaGuru.Text = "" Or txtAlamatGuru.Text = "" Or b = False Or txtPhoneGuru.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        db.insertGuru(txtKodeGuru.Text, txtNipGuru.Text, txtNamaGuru.Text, DateTimePickerGuru.Value.ToString("yyyy/MM/dd"), txtAlamatGuru.Text, jk, txtPhoneGuru.Text)
        db.insertAkun("kode_karyawan", txtKodeGuru.Text, txtKodeGuru.Text + DateTimePickerGuru.Value.ToString("MMddyy"), 2, txtKodeGuru.Text)

        MsgBox("Insert data success", MsgBoxStyle.Information)
        dataGuru()
        resetGuru()
    End Sub

    Private Sub txtPhoneGuru_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPhoneGuru.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub btnResetGuru_Click(sender As Object, e As EventArgs) Handles btnResetGuru.Click
        resetGuru()
        kodeGuru()
    End Sub

    Private Sub DataGridViewGuru_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewGuru.CellMouseClick
        Dim i As Integer = DataGridViewGuru.CurrentRow.Index

        txtKodeGuru.Text = DataGridViewGuru.Item(0, i).Value
        txtNipGuru.Text = DataGridViewGuru.Item(1, i).Value
        txtNamaGuru.Text = DataGridViewGuru.Item(2, i).Value
        DateTimePickerGuru.Value = DataGridViewGuru.Item(3, i).Value
        txtAlamatGuru.Text = DataGridViewGuru.Item(4, i).Value
        If DataGridViewGuru.Item(5, i).Value = "Laki-Laki" Then
            rbLkGuru.Checked = True
        ElseIf DataGridViewGuru.Item(5, i).Value = "Perempuan"
            rbPrGuru.Checked = True
        End If
        txtPhoneGuru.Text = DataGridViewGuru.Item(6, i).Value
    End Sub

    Private Sub btnEditGuru_Click(sender As Object, e As EventArgs) Handles btnEditGuru.Click
        btnSaveGuru.Enabled = True
    End Sub

    Private Sub btnDeleteGuru_Click(sender As Object, e As EventArgs) Handles btnDeleteGuru.Click
        db.deleteGuru(txtKodeGuru.Text)
        MsgBox("Delete data success", MsgBoxStyle.Information)
        resetGuru()
        dataGuru()
    End Sub

    Private Sub btnSaveGuru_Click(sender As Object, e As EventArgs) Handles btnSaveGuru.Click
        Dim jk As String

        If rbLkGuru.Checked = True Then
            jk = "Laki-Laki"
        ElseIf rbPrGuru.Checked = True
            jk = "Perempuan"
        End If

        If txtNipGuru.Text = "" Or txtNamaGuru.Text = "" Or txtAlamatGuru.Text = "" Or txtPhoneGuru.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        db.updateGuru(txtKodeGuru.Text, txtNipGuru.Text, txtNamaGuru.Text, DateTimePickerGuru.Value.ToString("yyyy/MM/dd"), txtAlamatGuru.Text, jk, txtPhoneGuru.Text)

        MsgBox("Update data success", MsgBoxStyle.Information)
        dataGuru()
        resetGuru()
        btnSaveGuru.Enabled = False
    End Sub

#End Region

#Region "Siswa"
    Private Sub dataSiswa()
        DataGridViewSiswa.Rows.Clear()
        sql = "SELECT s.kode_pendaftaran, s.nama_lengkap, s.alamat_siswa, s.tgl_lhr_siswa, s.no_telp_siswa, s.jk_siswa, s.asal_sekolah, s.prioritas_jurusan_1, s.prioritas_jurusan_2, s.prioritas_jurusan_3, j.kode_jurusan, j.nama_jurusan AS p1, ja.nama_jurusan AS p2, jb.nama_jurusan AS p3 FROM dbo.Siswa AS s INNER JOIN dbo.Jurusan AS j ON s.prioritas_jurusan_1 = j.kode_jurusan INNER JOIN dbo.Jurusan AS ja ON s.prioritas_jurusan_2 = ja.kode_jurusan INNER JOIN dbo.Jurusan AS jb ON s.prioritas_jurusan_3 = jb.kode_jurusan"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewSiswa.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"), dbread("alamat_siswa"), dbread("tgl_lhr_siswa").toshortdatestring, dbread("no_telp_siswa"), dbread("jk_siswa"), dbread("asal_sekolah"), dbread("prioritas_jurusan_1"), dbread("p1"), dbread("prioritas_jurusan_2"), dbread("p2"), dbread("prioritas_jurusan_3"), dbread("p3"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub cariSiswa()
        cbCariSiswa.DisplayMember = "Text"
        cbCariSiswa.ValueMember = "Key"
        cbCariSiswa.Items.Add(New With {Key .Text = "Kode Pendaftaran", Key .Value = "kode_pendaftaran"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Nama Lengkap", Key .Value = "nama_lengkap"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Alamat Siswa", Key .Value = "alamat_siswa"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Tanggal Lahir Siswa", Key .Value = "tgl_lhr_siswa"})
        cbCariSiswa.Items.Add(New With {Key .Text = "No Telp Siswa", Key .Value = "no_telp_siswa"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Jenis Kelamin Siswa", Key .Value = "jk_siswa"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Asal Sekolah", Key .Value = "asal_sekolah"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Prioritas Jurusan 1", Key .Value = "prioritas_jurusan_1"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Prioritas Jurusan 2", Key .Value = "prioritas_jurusan_2"})
        cbCariSiswa.Items.Add(New With {Key .Text = "Prioritas Jurusan 3", Key .Value = "prioritas_jurusan_3"})

        cbSortSiswa.DisplayMember = "Text"
        cbSortSiswa.ValueMember = "Key"
        cbSortSiswa.Items.Add(New With {Key .Text = "Kode Pendaftaran", Key .Value = "kode_pendaftaran"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Nama Lengkap", Key .Value = "nama_lengkap"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Alamat Siswa", Key .Value = "alamat_siswa"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Tanggal Lahir Siswa", Key .Value = "tgl_lhr_siswa"})
        cbSortSiswa.Items.Add(New With {Key .Text = "No Telp Siswa", Key .Value = "no_telp_siswa"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Jenis Kelamin Siswa", Key .Value = "jk_siswa"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Asal Sekolah", Key .Value = "asal_sekolah"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Prioritas Jurusan 1", Key .Value = "prioritas_jurusan_1"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Prioritas Jurusan 2", Key .Value = "prioritas_jurusan_2"})
        cbSortSiswa.Items.Add(New With {Key .Text = "Prioritas Jurusan 3", Key .Value = "prioritas_jurusan_3"})
    End Sub

    Private Sub searchSiswa(kolom, rows)
        DataGridViewSiswa.Rows.Clear()
        sql = "SELECT s.kode_pendaftaran, s.nama_lengkap, s.alamat_siswa, s.tgl_lhr_siswa, s.no_telp_siswa, s.jk_siswa, s.asal_sekolah, s.prioritas_jurusan_1, s.prioritas_jurusan_2, s.prioritas_jurusan_3, j.kode_jurusan, j.nama_jurusan AS p1, ja.nama_jurusan AS p2, jb.nama_jurusan AS p3 FROM dbo.Siswa AS s INNER JOIN dbo.Jurusan AS j ON s.prioritas_jurusan_1 = j.kode_jurusan INNER JOIN dbo.Jurusan AS ja ON s.prioritas_jurusan_2 = ja.kode_jurusan INNER JOIN dbo.Jurusan AS jb ON s.prioritas_jurusan_3 = jb.kode_jurusan where " & kolom & " like '%" & rows & "%'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewSiswa.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"), dbread("alamat_siswa"), dbread("tgl_lhr_siswa").toshortdatestring, dbread("no_telp_siswa"), dbread("jk_siswa"), dbread("asal_sekolah"), dbread("prioritas_jurusan_1"), dbread("p1"), dbread("prioritas_jurusan_2"), dbread("p2"), dbread("prioritas_jurusan_3"), dbread("p3"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub sortSiswa(kolom, by)
        DataGridViewSiswa.Rows.Clear()
        sql = "SELECT s.kode_pendaftaran, s.nama_lengkap, s.alamat_siswa, s.tgl_lhr_siswa, s.no_telp_siswa, s.jk_siswa, s.asal_sekolah, s.prioritas_jurusan_1, s.prioritas_jurusan_2, s.prioritas_jurusan_3, j.kode_jurusan, j.nama_jurusan AS p1, ja.nama_jurusan AS p2, jb.nama_jurusan AS p3 FROM dbo.Siswa AS s INNER JOIN dbo.Jurusan AS j ON s.prioritas_jurusan_1 = j.kode_jurusan INNER JOIN dbo.Jurusan AS ja ON s.prioritas_jurusan_2 = ja.kode_jurusan INNER JOIN dbo.Jurusan AS jb ON s.prioritas_jurusan_3 = jb.kode_jurusan order by " & kolom & " " & by & ""

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewSiswa.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"), dbread("alamat_siswa"), dbread("tgl_lhr_siswa").toshortdatestring, dbread("no_telp_siswa"), dbread("jk_siswa"), dbread("asal_sekolah"), dbread("prioritas_jurusan_1"), dbread("p1"), dbread("prioritas_jurusan_2"), dbread("p2"), dbread("prioritas_jurusan_3"), dbread("p3"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnCariSiswa_Click(sender As Object, e As EventArgs) Handles btnCariSiswa.Click
        If cbCariSiswa.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        searchSiswa(cbCariSiswa.SelectedItem.Value, txtCariSiswa.Text)
    End Sub

    Private Sub btnSortSiswa_Click(sender As Object, e As EventArgs) Handles btnSortSiswa.Click
        If cbCariSiswa.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim by As String
        If rbAscSiswa.Checked = True Then
            by = "ASC"
        ElseIf rbDescSiswa.Checked = True Then
            by = "DESC"
        End If
        Try
            sortSiswa(cbSortSiswa.SelectedItem.Value, by)
        Catch ex As Exception
            MsgBox("Tipe data ini tidak dapat di sort!", MsgBoxStyle.Exclamation)
        End Try
    End Sub


    Private Sub btnEditSiswa_Click(sender As Object, e As EventArgs) Handles btnEditSiswa.Click
        btnSaveSiswa.Enabled = True
    End Sub

    Private Sub resetSiswa()
        txtNamaSiswa.Text = ""
        txtAlamatSiswa.Text = ""
        DateTimePickerSiswa.Value = Date.Now
        txtPhoneSiswa.Text = ""
        rbLkSiswa.Checked = False
        rbPrSiswa.Checked = False
        txtSekolahSiswa.Text = ""
        cbP1.SelectedIndex = 0
        cbP2.SelectedIndex = 0
        cbP3.SelectedIndex = 0
    End Sub


    Private Sub btnResetSiswa_Click(sender As Object, e As EventArgs) Handles btnResetSiswa.Click
        resetSiswa()
    End Sub

    Private Sub DataGridViewSiswa_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewSiswa.CellMouseClick
        Dim i As Integer = DataGridViewSiswa.CurrentRow.Index

        txtNamaSiswa.Text = DataGridViewSiswa.Item(1, i).Value
        txtAlamatSiswa.Text = DataGridViewSiswa.Item(2, i).Value
        DateTimePickerSiswa.Value = DataGridViewSiswa.Item(3, i).Value
        txtPhoneSiswa.Text = DataGridViewSiswa.Item(4, i).Value
        If DataGridViewSiswa.Item(5, i).Value = "Laki-Laki" Then
            rbLkSiswa.Checked = True
        ElseIf DataGridViewSiswa.Item(5, i).Value = "Perempuan"
            rbPrSiswa.Checked = True
        End If
        txtSekolahSiswa.Text = DataGridViewSiswa.Item(6, i).Value
        cbP1.SelectedValue = DataGridViewSiswa.Item(7, i).Value
        cbP2.SelectedValue = DataGridViewSiswa.Item(9, i).Value
        cbP3.SelectedValue = DataGridViewSiswa.Item(11, i).Value
    End Sub

    Private Sub btnSaveSiswa_Click(sender As Object, e As EventArgs) Handles btnSaveSiswa.Click
        If txtNamaSiswa.Text = "" Or txtAlamatSiswa.Text = "" Or txtPhoneSiswa.Text = "" Or txtSekolahSiswa.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim jk As String

        If rbLkSiswa.Checked = False And rbPrSiswa.Checked = False Then
            MsgBox("Please select gender", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If rbLkSiswa.Checked = True Then
            jk = "Laki-Laki"
        ElseIf rbPrSiswa.Checked = True
            jk = "Perempuan"
        End If

        If Date.Now.ToString("yyyyMMdd") - DateTimePickerSiswa.Value.ToString("yyyyMMdd") < 180000 Then
            MsgBox("Registrant must older than 18th", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If cbP1.Text = cbP2.Text Or cbP1.Text = cbP3.Text Or cbP2.Text = cbP3.Text Then
            MsgBox("Program study priority must be different!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim i As Integer = DataGridViewSiswa.CurrentRow.Index

        db.updateSiswa(DataGridViewSiswa.Item(0, i).Value, txtNamaSiswa.Text, txtAlamatSiswa.Text, DateTimePickerSiswa.Value.ToString("yyyy/MM/dd"), txtPhoneSiswa.Text, jk, txtSekolahSiswa.Text, cbP1.SelectedValue, cbP2.SelectedValue, cbP3.SelectedValue)

        MsgBox("Update data success", MsgBoxStyle.Information)
        dataSiswa()
        resetSiswa()
        btnSaveSiswa.Enabled = False
    End Sub

    Private Sub btnDeleteSiswa_Click(sender As Object, e As EventArgs) Handles btnDeleteSiswa.Click
        Dim i As Integer = DataGridViewSiswa.CurrentRow.Index
        db.deleteSiswa(DataGridViewSiswa.Item(0, i).Value)

        MsgBox("Delete data success", MsgBoxStyle.Information)
        dataSiswa()
        resetSiswa()
    End Sub

#End Region

#Region "Akun"
    Public Sub dataAkun()
        DataGridViewAkun.Rows.Clear()
        sql = "select * from UserAccount"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                Dim role As String
                If dbread("role") = 1 Then
                    role = "Staff"
                ElseIf dbread("role") = 2 Then
                    role = "Guru"
                ElseIf dbread("role") = 3 Then
                    role = "Siswa"
                End If
                DataGridViewAkun.Rows.Add(dbread("username"), dbread("password"), role, dbread("kode_karyawan"), dbread("kode_pendaftaran"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnUpdateAkun_Click(sender As Object, e As EventArgs) Handles btnUpdateAkun.Click
        btnSaveAkun.Enabled = True
    End Sub

    Private Sub DataGridViewAkun_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewAkun.CellMouseClick
        Dim i As Integer = DataGridViewAkun.CurrentRow.Index

        txtUsernameAkun.Text = DataGridViewAkun.Item(0, i).Value
        txtPwAkun.Text = DataGridViewAkun.Item(1, i).Value
        If DataGridViewAkun.Item(2, i).Value = "Staff" Then
            rbStaffAkun.Checked = True
        ElseIf DataGridViewAkun.Item(2, i).Value = "Guru" Then
            rbGuruAkun.Checked = True
        ElseIf DataGridViewAkun.Item(2, i).Value = "Siswa" Then
            rbSiswaAkun.Checked = True
        End If
    End Sub

    Private Sub resetAkun()
        txtUsernameAkun.Text = ""
        txtPwAkun.Text = ""
        rbStaffAkun.Checked = False
        rbSiswaAkun.Checked = False
        rbGuruAkun.Checked = False
    End Sub

    Private Sub btnSaveAkun_Click(sender As Object, e As EventArgs) Handles btnSaveAkun.Click
        If txtUsernameAkun.Text = "" Or txtPwAkun.Text = "" Then
            MsgBox("All data must be fill!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If rbStaffAkun.Checked = False And rbSiswaAkun.Checked = False And rbGuruAkun.Checked = False Then
            MsgBox("Select role", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim i As Integer = DataGridViewAkun.CurrentRow.Index
        Dim role As String
        If rbStaffAkun.Checked = True Then
            role = 1
        ElseIf rbGuruAkun.Checked = True Then
            role = 2
        ElseIf rbSiswaAkun.Checked = True Then
            role = 3
        End If

        Dim a As String
        Dim b As String
        If DataGridViewAkun.Item(3, i).Value IsNot DBNull.Value Then
            a = DataGridViewAkun.Item(3, i).Value
            b = "kode_karyawan"
        ElseIf DataGridViewAkun.Item(4, i).Value IsNot DBNull.Value Then
            a = DataGridViewAkun.Item(4, i).Value
            b = "kode_pendaftaran"
        End If

        db.updateAkun(txtUsernameAkun.Text, txtPwAkun.Text, role, b, a)

        MsgBox("Update data success", MsgBoxStyle.Information)
        resetAkun()
        dataAkun()
        btnSaveAkun.Enabled = False
    End Sub
#End Region

#Region "Jadwal Ujian"
    Private Function kodeJadwal()
        Dim lastkode As String = db.maxKode("kode_ujian", "JadwalUjian")
        Dim kd As String = incrementKode(lastkode)
        txtKodeJadwal.Text = kd
        Return kd
    End Function

    Private Sub dataJadwalUjian()
        DataGridViewJadwal.Rows.Clear()
        sql = "select * from JadwalUjian"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewJadwal.Rows.Add(dbread("kode_ujian"), dbread("tanggal_ujian").toshortdatestring, dbread("durasi_ujian"), dbread("kapasitas_peserta"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub cariJadwal()
        cbCariJadwal.DisplayMember = "Text"
        cbCariJadwal.ValueMember = "Key"
        cbCariJadwal.Items.Add(New With {Key .Text = "Kode Ujian", Key .Value = "kode_ujian"})
        cbCariJadwal.Items.Add(New With {Key .Text = "Tanggal Ujian", Key .Value = "tanggal_ujian"})
        cbCariJadwal.Items.Add(New With {Key .Text = "Durasi Ujian", Key .Value = "durasi_ujian"})
        cbCariJadwal.Items.Add(New With {Key .Text = "Kapasitas Peserta", Key .Value = "kapasitas_peserta"})

        cbSortJadwal.DisplayMember = "Text"
        cbSortJadwal.ValueMember = "Key"
        cbSortJadwal.Items.Add(New With {Key .Text = "Kode Ujian", Key .Value = "kode_ujian"})
        cbSortJadwal.Items.Add(New With {Key .Text = "Tanggal Ujian", Key .Value = "tanggal_ujian"})
        cbSortJadwal.Items.Add(New With {Key .Text = "Durasi Ujian", Key .Value = "durasi_ujian"})
        cbSortJadwal.Items.Add(New With {Key .Text = "Kapasitas Peserta", Key .Value = "kapasitas_peserta"})
    End Sub

    Private Sub searchJadwal(kolom, rows)
        DataGridViewJadwal.Rows.Clear()
        sql = "select * from JadwalUjian where " & kolom & " like '%" & rows & "%'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewJadwal.Rows.Add(dbread("kode_ujian"), dbread("tanggal_ujian").toshortdatestring, dbread("durasi_ujian"), dbread("kapasitas_peserta"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub sortJadwal(kolom, by)
        DataGridViewJadwal.Rows.Clear()
        sql = "select * from JadwalUjian order by " & kolom & " " & by & ""

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewJadwal.Rows.Add(dbread("kode_ujian"), dbread("tanggal_ujian").toshortdatestring, dbread("durasi_ujian"), dbread("kapasitas_peserta"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnCariJadwal_Click(sender As Object, e As EventArgs) Handles btnCariJadwal.Click
        If cbCariJadwal.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        searchJadwal(cbCariJadwal.SelectedItem.Value, txtCariJadwal.Text)
    End Sub

    Private Sub btnSortJadwal_Click(sender As Object, e As EventArgs) Handles btnSortJadwal.Click
        If cbSortJadwal.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim by As String
        If rbAscJadwal.Checked = True Then
            by = "ASC"
        ElseIf rbDescJadwal.Checked = True Then
            by = "DESC"
        End If
        Try
            sortJadwal(cbSortJadwal.SelectedItem.Value, by)
        Catch ex As Exception
            MsgBox("Tipe data ini tidak dapat di sort!", MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub DataGridViewJadwal_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewJadwal.CellClick
        If e.ColumnIndex = 4 Then
            getsoal()
        End If
    End Sub

    Private Sub getsoal()
        Dim i As Integer = DataGridViewJadwal.CurrentRow.Index
        SaveFileDialog1.Reset()
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            Dim SaveLocation As String = SaveFileDialog1.FileName
            dbcomm = New SqlCommand("select soal_ujian from JadwalUjian where kode_ujian='" & DataGridViewJadwal.Item(0, i).Value & "'", db.conn)
            dbread = dbcomm.ExecuteReader
            If dbread.HasRows Then
                Do While dbread.Read
                    Dim x As Byte() = CType(dbread("soal_ujian"), Byte()).ToArray()
                    Dim fs As New FileStream(SaveLocation, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                    fs.Write(x, 0, x.Length)
                    fs.Close()
                    MsgBox("Download success", MsgBoxStyle.Information)
                    Process.Start(SaveLocation)
                    Exit Do
                Loop
            End If
            dbread.Close()
        End If
    End Sub

    Private Sub btnSoalJadwal_Click(sender As Object, e As EventArgs) Handles btnSoalJadwal.Click
        OpenFileDialog1.Reset()
        OpenFileDialog1.Filter = "DOCX Files(*.docx)|*.docx|PDF Files(*.pdf)|*.pdf"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            txtSoalJadwal.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub resetJadwal()
        DateTimePickerJadwal.Value = Date.Now
        txtDurasiJadwal.Text = ""
        txtKapasitasJadwal.Text = ""
        txtSoalJadwal.Text = ""
        kodeJadwal()
    End Sub

    Private Sub btnResetJadwal_Click(sender As Object, e As EventArgs) Handles btnResetJadwal.Click
        resetJadwal()
    End Sub

    Private Sub btnAddJadwal_Click(sender As Object, e As EventArgs) Handles btnAddJadwal.Click
        If txtDurasiJadwal.Text = "" Or txtKapasitasJadwal.Text = "" Or txtSoalJadwal.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim ms As New MemoryStream
        Dim fs As FileStream = New FileStream(txtSoalJadwal.Text, FileMode.Open, FileAccess.Read)
        fs.CopyTo(ms)

        db.insertJadwal(txtKodeJadwal.Text, DateTimePickerJadwal.Value.ToString("yyyy/MM/dd"), txtDurasiJadwal.Text, txtKapasitasJadwal.Text, ms.ToArray)
        MsgBox("Insert data success", MsgBoxStyle.Information)

        ms.Close()
        fs.Close()
        resetJadwal()
        dataJadwalUjian()
        jadwalUjian()
    End Sub

    Private Sub btnUpdateJadwal_Click(sender As Object, e As EventArgs) Handles btnUpdateJadwal.Click
        btnSaveJadwal.Enabled = True
    End Sub

    Private Sub btnSaveJadwal_Click(sender As Object, e As EventArgs) Handles btnSaveJadwal.Click
        If txtKodeJadwal.Text = "" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If txtDurasiJadwal.Text = "" Or txtKapasitasJadwal.Text = "" Or txtSoalJadwal.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If Not txtSoalJadwal.Text = "" Then
            Dim ms As New MemoryStream
            Dim fs As FileStream = New FileStream(txtSoalJadwal.Text, FileMode.Open, FileAccess.Read)
            fs.CopyTo(ms)
            db.updateJadwalSoal(txtKodeJadwal.Text, DateTimePickerJadwal.Value.ToString("yyyy/MM/dd"), txtDurasiJadwal.Text, txtKapasitasJadwal.Text, ms.ToArray)
            MsgBox("Update data success", MsgBoxStyle.Information)
            ms.Close()
            fs.Close()
            resetJadwal()
            dataJadwalUjian()
            jadwalUjian()
            btnSaveJadwal.Enabled = False
            Exit Sub
        End If

        If txtSoalJadwal.Text = "" Then
            db.updateJadwal(txtKodeJadwal.Text, DateTimePickerJadwal.Value.ToString("yyyy/MM/dd"), txtDurasiJadwal.Text, txtKapasitasJadwal.Text)
            MsgBox("Update data success", MsgBoxStyle.Information)
            resetJadwal()
            dataJadwalUjian()
            jadwalUjian()
            btnSaveJadwal.Enabled = False
            Exit Sub
        End If
    End Sub

    Private Sub btnDeleteJadwal_Click(sender As Object, e As EventArgs) Handles btnDeleteJadwal.Click
        If txtKodeJadwal.Text = "" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        db.deleteJadwal(txtKodeJadwal.Text)
        dataJadwalUjian()
        jadwalUjian()
    End Sub

    Private Sub DataGridViewJadwal_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewJadwal.CellMouseClick
        Dim i As Integer = DataGridViewJadwal.CurrentRow.Index

        txtKodeJadwal.Text = DataGridViewJadwal.Item(0, i).Value
        DateTimePickerJadwal.Value = DataGridViewJadwal.Item(1, i).Value
        txtDurasiJadwal.Text = DataGridViewJadwal.Item(2, i).Value
        txtKapasitasJadwal.Text = DataGridViewJadwal.Item(3, i).Value
    End Sub

#End Region

#Region "Alokasi Peserta Ujian"
    Private Sub KapasitasPeserta(kd)
        sql = "select kapasitas_peserta from JadwalUjian where kode_ujian='" & kd & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                lblKapasitas.Text = dbread("kapasitas_peserta")
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub jadwalUjian()
        sql = "select * from JadwalUjian"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                cbAlokasi.Items.Add(dbread("kode_ujian") & "-" & dbread("tanggal_ujian"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub cbAlokasi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbAlokasi.SelectedIndexChanged
        Dim a As String = cbAlokasi.Text
        KapasitasPeserta(a.Substring(0, 5))
        siswaNotNull()

        lblKapasitas.Text = lblKapasitas.Text - DataGridViewNotNull.RowCount()
    End Sub

    Private Sub siswaNull()
        DataGridViewNull.Rows.Clear()
        sql = "select s.kode_pendaftaran,s.nama_lengkap from Siswa as s left join AktivitasUjian as au on s.kode_pendaftaran=au.kode_pendaftaran where au.kode_pendaftaran is NULL"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewNull.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub siswaNotNull()
        DataGridViewNotNull.Rows.Clear()
        Dim a As String = cbAlokasi.Text
        sql = "select s.kode_pendaftaran,s.nama_lengkap from Siswa as s join AktivitasUjian as au on s.kode_pendaftaran=au.kode_pendaftaran where kode_ujian='" & a.Substring(0, 5) & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewNotNull.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnInsertAlokasi_Click(sender As Object, e As EventArgs) Handles btnInsertAlokasi.Click
        If cbAlokasi.Text = "Pilih" Then
            MsgBox("Select exam schedule first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If lblKapasitas.Text = 0 Then
            MsgBox("Capacity Max!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim cb As String = cbAlokasi.Text
        Dim kodeUjian As String = cb.Substring(0, 5)
        Dim i As Integer = DataGridViewNull.CurrentRow.Index
        db.insertSiswaUjian(kodeUjian, DataGridViewNull.Item(0, i).Value, idstaff)

        MsgBox("Insert data success", MsgBoxStyle.Information)
        siswaNull()
        siswaNotNull()

        lblKapasitas.Text -= 1
    End Sub

    Private Sub btnDeleteAlokasi_Click(sender As Object, e As EventArgs) Handles btnDeleteAlokasi.Click
        Dim i As Integer = DataGridViewNotNull.CurrentRow.Index
        db.deleteSiswaUjian(DataGridViewNotNull.Item(0, i).Value)

        MsgBox("Delete data success", MsgBoxStyle.Information)
        siswaNull()
        siswaNotNull()

        lblKapasitas.Text += 1
    End Sub
#End Region

#Region "Koreksi Hasil Ujian"
    Private Sub dataPassed()
        sql = "select * from JadwalUjian"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read()
                cbKodeUjian.Items.Add(dbread("kode_ujian") & " - " & dbread("tanggal_ujian") & " - " & dbread("durasi_ujian") & " Menit")
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub cariKoreksi()
        cbCariKoreksi.DisplayMember = "Text"
        cbCariKoreksi.ValueMember = "Key"
        cbCariKoreksi.Items.Add(New With {Key .Text = "Kode Pendaftaran", Key .Value = "s.kode_pendaftaran"})
        cbCariKoreksi.Items.Add(New With {Key .Text = "Nama Lengkap", Key .Value = "s.nama_lengkap"})
        cbCariKoreksi.Items.Add(New With {Key .Text = "Nilai PG", Key .Value = "au.nilai_pg"})
        cbCariKoreksi.Items.Add(New With {Key .Text = "Nilai Essay", Key .Value = "au.nilai_essay"})
        cbCariKoreksi.Items.Add(New With {Key .Text = "Nilai Kasus", Key .Value = "au.nilai_kasus"})
    End Sub

    Private Sub searchKoreksi(kolom, rows)
        DataGridViewKoreksi.Rows.Clear()
        Dim kd As String = cbKodeUjian.Text
        sql = "select * from AktivitasUjian as au join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian join Siswa as s on au.kode_pendaftaran=s.kode_pendaftaran where ju.kode_ujian='" & kd.Substring(0, 5) & "' and " & kolom & " like '%" & rows & "%'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read()
                DataGridViewKoreksi.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"), dbread("nilai_pg"), dbread("nilai_essay"), dbread("nilai_kasus"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnCariKoreksi_Click(sender As Object, e As EventArgs) Handles btnCariKoreksi.Click
        If cbCariKoreksi.Text = "Pilih" Then
            MsgBox("Select data first!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        searchKoreksi(cbCariKoreksi.SelectedItem.Value, txtCariKoreksi.Text)
    End Sub

    Private Sub cbKodeUjian_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbKodeUjian.SelectedIndexChanged
        dataKoreksi()
    End Sub

    Private Sub dataKoreksi()
        DataGridViewKoreksi.Rows.Clear()
        Dim kd As String = cbKodeUjian.Text
        sql = "select * from AktivitasUjian as au join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian join Siswa as s on au.kode_pendaftaran=s.kode_pendaftaran where ju.kode_ujian='" & kd.Substring(0, 5) & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read()
                DataGridViewKoreksi.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"), dbread("nilai_pg"), dbread("nilai_essay"), dbread("nilai_kasus"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub DataGridViewKoreksi_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewKoreksi.CellClick
        If e.ColumnIndex = 5 Then
            SaveFileDialog1.Reset()
            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                downloadJawaban(SaveFileDialog1.FileName)
            End If
        End If
    End Sub

    Private Sub downloadJawaban(loc)
        Dim i As Integer = DataGridViewKoreksi.CurrentRow.Index
        sql = "select jawaban_siswa from AktivitasUjian as au join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian join Siswa as s on au.kode_pendaftaran=s.kode_pendaftaran where s.kode_pendaftaran='" & DataGridViewKoreksi.Item(0, i).Value & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()

            If dbread.HasRows() Then
                Dim a As Byte() = CType(dbread("jawaban_siswa"), Byte()).ToArray()
                Dim fs = New FileStream(loc, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                fs.Write(a, 0, a.Length)
                fs.Close()
                MsgBox("Download success", MsgBoxStyle.Information)
                Process.Start(loc)
            Else
                MsgBox("Not yet", MsgBoxStyle.Exclamation)
            End If
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub DataGridViewKoreksi_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewKoreksi.CellEndEdit
        Dim i As Integer = DataGridViewKoreksi.CurrentRow.Index
        db.nilaiPg(DataGridViewKoreksi.Item(2, i).Value, DataGridViewKoreksi.Item(0, i).Value)
        Try
            db.nilaiEssay(DataGridViewKoreksi.Item(3, i).Value, DataGridViewKoreksi.Item(0, i).Value)
        Catch ex As Exception

        End Try

        Try
            db.nilaiKasus(DataGridViewKoreksi.Item(4, i).Value, DataGridViewKoreksi.Item(0, i).Value)
        Catch ex As Exception

        End Try

        lastUpdateKoreksi()
    End Sub

    Private Sub lastUpdateKoreksi()
        lblDatetimeKoreksi.Text = DateTime.Now().ToString("yyyy/MM/dd hh:mm:ss")
    End Sub
#End Region

#Region "Laporan Hasil Ujian"
    Private Sub kodeLaporan()
        sql = "select * from JadwalUjian"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read()
                cbKodeUjianLaporan.Items.Add(dbread("kode_ujian") & " - " & dbread("tanggal_ujian") & " - " & dbread("durasi_ujian") & " Menit")
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub cbKodeUjianLaporan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbKodeUjianLaporan.SelectedIndexChanged
        dataLaporan()
    End Sub

    Private Sub dataLaporan()
        DataGridViewLaporan.Rows.Clear()
        Dim a As String = cbKodeUjianLaporan.Text
        sql = "select * from AktivitasUjian as au join JadwalUjian as ju on au.kode_ujian=ju.kode_ujian join siswa as s on au.kode_pendaftaran=s.kode_pendaftaran where ju.kode_ujian='" & a.Substring(0, 5) & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                Dim status As String
                Dim totalScore As String = (30% * dbread("nilai_pg")) + (30% * dbread("nilai_essay")) + (40% * dbread("nilai_kasus"))
                If totalScore.Substring(0, 2) >= 65 Then
                    status = "Lulus"
                ElseIf totalScore.Substring(0, 2) < 65 Then
                    status = "Tidak Lulus"
                End If
                DataGridViewLaporan.Rows.Add(dbread("kode_pendaftaran"), dbread("nama_lengkap"), totalScore.Substring(0, 2), status)
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnHasilTest_Click(sender As Object, e As EventArgs) Handles btnHasilTest.Click
        Dim a As Integer
        Dim b As Integer
        For i = 0 To (DataGridViewLaporan.Rows.Count - 1)
            If DataGridViewLaporan.Rows(i).Cells("Column43").Value = "Lulus" Then
                a += 1
            End If
            If DataGridViewLaporan.Rows(i).Cells("Column43").Value = "Tidak Lulus" Then
                b += 1
            End If
        Next

        Test_Result_Graphic.Chart1.Series.Clear()
        Test_Result_Graphic.Chart1.Series.Add("Jumlah Peserta")

        Test_Result_Graphic.Chart1.Series("Jumlah Peserta").Points.AddXY("Lulus", a)
        Test_Result_Graphic.Chart1.Series("Jumlah Peserta").Points.AddXY("Tidak Lulus", b)
        Test_Result_Graphic.Label1.Text = cbKodeUjianLaporan.Text
        Test_Result_Graphic.Show()

    End Sub

#End Region
End Class