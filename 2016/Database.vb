Imports System.Data.SqlClient
Public Class Database
    Dim dbconn As New SqlConnection
    Dim dbcomm As New SqlCommand
    Dim dbread As SqlDataReader
    Dim sql As String
    Dim lastid As String
    Dim lastkode As String

    Public Function conn()
        dbconn = New SqlConnection("data source=.\SQLEXPRESS;database=2016;integrated security=true")

        Try
            dbconn.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dbconn
    End Function

    Public Function crud(sql As String)
        Try
            dbcomm = New SqlCommand(sql, conn)
            dbread = dbcomm.ExecuteReader()
            dbread.Read()
            Return dbread
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Function

    Public Function crud(sqlcommand As SqlCommand)
        Try
            dbread = dbcomm.ExecuteReader()
            dbread.Read()
            Return dbread
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Function

    Public Function crudid(sql)
        Try
            dbcomm = New SqlCommand(sql, conn)
            lastid = dbcomm.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return lastid
    End Function

    Public Function maxKode(kolom, table)
        sql = "select max(" & kolom & ") as max from " & table & ""

        crud(sql)
        lastkode = dbread("max")
        dbread.Close()
        Return lastkode
    End Function

    Public Sub insertAkun(kd, uname, pw, role, kdval)
        sql = "insert into UserAccount(username,password,role," & kd & ") values('" & uname & "','" & pw & "','" & role & "','" & kdval & "')"
        MsgBox(sql)

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub login(uname, pw, role)
        sql = "select * from UserAccount where username='" & uname & "' and CONVERT(VARCHAR, password)='" & pw & "' and role=" & role & ""

        crud(sql)
        If dbread.HasRows Then
            Login_Frm.role = dbread("role")
        Else
            Login_Frm.eror = True
        End If
        dbread.Close()
    End Sub

#Region "Jurusan"
    Public Sub insertJurusan(kd As String, nm As String, desc As String, foto As Byte())
        dbcomm = New SqlCommand("insert into Jurusan(kode_jurusan,nama_jurusan,desc_jurusan,foto_jurusan) values('" & kd & "','" & nm & "','" & desc & "',@fotoJurusan)", conn)
        dbcomm.Parameters.Add("@fotoJurusan", SqlDbType.Binary, foto.Length).Value = foto
        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub deleteJurusan(kd)
        sql = "delete from Jurusan where kode_jurusan='" & kd & "'"
        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateJurusan(kd As String, nm As String, desc As String)
        dbcomm = New SqlCommand("update Jurusan set nama_jurusan='" & nm & "', desc_jurusan='" & desc & "' where kode_jurusan='" & kd & "'", conn)
        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub updateJurusanImage(kd As String, nm As String, desc As String, foto As Byte())
        dbcomm = New SqlCommand("update Jurusan set nama_jurusan='" & nm & "', desc_jurusan='" & desc & "',foto_jurusan=@fotoJurusan where kode_jurusan='" & kd & "'", conn)
        dbcomm.Parameters.Add("@fotoJurusan", SqlDbType.Binary, foto.Length).Value = foto
        crud(dbcomm)
        dbread.Close()
    End Sub
#End Region

#Region "Karyawan dan Guru"
    Public Sub insertGuru(kd, nip, nm, lhr, alamat, jk, telp)
        sql = "insert into Karyawan(kode_karyawan,NIP,nama_karyawan,tgl_lhr_karyawan,alamat_karyawan,jk_karyawan,no_telp_karyawan) values('" & kd & "','" & nip & "','" & nm & "','" & lhr & "','" & alamat & "','" & jk & "','" & telp & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteGuru(kd)
        sql = "delete Karyawan where kode_karyawan='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateGuru(kd, nip, nm, lhr, alamat, jk, telp)
        sql = "update Karyawan set NIP='" & nip & "',nama_karyawan='" & nm & "',tgl_lhr_karyawan='" & lhr & "',alamat_karyawan='" & alamat & "',jk_karyawan='" & jk & "',no_telp_karyawan='" & telp & "' where kode_karyawan='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Siswa"
    Public Sub registerSiswa(kd, nm, alamat, lhr, telp, jk, sekolah, p1, p2, p3)
        sql = "insert into Siswa(kode_pendaftaran,nama_lengkap,alamat_siswa,tgl_lhr_siswa,no_telp_siswa,jk_siswa,asal_sekolah,prioritas_jurusan_1,prioritas_jurusan_2,prioritas_jurusan_3) values('" & kd & "','" & nm & "','" & alamat & "','" & lhr & "','" & telp & "','" & jk & "','" & sekolah & "','" & p1 & "','" & p2 & "','" & p3 & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateSiswa(kd, nm, alamat, lhr, telp, jk, sekolah, p1, p2, p3)
        sql = "update Siswa set nama_lengkap='" & nm & "',alamat_siswa='" & alamat & "',tgl_lhr_siswa='" & lhr & "',no_telp_siswa='" & telp & "',jk_siswa='" & jk & "',asal_sekolah='" & sekolah & "',prioritas_jurusan_1='" & p1 & "',prioritas_jurusan_2='" & p2 & "',prioritas_jurusan_3='" & p3 & "' where kode_pendaftaran='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteSiswa(kd)
        sql = "delete Siswa where kode_pendaftaran='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub

#End Region

#Region "UserAkun"
    Public Sub updateAkun(uname, pw, role, kd, kdval)
        sql = "update UserAccount set username='" & uname & "', password='" & pw & "', role=" & role & " where " & kd & "='" & kdval & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Jadwal Ujian"
    Public Sub insertJadwal(kd As String, tgl As String, drs As String, kps As String, soal As Byte())
        dbcomm = New SqlCommand("insert into JadwalUjian(kode_ujian,tanggal_ujian,durasi_ujian,kapasitas_peserta,soal_ujian) values('" & kd & "','" & tgl & "','" & drs & "','" & kps & "',@soalUjian)", conn)
        dbcomm.Parameters.Add("@soalUjian", SqlDbType.Binary, soal.Length).Value = soal

        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub updateJadwalSoal(kd As String, tgl As String, drs As String, kps As String, soal As Byte())
        dbcomm = New SqlCommand("update JadwalUjian set tanggal_ujian='" & tgl & "',durasi_ujian='" & drs & "',kapasitas_peserta='" & kps & "',soal_ujian=@soalUjian where kode_ujian='" & kd & "'", conn)
        dbcomm.Parameters.Add("@soalUjian", SqlDbType.Binary, soal.Length).Value = soal

        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub updateJadwal(kd As String, tgl As String, drs As String, kps As String)
        dbcomm = New SqlCommand("update JadwalUjian set tanggal_ujian='" & tgl & "',durasi_ujian='" & drs & "',kapasitas_peserta='" & kps & "' where kode_ujian='" & kd & "'", conn)

        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub deleteJadwal(kd)
        sql = "delete JadwalUjian where kode_ujian='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Alokasi Peserta Ujian"
    Public Sub insertSiswaUjian(ujian, daftar, staff)
        sql = "insert into AktivitasUjian(kode_ujian,kode_pendaftaran,kode_karyawan) values('" & ujian & "','" & daftar & "','" & staff & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteSiswaUjian(daftar)
        sql = "delete AktivitasUjian where kode_pendaftaran='" & daftar & "'"

        crud(sql)
        dbread.Close()
    End Sub

#End Region

#Region "Student Nav"
    Public Sub uploadJawaban1(waktu As DateTime, deadline As DateTime, jawaban As Byte(), kd As String)
        dbcomm = New SqlCommand("update AktivitasUjian set waktu_upload_1='" & waktu.ToString("yyyy/MM/dd hh:mm:ss") & "', deadline_pengumpulan='" & deadline.ToString("yyyy/MM/dd hh:mm:ss") & "', jawaban_siswa=@jawaban where kode_pendaftaran='" & kd & "'", conn)
        dbcomm.Parameters.Add("@jawaban", SqlDbType.Binary, jawaban.Length).Value = jawaban
        MsgBox(dbcomm.CommandText)
        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub uploadJawaban2(waktu1 As DateTime, waktu2 As DateTime, deadline As DateTime, jawaban As Byte(), kd As String)
        dbcomm = New SqlCommand("update AktivitasUjian set waktu_upload_1='" & waktu1.ToString("yyyy/MM/dd hh:mm:ss") & "',waktu_upload_2='" & waktu2.ToString("yyyy/MM/dd hh:mm:ss") & "', deadline_pengumpulan='" & deadline.ToString("yyyy/MM/dd hh:mm:ss") & "', jawaban_siswa=@jawaban where kode_pendaftaran='" & kd & "'", conn)
        dbcomm.Parameters.Add("@jawaban", SqlDbType.Binary, jawaban.Length).Value = jawaban

        crud(dbcomm)
        dbread.Close()
    End Sub

    Public Sub uploadJawaban3(waktu1 As DateTime, waktu2 As DateTime, waktu3 As DateTime, deadline As DateTime, jawaban As Byte(), kd As String)
        dbcomm = New SqlCommand("update AktivitasUjian set waktu_upload_1='" & waktu1.ToString("yyyy/MM/dd hh:mm:ss") & "',waktu_upload_2='" & waktu2.ToString("yyyy/MM/dd hh:mm:ss") & "',waktu_upload_3='" & waktu3.ToString("yyyy/MM/dd hh:mm:ss") & "', deadline_pengumpulan='" & deadline.ToString("yyyy/MM/dd hh:mm:ss") & "', jawaban_siswa=@jawaban where kode_pendaftaran='" & kd & "'", conn)
        dbcomm.Parameters.Add("@jawaban", SqlDbType.Binary, jawaban.Length).Value = jawaban

        crud(dbcomm)
        dbread.Close()
    End Sub
#End Region

#Region "Koreksi"
    Public Sub nilaiPg(pg, kd)
        sql = "update AktivitasUjian set nilai_pg='" & pg & "' where kode_pendaftaran='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub nilaiEssay(essay, kd)
        sql = "update AktivitasUjian set nilai_essay='" & essay & "' where kode_pendaftaran='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub nilaiKasus(kasus, kd)
        sql = "update AktivitasUjian set nilai_kasus='" & kasus & "' where kode_pendaftaran='" & kd & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region
End Class
