Imports System.Text
Imports System.IO
Imports System.Data.SqlClient
Public Class Main_Frm
    Dim db As New Database
    Dim dbcomm As New SqlCommand
    Dim sql As String
    Dim dbread As SqlDataReader
    Dim kodeSiswa As String = "S0001"
    Private Sub Main_Frm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the '_2016DataSet.Jurusan' table. You can move, or remove it, as needed.
        Me.JurusanTableAdapter.Fill(Me._2016DataSet.Jurusan)
        TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed

        db.conn()
        captcha()
        detailJurusan()
        getKodeSiswa()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectTab(0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TabControl1.SelectTab(1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TabControl1.SelectTab(2)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Login_Frm.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub

    Private Sub captcha()
        Dim randomText As New StringBuilder
        Dim r As New Random
        Dim alphabet As String = "qwertyuiopasdfghjklzxcvbnm1234567890!@#$%^&*()QWEIOPASDFGHJKLZXCVBNM"
        For i As Integer = 1 To 10
            randomText.Append(alphabet(r.[Next](alphabet.Length)))
        Next
        lblCaptcha.Text = randomText.ToString
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        captcha()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        detailJurusan()
    End Sub

    Private Sub detailJurusan()
        sql = "select * from Jurusan where kode_jurusan='" & ComboBox1.SelectedValue & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                lblNamaJurusan.Text = dbread("nama_jurusan")
                Dim gambar As Byte() = dbread("foto_jurusan")
                Dim ms As MemoryStream = New MemoryStream(gambar)
                PictureBoxJurusan.Image = Image.FromStream(ms)
                ms.Close()
                lblDescJurusan.Text = dbread("desc_jurusan")
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        resetRegister()
    End Sub

    Private Sub resetRegister()
        txtNama.Text = ""
        txtAlamat.Text = ""
        DateTimePicker1.Value = Date.Now
        txtPhone.Text = ""
        rbLk.Checked = False
        rbPr.Checked = False
        txtSekolah.Text = ""
        txtCaptcha.Text = ""
        cbP1.SelectedIndex = 0
        cbP2.SelectedIndex = 0
        cbP3.SelectedIndex = 0
        getKodeSiswa()
    End Sub

    Private Function incrementKode(kode)
        Dim huruf As String = kode.Substring(0, 1)
        Dim id As String = kode.Substring(1)
        Dim angka As Integer = Integer.Parse(id)
        angka += 1
        kode = huruf & angka.ToString("D" & id.Length)
        kodeSiswa = huruf & angka.ToString("D" & id.Length)
        Return kodeSiswa
    End Function

    Private Function getKodeSiswa()
        Dim lastkode As String = db.maxKode("kode_pendaftaran", "Siswa")
        Dim kd As String = incrementKode(lastkode)
        kodeSiswa = kd
        Return kd
    End Function

    Private Sub btnDaftar_Click(sender As Object, e As EventArgs) Handles btnDaftar.Click
        If txtNama.Text = "" Or txtAlamat.Text = "" Or txtPhone.Text = "" Or txtSekolah.Text = "" Then
            MsgBox("All data must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim jk As String

        If rbLk.Checked = False And rbPr.Checked = False Then
            MsgBox("Please select gender", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If rbLk.Checked = True Then
            jk = "Laki-Laki"
        ElseIf rbPr.Checked = True
            jk = "Perempuan"
        End If

        If Date.Now.ToString("yyyyMMdd") - DateTimePicker1.Value.ToString("yyyyMMdd") < 180000 Then
            MsgBox("Registrant must older than 18th", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If cbP1.Text = cbP2.Text Or cbP1.Text = cbP3.Text Or cbP2.Text = cbP3.Text Then
            MsgBox("Program study priority must be different!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If txtCaptcha.Text = "" Then
            MsgBox("Captcha must be fill", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If Not lblCaptcha.Text = txtCaptcha.Text Then
            MsgBox("Captcha wrong!", MsgBoxStyle.Exclamation)
            captcha()
            Exit Sub
        End If

        db.registerSiswa(kodeSiswa, txtNama.Text, txtAlamat.Text, DateTimePicker1.Value.ToString("yyyy/MM/dd"), txtPhone.Text, jk, txtSekolah.Text, cbP1.SelectedValue, cbP2.SelectedValue, cbP3.SelectedValue)
        db.insertAkun("kode_pendaftaran", kodeSiswa, kodeSiswa + DateTimePicker1.Value.ToString("MMddyy"), 3, kodeSiswa)

        MsgBox("Registration success", MsgBoxStyle.Information)

        Registration_Confirmation.lblNoPendaftaran.Text = kodeSiswa
        Registration_Confirmation.lblNama.Text = txtNama.Text
        Registration_Confirmation.lblP1.Text = cbP1.Text
        Registration_Confirmation.lblP2.Text = cbP2.Text
        Registration_Confirmation.lblP3.Text = cbP3.Text
        Registration_Confirmation.lblUname.Text = kodeSiswa
        Registration_Confirmation.lblPw.Text = kodeSiswa + DateTimePicker1.Value.ToString("MMddyy")
        Registration_Confirmation.Show()

        resetRegister()
    End Sub

    Private Sub txtPhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPhone.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class