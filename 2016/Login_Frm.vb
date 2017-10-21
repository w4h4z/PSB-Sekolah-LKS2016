Public Class Login_Frm
    Dim db As New Database
    Public role As String
    Public eror As Boolean
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            lblError.Visible = True
            lblError.Text = "All data must be fill"
            Exit Sub
        End If

        If ComboBox1.Text = "Pilih Role" Then
            lblError.Visible = True
            lblError.Text = "Please select Role!"
            Exit Sub
        End If

        Dim roleLogin As Integer = ComboBox1.SelectedIndex + 1
        db.login(TextBox1.Text, TextBox2.Text, roleLogin)

        If role = 1 Then
            Staff_Nav.idstaff = TextBox1.Text
            Staff_Nav.Show()
            Me.Close()
            Main_Frm.Close()
        ElseIf role = 2 Then
            Teacher_Nav.Show()
            Me.Close()
            Main_Frm.Close()
        ElseIf role = 3 Then
            Student_Nav.idsiswa = TextBox1.Text
            Student_Nav.Show()
            Me.Close()
            Main_Frm.Close()
        End If

        If eror = True Then
            lblError.Visible = True
            lblError.Text = "Data not found, check your Username and Password!"
        End If
    End Sub
End Class
