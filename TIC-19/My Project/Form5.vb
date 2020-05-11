Public Class frmbismillah
    Dim sqlnya As String
    Sub panggildata()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_tic", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_tic")
    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Sub hitung()
        TextBox1.Text = ListBox2.Items.Count
    End Sub
    Private Sub frmbismillah_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedItem = "" Then
            MsgBox("Silahkan pilih pilihan yang anda lakukan terlebih dahulu")
        Else
            ListBox2.Items.Add(ListBox1.SelectedItem)
            ListBox1.Items.Remove(ListBox1.SelectedItem)
        End If
        hitung()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox2.SelectedItem = "" Then
            MsgBox("Silahkan pilih pilihan yang ingin anda hapus terlebih dahulu")
        Else
            ListBox1.Items.Add(ListBox2.SelectedItem)
            ListBox2.Items.Remove(ListBox2.SelectedItem)
        End If
        hitung()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox1.Enabled = False
        If TextBox1.Text <= 7 Then
            Label2.Text = "Resiko Rendah"
        ElseIf TextBox1.Text >= 8 And TextBox1.Text <= 14 Then
            Label2.Text = "Resiko Sedang"
        Else
            Label2.Text = "Resiko Tinggi"
        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Label2.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        sqlnya = "insert into tb_tic (Nama,Telp,Alamat,Pilihan,Diagnosa)values('" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox1.Text & "','" & Label2.Text & "')"
        Call jalan()
        MsgBox("Data berhasil terinput")
        Call panggildata()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        sqlnya = "delete from tb_tic where Nama='" & TextBox2.Text & "'"
        Call jalan()
        MsgBox("Data berhasil dihapus")
        Call panggildata()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmawal.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click
        Me.Close()
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click
        If Me.WindowState = WindowState.Normal Then
            Me.WindowState = WindowState.Maximized
        Else
            Me.WindowState = WindowState.Normal
        End If
    End Sub

    Private Sub PictureBox14_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click
        Me.WindowState = WindowState.Minimized
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        frmregister.Show()
        Me.Hide()
    End Sub
End Class