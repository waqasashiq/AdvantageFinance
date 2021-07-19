Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        ComboBox2.Visible = False
        ComboBox2.Text = "choose"
        ComboBox2.Items.Add("Approved")
        ComboBox2.Items.Add("Declined")


        ComboBox1.Text = "choose"
        ComboBox1.Items.Add("ID")
        ComboBox1.Items.Add("First Name")
        ComboBox1.Items.Add("Surname")
        ComboBox1.Items.Add("Application Status")
        ComboBox1.Items.Add("Credit Score")



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 3 Then
            TextBox1.Visible = False
            ComboBox2.Visible = True


        End If

        If ComboBox1.SelectedIndex <> 3 Then
            ComboBox2.Visible = False
            TextBox1.Visible = True


        End If


    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        Dim iD As String
        Dim Firstname As String
        Dim Lastname As String
        Dim Status As String
        Dim CreditScore As String
        Dim textIt As String
        Dim sindex As Integer
        textIt = TextBox1.Text

        If ComboBox1.SelectedIndex > -1 Then
            sindex = ComboBox1.SelectedIndex


        End If


        Dim tfp As New TextFieldParser("C:\Users\track\Desktop\AdvatageFinance\data.csv")
        tfp.Delimiters = New String() {","}
        tfp.TextFieldType = FieldType.Delimited

        If sindex = 0 And textIt <> "" Then
            tfp.ReadLine() ' skip header
            While tfp.EndOfData = False
                Dim fields = tfp.ReadFields()
                iD = fields(0)
                Firstname = fields(1)
                Lastname = fields(2)
                Status = fields(3)
                CreditScore = fields(4)

                If textIt = iD Then
                    ListBox1.Items.Add(String.Format("{0} - {1} - {2} - {3} - {4}", iD, Firstname, Lastname, Status, CreditScore))


                End If
            End While
            If ListBox1.Items.Count = 0 Then
                ListBox1.Items.Add("No Record Found")

            End If
        End If

        If sindex = 1 And textIt <> "" Then
            tfp.ReadLine() ' skip header
            While tfp.EndOfData = False
                Dim fields = tfp.ReadFields()
                iD = fields(0)
                Firstname = fields(1)
                Lastname = fields(2)
                Status = fields(3)
                CreditScore = fields(4)

                If textIt = Firstname Then
                    ListBox1.Items.Add(String.Format("{0} - {1} - {2} - {3} - {4}", iD, Firstname, Lastname, Status, CreditScore))


                End If
            End While
            If ListBox1.Items.Count = 0 Then
                ListBox1.Items.Add("No Record Found")

            End If
        End If

        If sindex = 2 And textIt <> "" Then
            tfp.ReadLine() ' skip header
            While tfp.EndOfData = False
                Dim fields = tfp.ReadFields()
                iD = fields(0)
                Firstname = fields(1)
                Lastname = fields(2)
                Status = fields(3)
                CreditScore = fields(4)

                If textIt = Lastname Then
                    ListBox1.Items.Add(String.Format("{0} - {1} - {2} - {3} - {4}", iD, Firstname, Lastname, Status, CreditScore))


                End If
            End While

            If ListBox1.Items.Count = 0 Then
                ListBox1.Items.Add("No Record Found")

            End If
        End If

        If sindex = 3 And ComboBox2.SelectedIndex > -1 Then
            If ComboBox2.SelectedIndex = 0 Then
                tfp.ReadLine() ' skip header
                While tfp.EndOfData = False
                    Dim fields = tfp.ReadFields()
                    iD = fields(0)
                    Firstname = fields(1)
                    Lastname = fields(2)
                    Status = fields(3)
                    CreditScore = fields(4)

                    If "Approved" = Status Then
                        ListBox1.Items.Add(String.Format("{0} - {1} - {2} - {3} - {4}", iD, Firstname, Lastname, Status, CreditScore))


                    End If
                End While
                If ListBox1.Items.Count = 0 Then
                    ListBox1.Items.Add("No Record Found")

                End If
            End If

            If ComboBox2.SelectedIndex = 1 Then
                tfp.ReadLine() ' skip header
                While tfp.EndOfData = False
                    Dim fields = tfp.ReadFields()
                    iD = fields(0)
                    Firstname = fields(1)
                    Lastname = fields(2)
                    Status = fields(3)
                    CreditScore = fields(4)

                    If "Declined" = Status Then
                        ListBox1.Items.Add(String.Format("{0} - {1} - {2} - {3} - {4}", iD, Firstname, Lastname, Status, CreditScore))


                    End If
                End While
                If ListBox1.Items.Count = 0 Then
                    ListBox1.Items.Add("No Record Found")

                End If
            End If
        End If

        If sindex = 4 And textIt <> "" Then
            tfp.ReadLine() ' skip header
            While tfp.EndOfData = False
                Dim fields = tfp.ReadFields()
                iD = fields(0)
                Firstname = fields(1)
                Lastname = fields(2)
                Status = fields(3)
                CreditScore = fields(4)

                If textIt = CreditScore Then
                    ListBox1.Items.Add(String.Format("{0} - {1} - {2} - {3} - {4}", iD, Firstname, Lastname, Status, CreditScore))


                End If
            End While
            If ListBox1.Items.Count = 0 Then
                ListBox1.Items.Add("No Record Found")

            End If
        End If





    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class
