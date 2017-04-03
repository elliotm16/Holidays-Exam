Imports System.IO

Public Class Customers

    ' Sructure for customer's data
    Private Structure Customer

        Public CustomerID As String ' Used to uniquely identify a customer
        Public FirstName As String
        Public Surname As String
        Public EmailAddress As String
        Public PhoneNumber As String

    End Structure

    Private Sub Customers_Load() Handles MyBase.Load

        ' If there is no text file with this name
        If Dir$("Customers.txt") = "" Then

            Dim sw As New StreamWriter("Customers.txt", True)

            ' Write this to it
            sw.WriteLine("0")

            ' StreamWriter needs to be closed
            sw.Close()

            ' Give a warning that a new file has been created
            MsgBox("A new file has been created", vbExclamation, "Warning!")

        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        ' New customer created
        Dim CustomerData As New Customer

        Dim CustomersData() As String = File.ReadAllLines(Dir$("Customers.txt"))

        ' Boolean to indicate whether the data passes validation, set to false by default
        Dim Validated As Boolean
        Validated = False

        ' Sub used for generating the customer ID
        GenerateID(CustomersData)

        If Validation() = True Then

            Dim sw As New System.IO.StreamWriter("Customers.txt", True)

            ' Data in the textboxes is stored in the structure
            CustomerData.CustomerID = LSet(txtCustomerID.Text, 4)
            CustomerData.FirstName = LSet(txtFirstName.Text, 30)
            CustomerData.Surname = LSet(txtSurname.Text, 30)
            CustomerData.EmailAddress = LSet(txtEmailAddress.Text, 30)
            CustomerData.PhoneNumber = LSet(txtPhoneNumber.Text, 11)

            ' Write the data to the text file
            sw.WriteLine(CustomerData.CustomerID & CustomerData.FirstName & CustomerData.Surname & CustomerData.EmailAddress & CustomerData.PhoneNumber)

            ' StreamWriter needs to be closed
            sw.Close()

            ' Output that the file has been saved
            MsgBox("The customer data has been saved. Customer ID: " & txtCustomerID.Text)

        End If

    End Sub

    Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click

        Dim CustomerData() As String = File.ReadAllLines("Customers.txt")

        ' Used to indicate whether a customer has been found
        Dim CustomerFound As Boolean
        CustomerFound = False

        ' Used to count how many customers have been found
        Dim CustomerCount As Integer
        CustomerCount = 0

        ' Used so that if one customer is found, their details can be saved for outputting later
        Dim FoundCustomer As Integer

        ' i starts off at zero and incremements until it reaches the upper bound of customer data
        For i = 0 To UBound(CustomerData)

            ' If the customer is searching a customer ID
            If txtCustomerID.Text <> "" Then

                ' If an ID in the text file matches the ID in the textbox
                If Trim(Mid(CustomerData(i), 1, 4)) = txtCustomerID.Text Then

                    ' Output that a customer has been found
                    MsgBox("A customer with this Customer ID has been found.")

                    ' A customer has been found
                    CustomerFound = True

                    ' Output the customer's data to the textboxes
                    txtCustomerID.Text = Trim(Mid(CustomerData(i), 1, 4))
                    txtFirstName.Text = Trim(Mid(CustomerData(i), 5, 30))
                    txtSurname.Text = Trim(Mid(CustomerData(i), 35, 30))
                    txtEmailAddress.Text = Trim(Mid(CustomerData(i), 65, 30))
                    txtPhoneNumber.Text = Trim(Mid(CustomerData(i), 95, 11))

                End If

            Else

                ' If the data in the text file match the textboxes
                If (Trim(Mid(CustomerData(i), 5, 30)) = txtFirstName.Text Or txtFirstName.Text = "") And (Trim(Mid(CustomerData(i), 35, 30)) = txtSurname.Text Or txtSurname.Text = "") And (Trim(Mid(CustomerData(i), 65, 30)) = txtEmailAddress.Text Or txtEmailAddress.Text = "") And (Trim(Mid(CustomerData(i), 95, 11)) = txtPhoneNumber.Text Or txtPhoneNumber.Text = "") Then

                    ' A customer has been found, and the data is stored in the variable
                    FoundCustomer = i

                    ' Incrememnt number of customers
                    CustomerCount = CustomerCount + 1

                End If

            End If

        Next i

        ' If the user is not searching for a customer ID
        If txtCustomerID.Text = "" Then

            ' If only one customer was found
            If CustomerCount = 1 Then

                ' Output that one customer was found
                MsgBox("One customer was found.")

                ' Output the customer's data to the textboxes
                txtCustomerID.Text = Trim(Mid(CustomerData(FoundCustomer), 1, 4))
                txtFirstName.Text = Trim(Mid(CustomerData(FoundCustomer), 5, 30))
                txtSurname.Text = Trim(Mid(CustomerData(FoundCustomer), 35, 30))
                txtEmailAddress.Text = Trim(Mid(CustomerData(FoundCustomer), 65, 30))
                txtPhoneNumber.Text = Trim(Mid(CustomerData(FoundCustomer), 95, 11))

                Exit Sub

            Else

                ' Output the number of customers found
                MsgBox("There were " & CustomerCount & " customers found.")

            End If

        End If

        ' If the user is searching for customer ID and a customer has not been found
        If txtCustomerID.Text <> "" And CustomerFound = False Then

            ' Output that a customer with the ID entered doesn't exist
            MsgBox("A customer with this Customer ID has not been found.")

            ' Clear the textboxes so that data can be re-entered
            ClearTextboxes()

        End If

    End Sub

    Private Sub GenerateID(CustomersData)

        ' The value of the current highest customer ID
        Dim CurrentCustomerID As Integer

        ' i starts off at zero and incremements until it reaches the upper bound of customer data
        For i = 0 To UBound(CustomersData)

            ' If the highest ID is equal to the current customer ID, add one to it
            If Val(Trim(Mid(CustomersData(i), 1, 4))) = CurrentCustomerID Then

                ' Increment the value of the current customer ID
                CurrentCustomerID = CurrentCustomerID + 1

            End If

        Next

        ' THE ID is then saved to the textbox for the user to view
        txtCustomerID.Text = LSet(CurrentCustomerID, 4)

    End Sub

    Private Function Validation()

        If txtFirstName.Text = "" Or txtSurname.Text = "" Or txtEmailAddress.Text = "" Or txtPhoneNumber.Text = "" Then

            ' Presence check
            MsgBox("You must enter something for first and last name, email address and phone number.")

            Return False

        ElseIf IsNumeric(txtPhoneNumber.Text) = False Then

            ' Type check
            MsgBox("Phone number must be 11 numbers.")

            Return False

        ElseIf ((txtEmailAddress.Text).Contains("@") = False) Or ((txtEmailAddress.Text).Contains(".") = False) Then

            ' Format check
            MsgBox("Email address must contain an '@' symbol and a '.' symbol.")

            Return False

        ElseIf ((txtFirstName.Text).Length > 30 Or (txtSurname.Text).Length > 30) Then

            ' Length check
            MsgBox("Fist Name and Surname must be 30 characters maximum.")

        Else

            Return True

        End If

    End Function

    Private Sub ClearTextboxes()

        ' Clears all the textboxes
        txtCustomerID.Text = ""
        txtFirstName.Text = ""
        txtSurname.Text = ""
        txtEmailAddress.Text = ""
        txtPhoneNumber.Text = ""

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click

        ' Calls the sub to clear the textboxes
        ClearTextboxes()

    End Sub

End Class