Imports System.IO

Public Class Holidays

    Private Structure Holiday
        Public HolidayID As String
        Public HolidayName As String
        Public Location As String
        Public HolidayType As String                  'Creating the structure that will hold the  data.
        Public Rating As String
    End Structure

    Private Sub Holidays_Load() Handles MyBase.Load

        Customers.Show()

        If Dir$("Holidays.txt") = "" Then
            Dim sw As New StreamWriter("Holidays.txt", True)    'This makes sure there is actually a database to enter/read data. If not, it creates a new blank one.
            sw.WriteLine("")
            sw.Close()
            MsgBox("A new file has been created", vbExclamation, "Warning!")
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim HolidayData As New Holiday
        Dim sw As New System.IO.StreamWriter("Holidays.txt", True)
        HolidayData.HolidayID = LSet(txtHolidayID.Text, 50)
        HolidayData.HolidayName = LSet(txtHolidayName.Text, 50)
        HolidayData.HolidayType = LSet(txtHolidayType.Text, 50)
        HolidayData.Location = LSet(txtLocation.Text, 50)                      'Filling the structure with data.
        HolidayData.Rating = LSet(txtRating.Text, 50)

        sw.WriteLine(HolidayData.HolidayID & HolidayData.HolidayName & HolidayData.HolidayType & HolidayData.Location & HolidayData.Rating)
        sw.Close()                                                                  'Always need to close afterwards
        MsgBox("File Saved!")
    End Sub

    Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click
        Dim CountGot As Integer
        CountGot = 0
        Dim HolidayCount As Integer
        HolidayCount = 0

        ' ". . . ." Indicates missing or broken code.

        Dim HolidayData() As String = File.ReadAllLines("Holidays.txt")
        For i = 0 To UBound(HolidayData)

            CountGot = 0
            If Trim(Mid(HolidayData(i), 1, 50)) = txtHolidayID.Text And Not txtHolidayID.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(HolidayData(i), 51, 50)) = txtHolidayName.Text And Not txtHolidayName.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(HolidayData(i), 101, 50)) = txtLocation.Text And Not txtLocation.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(HolidayData(i), 151, 50)) = txtHolidayType.Text And Not txtHolidayType.Text = "" Then CountGot = CountGot + 1 'Counting how many attributes follow the search
            If Trim(Mid(HolidayData(i), 201, 50)) = txtRating.Text And Not txtRating.Text = "" Then CountGot = CountGot + 1
            If CountGot > 0 Then HolidayCount = HolidayCount + 1 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''If attributes match, add to the count.
        Next i
        MsgBox("There were: " & HolidayCount & " Found")

    End Sub

End Class
