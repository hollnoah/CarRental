'Noah Holloway
'Spring 2025
'RCET 2265
'Car Rental
'https://github.com/hollnoah/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm

    ' Summary totals
    Dim totalCustomers As Integer = 0
    Dim totalMiles As Double = 0
    Dim totalCharges As Decimal = 0D

    ' Set default values for form
    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        MilesradioButton.Checked = True
        NameTextBox.Focus()
    End Sub

    ' Validate inputs
    Function InputsAreValid() As Boolean
        Dim message As String = ""
        Dim valid As Boolean = True

        If NameTextBox.Text = "" Then
            message += "Customer Name is required." & vbCrLf
            valid = False
        End If
        If AddressTextBox.Text = "" Then
            message += "Address is required." & vbCrLf
            valid = False
        End If
        If CityTextBox.Text = "" Then
            message += "City is required." & vbCrLf
            valid = False
        End If
        If StateTextBox.Text = "" Then
            message += "State is required." & vbCrLf
            valid = False
        End If
        If ZipCodeTextBox.Text = "" Then
            message += "Zip Code is required." & vbCrLf
            valid = False
        End If

        Dim beginOdometer As Double
        If Not Double.TryParse(BeginOdometerTextBox.Text, beginOdometer) Then
            message += "Beginning Odometer must be a number." & vbCrLf
            valid = False
        End If

        Dim endOdometer As Double
        If Not Double.TryParse(EndOdometerTextBox.Text, endOdometer) Then
            message += "Ending Odometer must be a number." & vbCrLf
            valid = False
        End If

        If valid AndAlso beginOdometer >= endOdometer Then
            message += "Beginning Odometer must be less than Ending Odometer." & vbCrLf
            valid = False
        End If

        Dim days As Integer
        If Not Integer.TryParse(DaysTextBox.Text, days) Then
            message += "Number of Days must be a whole number." & vbCrLf
            valid = False
        ElseIf days <= 0 OrElse days > 45 Then
            message += "Number of Days must be between 1 and 45." & vbCrLf
            valid = False
        End If

        If Not valid Then
            MsgBox(message)
        End If

        Return valid
    End Function

    ' Calculate button
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If InputsAreValid() Then
            Dim beginOdometer As Double = Convert.ToDouble(BeginOdometerTextBox.Text)
            Dim endOdometer As Double = Convert.ToDouble(EndOdometerTextBox.Text)
            Dim days As Integer = Convert.ToInt32(DaysTextBox.Text)
            Dim distance As Double = endOdometer - beginOdometer

            If KilometersradioButton.Checked Then
                distance = distance * 0.62
            End If

            Dim mileageCharge As Decimal = 0D
            If distance > 200 AndAlso distance <= 500 Then
                mileageCharge = CDec((distance - 200) * 0.12)
            ElseIf distance > 500 Then
                mileageCharge = CDec((300 * 0.12) + ((distance - 500) * 0.1))
            End If

            Dim dayCharge As Decimal = CDec(days) * 15D
            Dim discountRate As Decimal = 0D

            If AAAcheckbox.Checked Then discountRate += 0.05D
            If Seniorcheckbox.Checked Then discountRate += 0.03D

            Dim discountAmount As Decimal = (mileageCharge + dayCharge) * discountRate
            Dim totalOwed As Decimal = (mileageCharge + dayCharge) - discountAmount

            TotalMilesTextBox.Text = distance.ToString("F2") & " mi"
            MileageChargeTextBox.Text = FormatCurrency(mileageCharge)
            DayChargeTextBox.Text = FormatCurrency(dayCharge)
            TotalDiscountTextBox.Text = FormatCurrency(discountAmount)
            TotalChargeTextBox.Text = FormatCurrency(totalOwed)

            totalCustomers += 1
            totalMiles += distance
            totalCharges += totalOwed

            SummaryButton.Enabled = True
        End If
    End Sub

    ' Clear button
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub

    ' Summary button
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim summary As String = "Total Customers: " & totalCustomers.ToString() & vbCrLf &
                                "Total Miles Driven: " & totalMiles.ToString("F2") & " mi" & vbCrLf &
                                "Total Charges: " & FormatCurrency(totalCharges)
        MsgBox(summary)
    End Sub

    ' Exit button
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub CarRentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetDefaults()
        SummaryButton.Enabled = False
    End Sub

End Class

