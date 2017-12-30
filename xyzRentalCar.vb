Option Strict On
Public Class xyzRentalCar
    '=============================================================================================================
    ' Author:      Holly Eaton
    ' 
    ' Program:     XYZ Rental Car
    ' 
    ' Dev Env:     Visual Studio
    ' 
    ' Description:
    '  Purpose:    Project that will determine:
    '                   The number of miles the car has been driven.
    '                   Rental charges for the customer.
    '  
    '  Input:      Customer Name, Beginning Odometer, Ending Odometer, Number of Days Rented. 
    '
    '  Process:    Calculate the following:
    '                   The number of miles driven.
    '                   The customer rental charge.
    '
    '  Output:     Textual information for the user inside the labels (totals) and textboxes (name, Beginning Odometer,
    '              Ending Odometer, Number of Days Rented)
    '              Format (as numbers) and display miles driven in the appropriate label.              
    '              Format (as currency) and display the customer charges inside the appropriate label.
    ' 
    '===============================================================================================================
    ' 	Declared Constants:
    '	dblDAILY_RENTAL_RATE
    '	dblMILEAGE_OVER_200
    '	dblMILEAGE_OVER_200_CUTOFF
    '
    '===============================================================================================================
    '	Variables for user entered data:
    '	strCustName
    '	dblBegOdometer
    '	dblEndOdometer
    '	dblNumDays
    '	Note: (dim and cast both directly to double then no type casting necessary in calculations)
    '	Example: Dim dblBegOdometer As Double = Cdbl(txtBegOdometer.Text)
    '
    '===============================================================================================================
    '	Variables for calculated values:
    '	dblNumMiles
    '	dblCharges
    '
    '================================================================================================================
    '	Calculations in pseudocode:
    '	Set option strict on at top
    '   Charges= (NumDays)* (DAILY_ RENTAL_RATE) + (((numMiles-200) * MILEAGE _OVER_200) if applicable)
    '	NumMiles = EndOdometer - BegOdometer
    '
    '==================================================================================================================
    '==================================================================================================================
    '==================================================================================================================

    'Declared Constants:
    Const dblDAILY_RENTAL_RATE = 49.95
    Const dblMILEAGE_OVER_200 = 0.32
    Const dblMILEAGE_OVER_200_CUTOFF = 200

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        'Variables for calculated values:
        Dim dblNumMiles As Double
        Dim dblCharges As Double

        'Variables for user entered data validation of user input:
        If txtCustName.Text <> String.Empty Then
            Dim strCustName As String = CStr(txtCustName.Text)
            Try
                Dim dblBegOdometer As Double = CDbl(txtBegOdometer.Text)
                Try
                    Dim dblEndOdometer As Double = CDbl(txtEndOdometer.Text)
                    Try
                        Dim dblNumDays As Double = CDbl(txtNumDays.Text)

                        'Calculations for the Number of Miles and Charges more user input validation.
                        ' Number of Miles
                        If dblBegOdometer < dblEndOdometer Then
                            dblNumMiles = dblEndOdometer - dblBegOdometer
                            If dblBegOdometer > 0 Then
                                If dblEndOdometer > 0 Then
                                    If dblNumDays > 0 Then

                                        'Charges
                                        If dblNumMiles > dblMILEAGE_OVER_200_CUTOFF Then
                                            dblCharges = (dblNumDays * dblDAILY_RENTAL_RATE) + ((dblNumMiles - dblMILEAGE_OVER_200_CUTOFF) * dblMILEAGE_OVER_200)
                                        Else
                                            dblCharges = dblNumDays * dblDAILY_RENTAL_RATE
                                        End If

                                        'format and output the results:
                                        lblNumMiles.Text = dblNumMiles.ToString("n")
                                        lblCharges.Text = dblCharges.ToString("n")

                                        'Send focus to the Clear button
                                        btnClear.Focus()
                                    Else
                                        MessageBox.Show("The NUMBER DAYS value must be at least 1")
                                        'Reset input area
                                        txtNumDays.Text = String.Empty
                                        'Put insertion point inside of Hours Rented textbox.
                                        txtNumDays.Focus()
                                    End If

                                Else
                                    MessageBox.Show("The ENDING odometer value must be a positive number.")
                                    'Reset input area
                                    txtEndOdometer.Text = String.Empty
                                    'Put insertion point inside of Hours Rented textbox.
                                    txtEndOdometer.Focus()
                                End If

                            Else
                                MessageBox.Show("The BEGINNING odometer value must be a positive number.")
                                'Reset input area
                                txtBegOdometer.Text = String.Empty
                                'Put insertion point inside of Hours Rented textbox.
                                txtBegOdometer.Focus()
                            End If

                        Else
                            MessageBox.Show("The Ending mileage must be larger than the beginning mileage. Please check the values and try again.")
                            'Reset input area
                            txtBegOdometer.Text = String.Empty
                            txtEndOdometer.Text = String.Empty
                            'Put insertion point inside of Hours Rented textbox.
                            txtBegOdometer.Focus()

                        End If

                    Catch ex As Exception
                        'What to do if user data entered into txtNumDays is invalid and cannot be cast to dbl.
                        'Tell user what to enter
                        MessageBox.Show("Please enter the NUMBER OF DAYS the car was rented using numeric characters.")
                        'Reset input area
                        txtNumDays.Text = String.Empty
                        'Put insertion point inside of Hours Rented textbox.
                        txtNumDays.Focus()
                    End Try

                Catch ex As Exception
                    'What to do if user data entered into txtEndOdometer is invalid and cannot be cast to dbl.
                    'Tell user what to enter
                    MessageBox.Show("Please enter the ENDING odometer reading using numeric characters.")
                    'Reset input area
                    txtEndOdometer.Text = String.Empty
                    'Put insertion point inside of Hours Rented textbox.
                    txtEndOdometer.Focus()
                End Try

            Catch ex As Exception
                'What to do if user data entered into txtBegOdometer is invalid and cannot be cast to dbl.
                'Tell user what to enter
                MessageBox.Show("Please enter the BEGINNING odometer reading using numeric characters.")
                'Reset input area
                txtBegOdometer.Text = String.Empty
                'Put insertion point inside of Hours Rented textbox.
                txtBegOdometer.Focus()
            End Try
        Else
            MessageBox.Show("Please enter the customers name.")
            txtCustName.Focus()
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear the textboxes and labels
        txtCustName.Text = String.Empty
        txtBegOdometer.Text = String.Empty
        txtEndOdometer.Text = String.Empty
        txtNumDays.Text = String.Empty
        lblNumMiles.Text = String.Empty
        lblCharges.Text = String.Empty

        'Give the focus to txtCustName
        txtCustName.Focus()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Close the form
        Me.Close()
    End Sub
End Class
