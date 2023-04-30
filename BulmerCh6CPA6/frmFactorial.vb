' Program Name:     Factorial Calculator
' Author:           Mark Bulmer
' Date:             3/27/2022
' Purpose:          The Factorial Calculator program offers numbers 1-12
'                   for the user to select to obtain that numbers' factorial
'                   power displayed in a pop-up box.

Option Strict On

' I couldn't figure out how to do the proper calculations to apply to a factorial and display it in a listbox,
' but I found an example which applied the calculation to a button and so I have applied what I could for maybe
' getting some partial credit.
Public Class frmFactorial

    ' Each button calls a Factorial function which applies the appropriate number to be computed.
    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click

        MessageBox.Show(Factorial("1"))

    End Sub

    Private Sub btn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2.Click

        MessageBox.Show(Factorial("2"))

    End Sub

    Private Sub btn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn3.Click

        MessageBox.Show(Factorial("3"))

    End Sub

    Private Sub btn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn4.Click

        MessageBox.Show(Factorial("4"))

    End Sub

    Private Sub btn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn5.Click

        MessageBox.Show(Factorial("5"))

    End Sub

    Private Sub btn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn6.Click

        MessageBox.Show(Factorial("6"))

    End Sub

    Private Sub btn7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn7.Click

        MessageBox.Show(Factorial("7"))

    End Sub

    Private Sub btn8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn8.Click

        MessageBox.Show(Factorial("8"))

    End Sub

    Private Sub btn9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn9.Click

        MessageBox.Show(Factorial("9"))

    End Sub

    Private Sub btn10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn10.Click

        MessageBox.Show(Factorial("10"))

    End Sub

    Private Sub btn11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn11.Click

        MessageBox.Show(Factorial("11"))

    End Sub

    Private Sub btn12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn12.Click

        MessageBox.Show(Factorial("12"))

    End Sub

    ' I had trouble figuring this out, and even by checking other examples I could find online I had a lot of trouble 
    ' making sense of it. Normally I would edit code to suit my needs and change appropriate variables, but in the interest
    ' of full disclosure I am going to leave this code as-is and attribute it to John Wein of the Microsoft Developer
    ' Network forums (https://social.msdn.microsoft.com/Forums/en-US/bb23b85f-9a88-4dae-a54e-df71cd25b59d/500-factorial-algorithm-?forum=vbgeneral)
    Private Function Factorial(ByVal ofThisStringsValue As String) As String

        Dim answer As String = "1"
        Dim loopEnd As Decimal
        Decimal.TryParse(ofThisStringsValue, loopEnd)
        loopEnd = Int(loopEnd)
        If loopEnd = 0 Then Return "You tried zero or some other text!!"
        If loopEnd = 1 Then Return "1"
        Dim start As Decimal = 2

        Do
            answer = Multiply(answer, start.ToString)
            start += 1
        Loop Until start > loopEnd

        Return answer
    End Function

    Public Function Multiply(ByVal S1 As String, ByVal S2 As String) As String

        Dim A(65535) As Byte
        Dim B(65535) As Byte
        Dim C(65535) As Byte
        Dim I, J, K, IJ, L1, L2 As Int32
        L1 = S1.Length - 1
        For I = 0 To L1
            A(I) = CType(S1.Substring(L1 - I, 1), Byte)
        Next
        L2 = S2.Length - 1
        For J = 0 To L2
            B(J) = CType(S2.Substring(L2 - J, 1), Byte)
        Next
        For I = 0 To L1
            For J = 0 To L2
                IJ = I + J
                C(IJ) += A(I) * B(J)
                K = 0
                While C(IJ + K) > 9
                    C(IJ + 1) += CByte(Fix(C(IJ) / 10))
                    C(IJ) = CByte(C(IJ) Mod 10)
                    K += 1
                End While
            Next
        Next
        Dim S As String = ""
        L1 += L2 + 2
        For I = L1 To 0 Step -1
            S += C(I).ToString
        Next
        'Remove any character "0" 's from the start of the string        
        'until a non-zero numeric character is reached.        
        Dim testChr As String = ""
        Do
            testChr = S.Substring(0, 1)
            If testChr = "0" Then
                S = S.Remove(0, 1)
            Else
                Exit Do
            End If
        Loop
        Return S
    End Function
    Private Sub mnuClear_Click(sender As Object, e As EventArgs) Handles mnuClear.Click

    End Sub

    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        ' The mnuExit click event closes the window and exits the application

        Close()
    End Sub
End Class