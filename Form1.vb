Public Class Form1
    Dim plainTextArray(13) As String
    Dim encryptArray(13) As Integer         'All the arrays are size 13 because log27(max long decimal value)=13
    Dim decryptArray(13) As Char            'One must send messages 13 characters at a time
    Dim number, ch, output As String
    Dim cipherText As Long
    Function Exponent(ByVal m As Long, ByVal exp As Long) As Long
        Dim k As Long
        k = m
        If exp = 0 Then
            m = 1
        Else                              'discrete exponent function computes m^exp 
            For i = 1 To exp - 1          'exp must be a whole number (0 and all positive integers)
                m = (m * k)
            Next
        End If
        Return m
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        number = TextBox1.Text
        TextBox4.Text = ""
        For i = 1 To Len(number)                             'convert the english letters to char values
            ch = Mid(number, i, 1)                          'ch=the i'th character of the string

            plainTextArray(i) = ch
            If (Asc(ch)) = 32 Then                  'converts to space = 0 a = 1 b = 2 ....... z = 26 (residues modulo 27)
                output = 0
                encryptArray(i - 1) = 0
            Else
                output = Asc(ch) - 96
                encryptArray(i - 1) = output
            End If
            TextBox4.Text = TextBox4.Text & " " & output
        Next


    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click    'CALCULATE
        TextBox2.Text = ""
        TextBox5.Text = ""
        cipherText = 0
        For j = 1 To Len(number)
            cipherText = cipherText + encryptArray(j - 1) * Exponent(27, Len(number) - j)  'calculates the ciphertext (encrypt)
        Next
        TextBox2.Text = cipherText

        For i = Len(number) - 1 To 0 Step -1                    'Reads out remainders backward so that the original message appears
            TextBox5.Text = TextBox5.Text & encryptArray(i) & " "
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click  'DECRYPT
        Dim remainder As Long
        Dim letter As String
        remainder = 0
        TextBox3.Text = ""
        cipherText = TextBox2.Text
        For i = 0 To 13
            remainder = cipherText Mod 27             'takes remainder
            cipherText = cipherText \ 27              'divides and rounds down
            If remainder = 0 Then
                decryptArray(i) = " "
            Else
                decryptArray(i) = ChrW(remainder + 96)     'Stores all the remainders as their corresponding letter
            End If
        Next
        For i = 13 To 0 Step -1
            letter = decryptArray(i)
            TextBox3.Text = TextBox3.Text & letter
        Next
    End Sub

End Class

