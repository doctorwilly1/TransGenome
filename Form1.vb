
Public Class Form1
    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
       
    End Sub
    Function StrToBinary(ByVal inputStr As String) As String

        Dim lValue As Integer
        Dim BinaryArr() As String
        Dim BinaryStr As String = ""  'This is fine 
        Dim i As Integer
        Dim LeadingZeroesCount As String
        Dim LeadingZeroes As String
        'Dim Chars = RichTextBox1.Text.Substring(0, 4)
        'Rich Text Box will not take a Null value. Will always use zero
        inputStr = RichTextBox1.Text
        If RichTextBox1.Text = "" Then
            RichTextBox1.Text = "0"
        Else
            'Translates text to ASCII and stores value in 1Value
            For Each ch As Char In RichTextBox1.Text
                lValue = Asc(inputStr)
                i = 0
                'Calls values from ASCII array by using i as index then converts ASCII to binary by adjusting binary array size using ReDim Preserve.
                'ReDim Preserve stores each bit
                ReDim BinaryArr(i)
                While lValue <> 0
                    ReDim Preserve BinaryArr(i)
                    BinaryArr(i) = lValue Mod 2

                    lValue = lValue \ 2
                    i = i + 1
                End While

                'Binary string is determined 
                If UBound(BinaryArr) >= 0 Then
                    For i = 0 To UBound(BinaryArr)
                        BinaryStr = BinaryArr(i) & BinaryStr
                    Next
                    'Determines how many zeroes should preceed the binary outputs' significant digits. 
                    'EX. >>110010(BinaryStr)<< only has SIX digits (Len(BinaryStr)) but it needs EIGHT (Total bits in a byte). TWO zeroes are missing (LeadingZeroesCount). 
                    LeadingZeroesCount = 8 - Len(BinaryStr)
                    'LeadingZeroesCount begins to count down to 0 from 2 (LeadingZeroesCount) in increments of 1.
                    'After each increment a zero is added to the LeadingZeroes string...."00" (LeadingZeroes)
                    For x As Integer = LeadingZeroesCount To 0 Step -1
                        If x > 0 Then
                            LeadingZeroes = "0" & LeadingZeroes
                        End If
                    Next
                    'Takes LeadingZeroes and BinaryStr and combines them to redefine BinaryStr.
                    'LeadingZeroes (00) & BinaryStr (110010) = 00110010
                    If LeadingZeroesCount > 0 Then
                        BinaryStr = String.Format(LeadingZeroes) & BinaryStr
                    End If
                End If
                'Value of BinaryStr is converted to StrToBinary which can then be called as a function 
                StrToBinary = BinaryStr
            Next
        End If
        TestLabel.Text = leadingzeroes
    End Function



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim TestBinary As String
        TestBinary = StrToBinary(TestBinary)
        BinaryLabel.Text = TestBinary
    End Sub
End Class
