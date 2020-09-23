Attribute VB_Name = "Module1"
' XOR Password cracker
' Copyright by Robin Schuil, 2001
' E-mail: beast@valleyalley.co.uk

Function XorCrypt(PlainText As String, Password As String) As String

    Dim PwdPos As Integer
    Dim MsgPos As Integer
    Dim CipherText As String
    
    ' Preserve buffer
    CipherText = Space(Len(PlainText))
    
    ' Encode / decode message by XOR'ing the data with the password
    For MsgPos = 1 To Len(PlainText)
        PwdPos = (PwdPos Mod Len(Password)) + 1
        Mid(CipherText, MsgPos, 1) = Chr(Asc(Mid(PlainText, MsgPos, 1)) Xor Asc(Mid(Password, PwdPos, 1)))
    Next MsgPos

    ' Return the result
    XorCrypt = CipherText

End Function

Function GetKeyLength(CipherText As String) As Integer

    Dim i As Integer, maxI As Integer
    Dim j As Integer, pos As Integer
    Dim Count As Double, maxCount As Double
    Dim ResultString As String
    Dim TestByte As Byte
    
    Dim pct As Double
    Dim bestI As Integer
    
    For i = 1 To Len(CipherText) - 1
        
        pos = i: Count = 0
        
        ' Xor the message with shifted message
        For j = 1 To Len(CipherText)
            pos = pos + 1: If pos > Len(CipherText) Then pos = 1
            TestByte = Asc(Mid(CipherText, j, 1)) Xor Asc(Mid(CipherText, pos, 1))
            If TestByte = 0 Then Count = Count + 1
        Next j
        
        ' Calculate percentage of 0-value bytes in message
        pct = Count / Len(CipherText)
        
        ' If percentage is larger then 0.5% of the message, we found the keylength
        If pct > 0.05 Then
            GetKeyLength = i
            Exit Function
        End If
        
    Next i
    
    MsgBox "Error: cannot determine keylength."
    End
    
End Function

Sub XorCrack(CipherText As String)

    Dim BestChar As Byte
    Dim BestCount As Double

    Dim Password As String
    Dim PwdLength As Integer
    Dim PwdPosition As Integer
    Dim PwdChar As Byte
    Dim MsgPosition As Integer
    Dim DecodedChar As Byte
    
    Dim CurCount As Double
        
    ' Get the keylength
    PwdLength = GetKeyLength(CipherText)
    
    ' Put dots in the label on the form
    Form1.lblPassword.Caption = String(PwdLength, ".")
    Form1.lblPassword.Refresh
    
    ' Start at the first character
    PwdPosition = 1
    
    Do
            
        BestCount = 0
    
        ' Try each character
        For PwdChar = 1 To 254
                
            ' Update label for animation
            ' To gain more speed, you should remove this.
            Form1.lblPassword.Caption = Password & Chr(PwdChar) & String(PwdLength - PwdPosition, ".")
            Form1.lblPassword.Refresh
            DoEvents
                
            ' Reset counter
            CurCount = 0
    
            ' Decode the message and calculate totals
            For MsgPosition = PwdPosition To Len(CipherText) Step PwdLength
            
                DecodedChar = Asc(Mid(CipherText, MsgPosition, 1)) Xor PwdChar
                
                ' Frequency table for english text
                Select Case Chr(DecodedChar)
                    Case "E", "e"
                        CurCount = CurCount + (100 * 0.127)
                    Case "T", "t"
                        CurCount = CurCount + (100 * 0.097)
                    Case "I", "i"
                        CurCount = CurCount + (100 * 0.075)
                    Case "A", "a"
                        CurCount = CurCount + (100 * 0.073)
                    Case "O", "o"
                        CurCount = CurCount + (100 * 0.068)
                    Case "N", "n"
                        CurCount = CurCount + (100 * 0.067)
                    Case "S", "s"
                        CurCount = CurCount + (100 * 0.067)
                    Case "R", "r"
                        CurCount = CurCount + (100 * 0.064)
                    Case "H", "h"
                        CurCount = CurCount + (100 * 0.049)
                    Case "C", "c"
                        CurCount = CurCount + (100 * 0.045)
                    Case "L", "l"
                        CurCount = CurCount + (100 * 0.04)
                    Case " "
                        CurCount = CurCount + (100 * 0.038)
                    Case "D", "d"
                        CurCount = CurCount + (100 * 0.031)
                    Case "P", "p"
                        CurCount = CurCount + (100 * 0.03)
                    Case "Y", "y"
                        CurCount = CurCount + (100 * 0.027)
                    Case "U", "u"
                        CurCount = CurCount + (100 * 0.024)
                    Case "M", "m"
                        CurCount = CurCount + (100 * 0.024)
                    Case "F", "f"
                        CurCount = CurCount + (100 * 0.021)
                    Case "B", "b"
                        CurCount = CurCount + (100 * 0.017)
                    Case "G", "g"
                        CurCount = CurCount + (100 * 0.016)
                    Case "W", "w"
                        CurCount = CurCount + (100 * 0.013)
                    Case "V", "v"
                        CurCount = CurCount + (100 * 0.008)
                    Case "K", "k"
                        CurCount = CurCount + (100 * 0.008)
                    Case "X", "x"
                        CurCount = CurCount + (100 * 0.005)
                    Case "Q", "q"
                        CurCount = CurCount + (100 * 0.002)
                    Case "Z", "z"
                        CurCount = CurCount + (100 * 0.001)
                    Case "J", "j"
                        CurCount = CurCount + (100 * 0.001)
                End Select
                            
            Next MsgPosition
        
            ' If total is highest seen, remember character
            If CurCount > BestCount Then
                BestCount = CurCount
                BestChar = PwdChar
            End If
                
        Next
    
        ' Add highest scoring character to the password
        Password = Password & Chr(BestChar)
        
        ' Process next password character
        PwdPosition = PwdPosition + 1
            
    Loop Until PwdPosition > PwdLength
    
    ' Display password in label
    Form1.lblPassword = Shorten(Password)
    Form1.lblPassword.Refresh
    
    'MsgBox XorCrypt(CipherText, Password)
    
End Sub

Public Function Shorten(Password As String) As String
    ' This function shortens a password .
    ' It may happen that the password found is like 'secretsecret'
    ' This function then returns only 'secret' as the password.
    Dim PwdLength As Double
    Dim TestVal As Integer
    TestVal = Len(Password)
    While TestVal > 0
        PwdLength = Len(Password) / TestVal
        If PwdLength = Int(PwdLength) Then
            If Mid(Password, 1, PwdLength) = Mid(Password, 1 + PwdLength, PwdLength) Then
                Shorten = Mid(Password, 1, PwdLength)
                Exit Function
            End If
        End If
        TestVal = TestVal - 1
    Wend
    Shorten = Password
End Function
