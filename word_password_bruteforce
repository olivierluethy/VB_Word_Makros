' Basiscode von: https://stackoverflow.com/questions/30752525/i-forgot-the-password-to-open-a-word-document-how-can-i-retrieve-the-password

Sub PasswordBreakerWord()

    ' Declare Word application and document objects
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim strPath As String
    Dim passAttempt As String

    ' Turn off alerts and screen updating to improve performance and hide operations
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Create a new instance of Word application
    Set WordApp = CreateObject("Word.Application")
    
    ' Define the path to the Word document to be cracked
    strPath = Environ("USERPROFILE") & "\Desktop\blah.docx"

    ' Loop through all combinations of ASCII characters from 31 to 126 (printable characters)
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    Dim i7 As Integer, i8 As Integer, i9 As Integer
    Dim i10 As Integer, i11 As Integer, i12 As Integer

    ' Nested loops to generate all possible 12-character combinations
    For i1 = 31 To 126
        For i2 = 31 To 126
            For i3 = 31 To 126
                For i4 = 31 To 126
                    For i5 = 31 To 126
                        For i6 = 31 To 126
                            For i7 = 31 To 126
                                For i8 = 31 To 126
                                    For i9 = 31 To 126
                                        For i10 = 31 To 126
                                            For i11 = 31 To 126
                                                For i12 = 31 To 126
                                                    ' Generate password attempt from character codes
                                                    passAttempt = Chr(i12) & Chr(i11) & Chr(i10) & Chr(i9) & Chr(i8) & Chr(i7) & Chr(i6) & Chr(i5) & Chr(i4) & Chr(i3) & Chr(i2) & Chr(i1)
                                                    
                                                    ' Debugging print for each attempt (can be removed for performance)
                                                    Debug.Print passAttempt
                                                    
                                                    ' Attempt to open the document with the current password attempt
                                                    On Error Resume Next
                                                    Set WordDoc = WordApp.Documents.Open(strPath, , True, , passAttempt)
                                                    
                                                    ' Check if the password attempt was successful
                                                    If Err = 0 Then
                                                        MsgBox "Password found: " & passAttempt
                                                        Debug.Print "Password found: " & passAttempt
                                                        
                                                        ' Clean up and exit
                                                        WordApp.Quit
                                                        Application.DisplayAlerts = True
                                                        Application.ScreenUpdating = True
                                                        Exit Sub
                                                    End If
                                                    
                                                    ' Reset error handling after a failed attempt
                                                    On Error GoTo 0
                                                Next i12
                                            Next i11
                                        Next i10
                                    Next i9
                                Next i8
                            Next i7
                        Next i6
                    Next i5
                Next i4
            Next i3
        Next i2
    Next i1
    
    ' Ensure proper cleanup if the password is not found after all attempts
    WordApp.Quit
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
