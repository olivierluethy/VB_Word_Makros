Private Sub Document_Open()
    ' Get and display the public IP address
    Dim publicIP As String
    publicIP = GetPublicIPAddress()
    
    ' Get and display the MAC address
    Dim macAddress As String
    macAddress = GetMacAddress()
    
    ' Check if both values are retrieved successfully
    Dim message As String
    message = ""
    
    If publicIP = "" Then
        message = "Could not retrieve the public IP address."
    End If
    
    If macAddress = "" Then
        If message <> "" Then
            message = message & vbCrLf ' Add a new line if there was already a message
        End If
        message = message & "Could not retrieve the MAC address."
    End If
    
    ' Display results if both are available, otherwise show error message
    If publicIP <> "" And macAddress <> "" Then
        MsgBox "Your public IP address is: " & publicIP & vbCrLf & _
               "And your MAC address is: " & macAddress
    Else
        MsgBox message
    End If
End Sub

Private Function GetMacAddress() As String
    Dim macAddress As String
    macAddress = ""
    
    Dim nic As Object
    For Each nic In CreateObject("WinMgmts://").ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
        macAddress = nic.macAddress
        Exit For
    Next

    GetMacAddress = macAddress
End Function

Private Function GetPublicIPAddress() As String
    On Error GoTo ErrorHandler
    
    ' Use an external service to get the public IP address
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' URL of a service that returns the public IP address as plain text
    http.Open "GET", "https://api.ipify.org", False
    http.Send
    
    ' Get the response text, which is the IP address
    GetPublicIPAddress = http.responseText
    Exit Function
    
ErrorHandler:
    GetPublicIPAddress = ""
End Function
