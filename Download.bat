Dim URL, SavePath, WinHttpReq

URL = "https://raw.githubusercontent.com/Aleshhh/ucases-s4.github.io/main/Not_A_Virus.sh"
SavePath = "C:\path\to\save\Not_A_Virus.sh"

Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")

WinHttpReq.Open "GET", URL, False
WinHttpReq.Send

If WinHttpReq.Status = 200 Then
    Dim Stream
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1 'Binary
    Stream.Write WinHttpReq.ResponseBody
    Stream.Position = 0
    Stream.SaveToFile SavePath, 2 '2 = Overwrite
    Stream.Close
    MsgBox "Download Complete", vbInformation
Else
    MsgBox "Error downloading file", vbExclamation
End If