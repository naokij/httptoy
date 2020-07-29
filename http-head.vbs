Dim sUrl
sUrl = InputBox("Please input url")
Dim WinHttpReq
Set WinHttpReq = WScript.CreateObject("WinHttp.WinHttpRequest.5.1")
WinHttpReq.Open "GET", sUrl
WinHttpReq.Send
Dim output
output = WinHttpReq.Status & " - " + WinHttpReq.StatusText & vbCrLf + WinHttpReq.GetAllResponseHeaders()
WScript.Echo(output)
