<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Function ConvertToUnixTimeStamp(input_datetime) 
    Dim d
    d = CDate(input_datetime) 
    ConvertToUnixTimeStamp = CStr(DateDiff("s", "01/01/1970 00:00:00", d)) 
End Function

Dim objUpload

Set objUpload = Server.CreateObject("Dundas.Upload.2")
objUpload.UseUniqueNames = true
objUpload.MaxFileSize = 500000000
objUpload.UseVirtualDir = True
objUpload.SaveToMemory

nameFile = ConvertToUnixTimeStamp(now())

for each UploadedFile in objUpload.Files
    extensao = right(objUpload.GetFileName(UploadedFile.Originalpath),3)
    arquivo = nameFile & "." & extensao
    UploadedFile.SaveAs "./storage/"&arquivo&""
next

Set objUpload = Nothing
Response.Write "http://storage/"&arquivo
%>