<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" ENABLESESSIONSTATE="FALSE" LCID="1046" %>
<!--#include file="upload.lib.asp"-->
<% Response.Charset = "utf-8"

Dim Form : Set Form = New ASPForm
Server.ScriptTimeout = 1440 ' Limite de 24 minutos de execução de código, o upload deve acontecer dentro deste tempo ou então ocorre erro de limite de tempo.
Const MaxFileSize = 1200000 ' Limite de 1,2 Mb de arquivo.
If Form.State = 0 Then
 For each Field in Form.Files.Items
  ' # Field.Filename : Nome do Arquivo que chegou.
  ' # Field.ByteArray : Dados binários do arquivo, útil para subir em blobstore (MySQL).
  Field.SaveAs Server.MapPath(".") & "\" & Field.FileName
 Next
End If
%>
