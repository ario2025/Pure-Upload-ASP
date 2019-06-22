<%@ LANGUAGE="VBSCRIPT" CODEPAGE="1252" ENABLESESSIONSTATE="FALSE" LCID="1046" %>
<!--#include file="upload.lib.asp"-->
<% Response.Charset = "ISO-8859-1"

Dim Form : Set Form = New ASPForm
Server.ScriptTimeout = 1440 ' Limite de 24 minutos de execução de código, o upload deve acontecer dentro deste tempo ou então ocorre erro de limite de tempo.
Const MaxFileSize = 10240000 ' Bytes. Aqui está configurado o limite de 100 MB por upload (inclui todos os tamanhos de arquivos e conteúdos dos formulários).
If Form.State = 0 Then

	For each Key in Form.Texts.Keys
		Response.Write "Elemento: " & UTF8_ANSI(Key) & " = " & UTF8_ANSI(Form.Texts.Item(Key)) & vbcrlf
	Next

	For each Field in Form.Files.Items
		' # Field.Filename : Nome do Arquivo que chegou.
		' # Field.ByteArray : Dados binários do arquivo, útil para subir em blobstore (MySQL).
		Field.SaveAs Server.MapPath(".") & "\upload\" & Field.FileName
		Response.Write "Arquivo: " & UTF8_ANSI(Field.FileName) & " foi salvo com sucesso." & vbcrlf
	Next
End If

function UTF8_ANSI(x)
    ' Check if do you are using the codepage 1252 or this script doesn't works properly.
    ' Verifique se você está usando o código de página 1252 ou este não funcionará corretamente.
    ' <.%@LANGUAGE="VBSCRIPT" CODEPAGE = "1252" %.>
    Cod = second(now()) + minute(now())
    x=replace(x,chr(226)&chr(128)&chr(156),chr(34))
    x=replace(x,chr(226)&chr(128)&chr(157),chr(34))
    x=replace(x,chr(226)&chr(128)&chr(147),chr(150))
    for ife = 1 to 191 : x=replace(x,chr(195)&chr(ife),chr(ife+64)) : next
    UTF8_ANSI=x
end function


%>
