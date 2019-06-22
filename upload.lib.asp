<%
Const adTypeBinary = 1
Const adTypeText = 2
Const xfsCompleted    = &H0 '0  Formulário foi processado com sucesso.
Const xfsNotPost      = &H1 '1  Método de envio não é POST.
Const xfsZeroLength   = &H2 '2  Conteúdo de tamanho zero (não há conteúdo no formulário de origem)
Const xfsInProgress   = &H3 '3  Formulário está no meio do processo.
Const xfsNone         = &H5 '5  Estado inicial do formulário.
Const xfsError        = &HA '10  Se houve algum erro.  
Const xfsNoBoundary   = &HB '11  Boundary dos arquivos em formulário multipart/form-data não foi especificado.
Const xfsUnknownType  = &HC '12  Formulário desconhecido (Content-type necessita ser application/x-www-form-urlencoded ou multipart/form-data) 
Const xfsSizeLimit    = &HD '13  O tamanho excede o limite permitido.
Const xfsTimeOut      = &HE '14  O tempo excede o limite permitido.
Const xfsNoConnected  = &HF '15  O cliente desconectou antes de completar o upload.
Const xfsErrorBinaryRead = &H10 '16  Erro inexperado ocorreu (Erro ASP).
Const MaxLicensedLimit = &H77359400 ' Tamanho máximo do upload para qualquer situação = 2 GB (limite do IIS).
Class ASPForm
 Public ChunkReadSize, BytesRead, TotalBytes, UploadID
 Private m_ReadTime
 Public TempPath, MaxMemoryStorage, CharSet, FormType, SourceData, ReadTimeout
 public Default Property Get Item(Key)
  Set Item = m_Items.Item(Key)
 End Property
 public Property Get Items
  Read
  Set Items = m_Items
 End Property
 public Property Get Files
  Read
  Set Files = m_Items.Files
 End Property
 public Property Get Texts
  Read
  Set Texts = m_Items.Texts
 End Property
 public Property Get NewUploadID
  Randomize
  NewUploadID = clng(rnd * &H7FFFFFFF)
 End Property
 Public Property Get ReadTime
  if isempty(m_ReadTime) then
   if not isempty(StartUploadTime) then ReadTime = Clng((Now() - StartUploadTime) * 86400 * 1000)
  else
   ReadTime = m_ReadTime
  end if
 End Property
 Public Property Get State
  if m_State = xfsNone Then Read
  State = m_State
 End Property
 Private Function CheckRequestProperties
   If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then 'Request method must be "POST"
   m_State = xfsNotPost
   Exit Function
  End If 
  Dim CT
  CT = Request.ServerVariables("HTTP_Content_Type")
  if len(CT) = 0 then CT = Request.ServerVariables("CONTENT_TYPE")
   If LCase(Left(CT, 19)) <> "multipart/form-data" Then
   m_State = xfsUnknownType 
   Exit Function
  End If
  Dim PosB
  PosB = InStr(LCase(CT), "boundary=")
  If PosB = 0 Then
   m_State = xfsNoBoundary
   Exit Function
  End If
  If PosB > 0 Then Boundary = Mid(CT, PosB + 9)
  PosB = InStr(LCase(CT), "boundary=") 
  If PosB > 0 then
   PosB = InStr(Boundary, ",")
   If PosB > 0 Then Boundary = Left(Boundary, PosB - 1)
  end if
  On Error Resume next
  TotalBytes = Request.TotalBytes
  If Err<>0 Then
   TotalBytes = CLng(Request.ServerVariables("HTTP_Content_Length"))
   if len(TotalBytes)=0 then TotalBytes = CLng(Request.ServerVariables("CONTENT_LENGTH"))
  End If
  If TotalBytes = 0 then
   m_State = xfsZeroLength 
   Exit Function
  End If
  If IsInSizeLimit(TotalBytes) Then
   CheckRequestProperties = True
   m_State = xfsInProgress 
  Else
   m_State = xfsSizeLimit 
  End if
 End Function
 Public Sub Read()
  if m_State <> xfsNone Then Exit Sub
  If Not CheckRequestProperties Then 
   WriteProgressInfo
   Exit Sub
  End If
  if isempty(bSourceData) then Set bSourceData = createobject("ADODB.Stream")
  bSourceData.Open
  bSourceData.Type = 1
  Dim DataPart, PartSize
  BytesRead = 0
  StartUploadTime = Now
  Do While BytesRead < TotalBytes
   PartSize = ChunkReadSize
   if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
   DataPart = Request.BinaryRead(PartSize)
   BytesRead = BytesRead + PartSize
   bSourceData.Write DataPart
   WriteProgressInfo
   If Not Response.IsClientConnected Then
    m_State = xfsNoConnected  
    Exit Sub
   End If
  Loop
  m_State = xfsCompleted
  ParseFormData
 End Sub
 Private Sub ParseFormData
  Dim Binary
  bSourceData.Position = 0
  Binary = bSourceData.Read
  m_Items.mpSeparateFields Binary, Boundary
 End Sub
 Public Function getForm(FormID)
  if isempty(ProgressFile.UploadID) Then
   ProgressFile.UploadID = FormID
  End If
  Dim ProgressData
  ProgressData = ProgressFile
  if len(ProgressData) > 0 then
   if ProgressData = "DONE" Then
    ProgressFile.Done
    Err.Raise 1, "getForm", "Concluído"
   Else
    ProgressData = Split (ProgressData, vbCrLf)
    if ubound(ProgressData) = 3 Then
     m_State = clng(ProgressData(0))
     TotalBytes = clng(ProgressData(1))
     BytesRead = clng(ProgressData(2))
     m_ReadTime = clng(ProgressData(3))
    End If
   End If
  end if
  Set getForm = Me
 End Function
 Private Sub WriteProgressInfo
  If UploadID > 0 Then
   if isempty(ProgressFile.UploadID) Then
    ProgressFile.UploadID = UploadID
   End If
   Dim ProgressData, FileName
   ProgressData = m_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
   ProgressFile.Contents = ProgressData
  End If
 End Sub
 Private Sub Class_Initialize()
  ChunkReadSize = &H10000 '64 kB
  SizeLimit = &H100000 '1MB
  BytesRead = 0
  m_State = xfsNone
  TotalBytes = Request.TotalBytes
  Set ProgressFile = New cProgressFile
  Set m_Items = New cFormFields
 End Sub
 Private Sub Class_Terminate()
  If UploadID > 0 Then
   ProgressFile.Contents = "DONE"
  End If
 End Sub
 Private Function IsInSizeLimit(TotalBytes)
  IsInSizeLimit = (m_SizeLimit = 0 or m_SizeLimit > TotalBytes) and (MaxLicensedLimit > TotalBytes)
 End Function
  Public Property Get SizeLimit
  SizeLimit = m_SizeLimit
 End Property 
  Public Property Let SizeLimit(NewLimit)
 if NewLimit > MaxLicensedLimit Then
   Err.Raise 1, "Upload grande", "Seu upload ultrapassou o limite máximo suportado por este servidor."
   m_SizeLimit = MaxLicensedLimit
  Else
   m_SizeLimit = NewLimit
  end if
 End Property 
 Public Boundary
 Private m_Items 
 Private m_State
 Private m_SizeLimit 'Define o tamanho limite do formulário.
 Private bSourceData 'ADODB.Stream
 Private StartUploadTime , TempFiolder 
 Private ProgressFile 'Informa o atual progresso de upload do arquivo
End Class
Class cFormFields
 Dim m_Keys()
 Dim m_Items()
 Dim m_Count
 Public Default Property Get Item(Key)
  If vartype(Key) = vbInteger or vartype(Key) = vbLong then
   if Key<1 or Key>m_Count Then Err.raise "Item não encontrado"
   Set Item = m_Items(Key-1)
   Exit Property
  end if
  Dim Count
  Count = ItemCount(Key)
  If Count > 0 then
   If Count>1 Then
    Dim OutItem, ItemCounter
    Set OutItem = New cFormFields
    ItemCounter = 0
    For ItemCounter = 0 To Ubound(m_Keys)
     If m_Keys(ItemCounter) = Key then OutItem.Add Key, m_Items(ItemCounter)
    Next
    Set Item = OutItem
   Else 
    For ItemCounter = 0 To Ubound(m_Keys)
     If m_Keys(ItemCounter) = Key then exit for
    Next
    if isobject (m_Items(ItemCounter)) then
     Set Item = m_Items(ItemCounter)
    else
     Item = m_Items(ItemCounter)
    end if
   End If
  Else
   Set Item = New cFormField
  End if
 End Property
 Public Property Get xA_NewEnum
  Set xA_NewEnum = m_Items
 End Property
 Public Property Get Items()
  Items = m_Items
 End Property
 Public Property Get Keys()
  Keys = m_Keys
 End Property
 public Property Get Files
  Dim cItem, OutItem, ItemCounter
  Set OutItem = New cFormFields 
  ItemCounter = 0
  if m_Count > 0 then
   For ItemCounter = 0 To Ubound(m_Keys)
    Set cItem = m_Items(ItemCounter)
    if cItem.IsFile then
     OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
    end if
   Next
  End If
  Set Files = OutItem 
 End Property
 Public Property Get Texts
  Dim cItem, OutItem, ItemCounter
  Set OutItem = New cFormFields 
  ItemCounter = 0
  For ItemCounter = 0 To Ubound(m_Keys)
   Set cItem = m_Items(ItemCounter)
   if Not cItem.IsFile then
    OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
   end if
  Next
  Set Texts = OutItem 
 End Property
 Public Sub Save(Path)
  Dim Item
  For Each Item In m_Items
   If Item.isFile Then
    Item.Save Path
   End If
  Next
 End Sub
 Public Property Get ItemCount(ByVal Key)
  Dim cKey, Counter
  Counter = 0
  For Each cKey In m_Keys
   If cKey = Key then Counter = Counter + 1
  Next
  ItemCount = Counter
 End Property
 Public Property Get Count()
  Count = m_Count
 End Property
 Public Sub Add(byval Key, Item)
  Key = "" & Key
  ReDim Preserve m_Items(m_Count)
  ReDim Preserve m_Keys(m_Count)
  m_Keys(m_Count) = Key
  Set m_Items(m_Count) = Item
  m_Count = m_Count + 1
 End Sub
 Private Sub Class_Initialize()
  m_Count = 0
 End Sub
 Public Sub mpSeparateFields(Binary, ByVal Boundary)
  Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary
  Boundary = "--" & Boundary   
  Boundary = StringToBinary(Boundary)
  PosOpenBoundary = InStrB(Binary, Boundary)
  PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)
  Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)
   Dim HeaderContent, bFieldContent
   Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type
   Dim TwoCharsAfterEndBoundary
   PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))
   HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
   bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)
   GetHeadFields BinaryToString(HeaderContent), FormFieldName, SourceFileName, Content_Disposition, Content_Type
   Dim Field
   Set Field = New cFormField
   Field.ByteArray = MultiByteToBinary(bFieldContent)
   Field.Name = FormFieldName
   Field.ContentDisposition = Content_Disposition
   if len(SourceFileName) > 0 then
    Field.FilePath = SourceFileName
    Field.FileName = GetFileName(SourceFileName)
   End If
   Field.ContentType = Content_Type
   Add FormFieldName, Field
   TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
   isLastBoundary = TwoCharsAfterEndBoundary = "--"
   If Not isLastBoundary Then 'Final do boundary. Segue para próximo arquivo.
    PosOpenBoundary = PosCloseBoundary
    PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
   End If
  Loop
 End Sub
End Class
Class cProgressFile
 Private fs
 Public TempFolder
 Public m_UploadID
 Public TempFileName
 Public Default Property Get Contents()
  Contents = GetFile(TempFileName)
 End Property
 Public Property Let Contents(inContents)
  WriteFile TempFileName, inContents
 End Property
 Public Sub Done 'Remove arquivo temporário quando alcança sucesso no upload.
  FS.DeleteFile TempFileName
 End Sub
 Public Property Get UploadID()
  UploadID = m_UploadID
 End Property
 Public Property Let UploadID(inUploadID)
  if isempty(FS) then Set fs = CreateObject("Scripting.FileSystemObject")
  TempFolder = fs.GetSpecialFolder(2)
  m_UploadID = inUploadID
  TempFileName = TempFolder & "\temporary" & m_UploadID & ".~tmp"
  Dim DateLastModified
  on error resume next
  DateLastModified = fs.GetFile(TempFileName).DateLastModified
  on error goto 0
  if isempty(DateLastModified) then
  elseif Now-DateLastModified>1 Then
   FS.DeleteFile TempFileName
  end if
 End Property
 Private Function GetFile(Byref FileName)
  Dim InStream
  On Error Resume Next
  Set InStream = fs.OpenTextFile(FileName, 1)
  GetFile = InStream.ReadAll
  On Error Goto 0
 End Function
 Private Function WriteFile(Byref FileName, Byref Contents)
  Dim OutStream
  On Error Resume Next
  Set OutStream = fs.OpenTextFile(FileName, 2, True)
  OutStream.Write Contents
 End Function
 Private Sub Class_Initialize()
 End Sub
End Class
Class cFormField
 Public ContentDisposition, ContentType, FileName, FilePath, Name
 Public ByteArray
 Public CharSet, HexString, InProgress, SourceLength, RAWHeader, Index, ContentTransferEncoding
 Public Default Property Get String()
  String = BinaryToString(ByteArray)
 End Property 
 Public Property Get IsFile()
  IsFile = not isempty(FileName)
 End Property
 Public Property Get Length()
  Length = LenB(ByteArray)
 End Property
 Public Property Get Value()
  Set Value = Me
 End Property
 Public Sub Save(Path)
  if IsFile Then
   Dim fullFileName
   fullFileName = Path & "\" & FileName
   SaveAs fullFileName
  Else
   Err.raise "O campo de texto " & Name & " não tem um nome de arquivo."
  End If
 End Sub
 Public Sub SaveAs(newFileName)
  SaveBinaryData newFileName, ByteArray
 End Sub
End Class
Function StringToBinary(String)
  Dim I, B : For I=1 to len(String) : B = B & ChrB(Asc(Mid(String,I,1))) : Next 
  StringToBinary = B
End Function
Function BinaryToString(Binary)
  Dim TempString 
  On Error Resume Next 
  TempString = RSBinaryToString(Binary) 
  If Len(TempString) <> LenB(Binary) then 
  TempString = MBBinaryToString(Binary)
  end if
  BinaryToString = TempString
End Function
Function MBBinaryToString(Binary)
  dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    if cl3>300 then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      if cl2>200 then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  MBBinaryToString = pl1 & pl2 & pl3
End Function
Function MultiByteToBinary(MultiByte)
  Dim RS, LMultiByte, Binary
  Const adLongVarBinary = 205
  Set RS = CreateObject("ADODB.Recordset")
  LMultiByte = LenB(MultiByte)
 if LMultiByte>0 then
  RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
  RS.Open
  RS.AddNew
   RS("mBinary").AppendChunk MultiByte & ChrB(0)
  RS.Update
  Binary = RS("mBinary").GetChunk(LMultiByte)
 End If
  MultiByteToBinary = Binary
End Function
Function GetHeadFields(ByVal Head, Name, FileName, Content_Disposition, Content_Type)
  Name = (SeparateField(Head, "name=", ";")) 'ltrim
  If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)
  FileName = (SeparateField(Head, "filename=", ";")) 'ltrim
  If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)
  Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
  Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function
Function SeparateField(From, ByVal sStart, ByVal sEnd)
  Dim PosB, PosE, sFrom
  sFrom = LCase(From)
  PosB = InStr(sFrom, sStart)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    PosE = InStr(PosB, sFrom, sEnd)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(From, PosB, PosE - PosB)
  Else
    SeparateField = Empty
  End If
End Function
Function SplitFileName(FullPath)
  Dim Pos, PosF
  PosF = 0
  For Pos = Len(FullPath) To 1 Step -1
    Select Case Mid(FullPath, Pos, 1)
      Case ":", "/", "\": PosF = Pos + 1: Pos = 0
    End Select
  Next
  If PosF = 0 Then PosF = 1
 SplitFileName = PosF
End Function
Function GetPath(FullPath)
  GetPath = left(FullPath, SplitFileName(FullPath)-1)
End Function
Function GetFileName(FullPath)
  GetFileName = Mid(FullPath, SplitFileName(FullPath))
End Function
Function RecurseMKDir(ByVal Path)
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  Path = Replace(Path, "/", "\")
  If Right(Path, 1) <> "\" Then Path = Path & "\"   '"
  Dim Pos, n
  Pos = 0: n = 0
  Pos = InStr(Pos + 1, Path, "\")   '"
  Do While Pos > 0
    On Error Resume Next
    FS.CreateFolder Left(Path, Pos - 1)
    If Err = 0 Then n = n + 1
    Pos = InStr(Pos + 1, Path, "\")   '"
  Loop
  RecurseMKDir = n
End Function
Function SaveBinaryData(FileName, ByteArray)
 SaveBinaryData = SaveBinaryDataStream(FileName, ByteArray)
End Function
Function SaveBinaryDataTextStream(FileName, ByteArray)
  Dim FS : Set FS = CreateObject("Scripting.FileSystemObject")
 On error Resume next
  Dim TextStream 
 Set TextStream = FS.CreateTextFile(FileName)
 if Err = &H4c then 'Path not found.
  On error Goto 0
  RecurseMKDir GetPath(FileName)
  On error Resume next
  Set TextStream = FS.CreateTextFile(FileName)
 end if
  TextStream.Write BinaryToString(ByteArray) 'BinaryToString is in upload.inc.
  TextStream.Close
 Dim ErrMessage, ErrNumber
 ErrMessage = Err.Description
 ErrNumber = Err
 On Error Goto 0
 if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 
End Function
Function SaveBinaryDataStream(FileName, ByteArray)
 Dim BinaryStream
 Set BinaryStream = createobject("ADODB.Stream")
 BinaryStream.Type = 1 'Binary
 BinaryStream.Open
 BinaryStream.Write ByteArray
 On error Resume next
 BinaryStream.SaveToFile FileName, 2 'Overwrite
 if Err = &Hbbc then 'Path not found.
  On error Goto 0
  RecurseMKDir GetPath(FileName)
  On error Resume next
  BinaryStream.SaveToFile FileName, 2 'Overwrite
 end if
 Dim ErrMessage, ErrNumber
 ErrMessage = Err.Description
 ErrNumber = Err
 On Error Goto 0
 if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 
End Function
Class cResponse
 Public Property Get IsClientConnected
  randomize
  IsClientConnected = cbool(clng(rnd * 4))
  IsClientConnected = True
 End Property 
End Class 
%>
