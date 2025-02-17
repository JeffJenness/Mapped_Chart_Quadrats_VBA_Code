Attribute VB_Name = "aml_func_mod"
Option Explicit

Public Enum enumArcGISFolderTypes
   enumLastBrowsedLocation
   enumLastExportToLocation
   enumLastLocation
   enumLastSaveToLocation
   enumArcGISInstallLocation
End Enum

Public Function ReturnArcGISVersionAlt(pMxDoc As IMxDocument) As Long

   Dim pDocVersion As IDocumentVersion
   Set pDocVersion = pMxDoc

   ReturnArcGISVersionAlt = CLng(pDocVersion.DocumentVersion)

End Function

Public Function ReturnArcGISVersionAlt2(pMxDoc As IMxDocument, Optional strDecimalVersion As String) As Long

   Dim pDocVersion As IDocumentVersion
   Set pDocVersion = pMxDoc

   ReturnArcGISVersionAlt2 = CLng(pDocVersion.DocumentVersion)

  Select Case ReturnArcGISVersionAlt2
    Case 0
      strDecimalVersion = "8.3"
    Case 1
      strDecimalVersion = "9.0"
    Case 2
      strDecimalVersion = "9.2"
    Case 3
      strDecimalVersion = "9.3"
    Case 4
      strDecimalVersion = "10.0"
    Case 5
      strDecimalVersion = "10.1"
    Case 6
      strDecimalVersion = "10.2"
    Case Else
      strDecimalVersion = "Unable to determine ArcGIS Version!"
  End Select

End Function

Public Function QueryForNewFilename(strSampleName As String) As String

End Function

Public Function ReturnShapeName(pEnum As esriGeometryType) As String

  Select Case pEnum
  Case 0
    ReturnShapeName = "Unknown Geometry"
  Case 1
    ReturnShapeName = "Point"
  Case 2
    ReturnShapeName = "Multipoint"
  Case 3
    ReturnShapeName = "Polyline"
  Case 4
    ReturnShapeName = "Polygon"
  Case 5
    ReturnShapeName = "Envelope"
  Case 6
    ReturnShapeName = "Path"
  Case 7
    ReturnShapeName = "Unknown Geometry"
  Case 9
    ReturnShapeName = "Multipatch"
  Case 11
    ReturnShapeName = "Ring"
  Case 13
    ReturnShapeName = "Line"
  Case 14
    ReturnShapeName = "Circular Arc"
  Case 15
    ReturnShapeName = "Bezier Curve"
  Case 16
    ReturnShapeName = "Elliptic Arc"
  Case 17
    ReturnShapeName = "Geometry Bag"
  Case 18
    ReturnShapeName = "Triangle Strip"
  Case 19
    ReturnShapeName = "Triangle Fan"
  Case 20
    ReturnShapeName = "Ray"
  Case 21
    ReturnShapeName = "Sphere"
  Case 22
    ReturnShapeName = "Triangles"
  Case Else
    ReturnShapeName = "Unknown Geometry"
  End Select

End Function

Public Function FileExists(strFilename As String) As Boolean

  FileExists = Dir(strFilename) <> ""

End Function

Public Function ReturnArcGISVersion() As String

   ReturnArcGISVersion = "Unable to determine ArcGIS Version!"

   Dim nRet As Long
   Dim hKey As Long
   Dim nType As Long
   Dim nBytes As Long
   Dim Buffer As String

   Dim strSection As String
   strSection = ""
   Dim strKey As String
   strKey = "RealVersion"
   Dim strDefault As String
   strDefault = ""
   Dim strDir As String

   nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey("ESRI", "ArcGIS", strSection), 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
      If strKey = "*" Then strKey = vbNullString

      nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               strDir = Left(Buffer, nBytes - 1)
               ReturnArcGISVersion = strDir
            End If
         End If
      End If
      Call RegCloseKey(hKey)
   End If

End Function

Public Function ReturnArcGISInstallDir() As String

   ReturnArcGISInstallDir = "Unable to determine ArcGIS Install location!"

   Dim nRet As Long
   Dim hKey As Long
   Dim nType As Long
   Dim nBytes As Long
   Dim Buffer As String

   Dim strSection As String
   strSection = ""
   Dim strKey As String
   strKey = "InstallDir"
   Dim strDefault As String
   strDefault = ""
   Dim strDir As String

   nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey("ESRI", "ArcGIS", strSection), 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
      If strKey = "*" Then strKey = vbNullString

      nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               strDir = Left(Buffer, nBytes - 1)
               If Right(strDir, 1) = "\" Then strDir = Left(strDir, Len(strDir) - 1)
               ReturnArcGISInstallDir = strDir
            End If
         End If
      End If
      Call RegCloseKey(hKey)
   End If

End Function

Public Function ReturnArcGISGeneralDir(pArcGISFolderType As enumArcGISFolderTypes) As String

  Dim nRet As Long
  Dim hKey As Long
  Dim nType As Long
  Dim nBytes As Long
  Dim Buffer As String
  Dim strKey As String
  Dim strDir As String

  Select Case pArcGISFolderType
    Case enumArcGISInstallLocation
      ReturnArcGISGeneralDir = "Unable to determine ArcGIS Install location!"
      nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\ESRI\ArcGIS", 0&, KEY_READ, hKey)
      strKey = "InstallDir"
    Case enumLastSaveToLocation
      ReturnArcGISGeneralDir = "Unable to determine ArcGIS Last Saved To location!"
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\ESRI\ArcCatalog\Settings", 0&, KEY_READ, hKey)
      strKey = "LastSaveToLocation"
    Case enumLastBrowsedLocation
      ReturnArcGISGeneralDir = "Unable to determine ArcGIS Last Browsed To location!"
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\ESRI\ArcCatalog\Settings", 0&, KEY_READ, hKey)
      strKey = "LastBrowseLocation"
    Case enumLastExportToLocation
      ReturnArcGISGeneralDir = "Unable to determine ArcGIS Last Exported To location!"
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\ESRI\ArcCatalog\Settings", 0&, KEY_READ, hKey)
      strKey = "LastExportToLocation"
    Case enumLastLocation
      ReturnArcGISGeneralDir = "Unable to determine ArcGIS Last location!"
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\ESRI\ArcCatalog\Settings", 0&, KEY_READ, hKey)
      strKey = "LastLocation"
  End Select

   If nRet = ERROR_SUCCESS Then
      If strKey = "*" Then strKey = vbNullString

      nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               strDir = Left(Buffer, nBytes - 1)
               If Right(strDir, 1) = "\" Then strDir = Left(strDir, Len(strDir) - 1)
               ReturnArcGISGeneralDir = strDir
            End If
         End If
      End If
      Call RegCloseKey(hKey)
   End If

End Function

Private Function SubKey(ByVal strCompany As String, ByVal strAppName As String, Optional ByVal Section As String = "") As String

   SubKey = "Software\" & strCompany & "\" & strAppName
   If Len(Section) Then
      SubKey = SubKey & "\" & Section
   End If

End Function

Public Function ReturnFolders(ByVal DirPath As String, Optional ByVal FolderName As String) As Collection

  Dim returnCollection As Collection
  Set returnCollection = New Collection

  Dim strFolders() As String

  Dim strSearchString As String
  If FolderName = "" Then
    strSearchString = "*"
  Else
    strSearchString = FolderName
  End If

  Dim strSearchPath As String

  DirPath = Trim(DirPath)
  If Not Right(DirPath, 1) = "\" Then DirPath = DirPath & "\"
  strSearchPath = DirPath & strSearchString

  Dim strFoundFolder As String
  Dim anIndex As Integer
  anIndex = -1

  strFoundFolder = Dir(strSearchPath, 16)

  Do While Not strFoundFolder = ""
    If (Not strFoundFolder = ".") And (Not strFoundFolder = "..") Then
      If GetAttr(DirPath & strFoundFolder) = vbDirectory Then
        anIndex = anIndex + 1
        ReDim Preserve strFolders(anIndex)
        strFolders(anIndex) = DirPath & strFoundFolder
      End If
    End If
    strFoundFolder = Dir
  Loop

  returnCollection.Add (UBound(strFolders) - LBound(strFolders) + 1)
  returnCollection.Add (strFolders)

  Set ReturnFolders = returnCollection

End Function

Public Function ReturnFiles(ByVal DirPath As String, Optional ByVal FileName As String, Optional ByVal IncludeReadOnlyFiles As Boolean, _
                            Optional ByVal IncludeHiddenFiles As Boolean, Optional ByVal IncludeSystemFiles As Boolean) As Collection

  Dim returnCollection As Collection
  Set returnCollection = New Collection

  Dim strFiles() As String
  Dim intOption As Integer
  intOption = 0
  If IncludeReadOnlyFiles Then intOption = intOption + 1
  If IncludeHiddenFiles Then intOption = intOption + 2
  If IncludeSystemFiles Then intOption = intOption + 4

  Dim strSearchString As String
  If FileName = "" Then
    strSearchString = "*"
  Else
    strSearchString = FileName
  End If

  Dim strSearchPath As String

  DirPath = Trim(DirPath)
  If Not Right(DirPath, 1) = "\" Then DirPath = DirPath & "\"
  strSearchPath = DirPath & strSearchString

  Dim strFoundFile As String
  Dim anIndex As Integer

  strFoundFile = Dir(strSearchPath, intOption)

  Do While Not strFoundFile = ""
    ReDim Preserve strFiles(anIndex)
    strFiles(anIndex) = DirPath & strFoundFile
    strFoundFile = Dir
    anIndex = anIndex + 1
  Loop

  returnCollection.Add anIndex
  returnCollection.Add (strFiles)

  Set ReturnFiles = returnCollection

End Function

Public Function QuoteString(strInput As String) As String

  Dim strQuoted As String
  strQuoted = m_Quotation & SubstituteString(strInput, Chr(34), m_Quotation & m_Quotation) & m_Quotation

  QuoteString = strQuoted

End Function

Public Function ContainsString(strInText As String, strSearchText As String) As Boolean

  If strInText = "" Or strSearchText = "" Then
    ContainsString = False
  Else
    ContainsString = InStr(1, strInText, strSearchText, vbTextCompare) > 0
  End If

End Function

Public Function SubstituteString(strFullText As String, strSearchText As String, strSubstituteText As String)

  Dim lngIndex As Long
  Dim lngStartPos As Long
  Dim lngSearchLength As Long
  Dim lngFullLength As Long

  lngStartPos = 1
  lngSearchLength = Len(strSearchText)
  lngFullLength = Len(strFullText)

  lngIndex = InStr(lngStartPos, strFullText, strSearchText, vbTextCompare)

  Dim strNewString As String
  If lngIndex = 0 Then
    strNewString = strFullText
  Else
    strNewString = Left(strFullText, lngIndex - 1) & strSubstituteText
    lngStartPos = lngIndex + lngSearchLength
  End If

  Do While lngIndex <> 0
    lngIndex = InStr(lngStartPos, strFullText, strSearchText, vbTextCompare)
    If lngIndex = 0 Then
      strNewString = strNewString & Right(strFullText, lngFullLength - lngStartPos + 1)
    Else
      strNewString = strNewString & Mid(strFullText, lngStartPos, lngIndex - lngStartPos) & strSubstituteText
      lngStartPos = lngIndex + lngSearchLength
    End If

  Loop
  SubstituteString = strNewString

End Function

Public Function GetMxDocPath(pApp As IApplication) As String

  Dim pTemplates As ITemplates
  Dim lTempCount As Long

  Set pTemplates = pApp.Templates
  lTempCount = pTemplates.Count

  GetMxDocPath = pTemplates.Item(lTempCount - 1)

End Function

Public Function CreatedBASETable(strFullName As String, Optional pFields As IFields) As ITable

  Dim strName As String
  Dim strFolder As String

  strFolder = aml_func_mod.ReturnDir(strFullName)
  strName = aml_func_mod.ReturnFilename(strFullName)
  If Right(strName, 4) = ".dbf" Then strName = Left(strName, Len(strName) - 4)

  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Dim fs As Object
  Dim pFieldsEdit As IFieldsEdit
  Dim pFieldEdit As IFieldEdit
  Dim pField As iField

  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  If Not aml_func_mod.ExistFileDir(strFolder) Then
    MsgBox "Folder does not exist: " & vbCr & strFolder
    Exit Function
  End If

  Set pFWS = pWorkspaceFactory.OpenFromFile(strFolder, 0)

  If pFields Is Nothing Then
    Set pFields = New Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 1

    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
        .Precision = 8
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
    End With
    Set pFieldsEdit.Field(0) = pField
  End If

  Dim strString As String
  Dim lngIndex As Long
  Dim pFieldInfo As IFieldInfo

  Set pField = pFields.Field(0)

  Set CreatedBASETable = pFWS.CreateTable(strName, pFields, Nothing, Nothing, "")

End Function

Public Function GetTheUserName() As String

  Dim sBuffer As String
  Dim sUName As String
  Dim lSize As Long
  sBuffer = Space$(255)
  lSize = Len(sBuffer)
  Call GetUserName(sBuffer, lSize)
  If lSize > 0 Then
    sUName = Left$(sBuffer, lSize)
  Else
    sUName = vbNullString
  End If
  GetTheUserName = BasicTrimAvenue(sUName, "", Chr(0))     ' NEED TO PEEL OFF THAT LAST ODD CHARACTER

End Function

Public Function BasicTrimAvenue(aString As String, aTrimLeft As String, aTrimRight As String) As String

  Do While (aString <> "") And (InStr(1, aTrimRight, Right(aString, 1), vbTextCompare) > 0)
    aString = Left(aString, Len(aString) - 1)
  Loop
  Do While (aString <> "") And (InStr(1, aTrimLeft, Left(aString, 1), vbTextCompare) > 0)
    aString = Right(aString, Len(aString) - 1)
  Loop

  BasicTrimAvenue = aString

End Function

Public Function ClipExtension2(strPathName As String) As String

  Dim lngLastDot As Long
  Dim lngLastSlash As Long
  Dim lngLastForwardSlash As Long

  lngLastDot = InStrRev(strPathName, ".")
  lngLastSlash = InStrRev(strPathName, "\")
  lngLastForwardSlash = InStrRev(strPathName, "/")

  Dim strSplit() As String
  Dim strFinalPath As String
  Dim lngIndex As Long

  If lngLastDot > 0 And lngLastDot > lngLastSlash And lngLastDot > lngLastForwardSlash Then
    strSplit = Split(strPathName, ".")
    strFinalPath = strSplit(0)
    For lngIndex = 1 To UBound(strSplit) - 1
      strFinalPath = strFinalPath & "." & strSplit(lngIndex)
    Next lngIndex
    ClipExtension2 = strFinalPath
  Else
    ClipExtension2 = strPathName
  End If

  Erase strSplit

End Function

Public Function ClipExtension(strPathName As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String

  aml_func_mod.ParseString strPathName, strDirTokens, "."
  strDirPath = strDirTokens(0)

  If (UBound(strDirTokens) = 0) Then
    ClipExtension = strDirPath
  Else
    Dim anIndex As Long
    For anIndex = 1 To (UBound(strDirTokens) - 1)
      strDirPath = strDirPath & "." & strDirTokens(anIndex)
    Next anIndex
    ClipExtension = strDirPath
  End If

  Erase strDirTokens

End Function

Public Function SetExtension(strPathName As String, strExtension As String) As String

  Dim theClippedPath As String
  SetExtension = ClipExtension(strPathName) & "." & strExtension

End Function

Public Function GetExtensionText(strPathName As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String

  aml_func_mod.ParseString strPathName, strDirTokens, "."
  If UBound(strDirTokens) = 0 Then
    GetExtensionText = ""
  Else
    GetExtensionText = strDirTokens(UBound(strDirTokens))
  End If

End Function

Public Function ReturnDir(strPathName As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String

  If InStr(1, strPathName, "\") = 0 Then
    ReturnDir = ""
  Else

    aml_func_mod.ParseString strPathName, strDirTokens, "\"
    strDirPath = strDirTokens(0)

    If (UBound(strDirTokens) = 0) Then
      ReturnDir = strDirPath
    Else
      Dim anIndex As Long
      For anIndex = 1 To (UBound(strDirTokens) - 1)
        strDirPath = strDirPath & "\" & strDirTokens(anIndex)
      Next anIndex
      ReturnDir = strDirPath
    End If

    ReturnDir = ReturnDir & "\"
  End If

End Function

Public Function ReturnFilename(strPathName As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String

  If InStr(1, strPathName, "\") = 0 Then
    ReturnFilename = strPathName
  Else

    aml_func_mod.ParseString strPathName, strDirTokens, "\"
    ReturnFilename = strDirTokens(UBound(strDirTokens))
  End If

End Function

Public Sub ParseString(str As String, strArray() As String, Delim As String)

Dim i As Long
Dim tokenlen As Long
Dim tmpstr As String
Dim position As Long
Dim length As Long

Dim switch As Long
Dim position1 As Long
Dim position2 As Long
Dim pair As Long

On Error Resume Next

  If Trim(Subst(str, Delim)) = "" Then
    err.Raise vbObjectError + 1, "aml_func.ParseString", _
    "StringPassed"
    Exit Sub
  End If

  ReDim strArray(0)

  pair = False
  switch = 0
  length = Len(str)
  position = 1
  i = 0
  tmpstr = str

  Do While position < length
    If Mid(tmpstr, position, 1) = "'" Then
      If Not (switch = 1) Then
        switch = 1
        position1 = position
        pair = False
      Else
        switch = 2
        position2 = position
        pair = True
      End If
    End If

    If pair = True Then
      Mid(tmpstr, position1, 1) = " "
      Mid(tmpstr, position2, 1) = " "
      strArray(i) = Mid(tmpstr, position1, position2 - position1)
      strArray(i) = Trim(strArray(i))
      pair = False
      switch = 0

    Else
      If switch = 0 Then
        If Mid(tmpstr, position, 1) = Delim Then
          strArray(i) = Left(tmpstr, position)
          tokenlen = Len(strArray(i))
          strArray(i) = Trim(strArray(i))
          If Not (Len(strArray(i)) = 0) Then
            Mid(tmpstr, 1, tokenlen) = String(tokenlen, " ")
            ReDim Preserve strArray(LBound(strArray) To UBound(strArray) + 1)
            i = i + 1
          End If
        End If
      End If
    End If
    position = position + 1
  Loop
  strArray(i) = Trim(tmpstr)

  position = 1
  For i = 0 To UBound(strArray)
    position = Len(strArray(i))
    If Mid(strArray(i), position, 1) = Delim Then
      Mid(strArray(i), position, 1) = " "
      strArray(i) = Trim(strArray(i))
    End If
  Next i

End Sub

Public Function After(str As String, SearchStr As String) As String

Dim position As Long
Dim length As Long

  position = InStr(str, SearchStr)
  length = Len(SearchStr)
  If Not (position = 0) Then
   After = Mid(str, position + length)
  End If

End Function

Public Function Before(str As String, SearchStr As String) As String

Dim position As Long
Dim length As Long

  position = InStr(str, SearchStr)
  length = Len(SearchStr)
  If Not (position = 0) Then
   Before = Mid(str, 1, position - 1)
  End If

End Function

Function ExistFileDir(sTest As String) As Boolean

  Dim af As Long
  af = -1
  On Error Resume Next
  af = GetAttr(sTest)
  ExistFileDir = (af <> -1)

End Function

Public Function MakeUniqueFilename(strFilename As String) As String

  Dim theFilename As String
  Dim theBaseName As String
  Dim thePointPos As Long
  Dim theExtension As String

  If InStr(1, Right(strFilename, 5), ".", vbTextCompare) > 0 Then
    thePointPos = InStrRev(strFilename, ".", -1, vbTextCompare)
    theExtension = Right(strFilename, Len(strFilename) - (thePointPos - 1))
    theFilename = Left(strFilename, thePointPos - 1)
  Else
    theExtension = ""
    theFilename = strFilename
  End If

  If Not ExistFileDir(strFilename) And Not ExistFileDir(theFilename) Then
    MakeUniqueFilename = strFilename
    Exit Function
  Else

    Dim theCounter As Long
    theCounter = 1

    theBaseName = theFilename

    Do While ExistFileDir(theFilename & theExtension) Or ExistFileDir(theFilename)
      theCounter = theCounter + 1
      theFilename = theBaseName & "_" & CStr(theCounter)
    Loop

    MakeUniqueFilename = theFilename & theExtension

  End If

End Function

Public Sub ParseStringR(str As String, strArray() As String, Optional ReturnQuoted)

Dim i As Long
Dim tokenlen As Long
Dim switch As Long
Dim position1 As Long
Dim position2 As Long
Dim pair As Long
Dim parseAgain As Boolean
Dim tmpstr As String
Dim position As Long
Dim length As Long

On Error Resume Next

  If Trim(Subst(str, ",")) = "" Then
    err.Raise vbObjectError + 1, "aml_func.ParseStringR", _
    "StringPassed"
    Exit Sub
  End If

  ReDim strArray(0)
  pair = False
  switch = 0
  length = Len(str)
  position = 1
  i = 0
  tmpstr = str

  If IsMissing(ReturnQuoted) Then
    ReturnQuoted = False
  End If
  If Not (ReturnQuoted = False) Then
    ReturnQuoted = True
  End If

  Do While position < length
    If Mid(tmpstr, position, 1) = "'" Then
      If Not (switch = 1) Then
      switch = 1
      position1 = position
      pair = False
    Else
      switch = 2
      position2 = position
      pair = True
      End If
    End If

    If pair = True Then
      Mid(tmpstr, position1, 1) = " "
      Mid(tmpstr, position2, 1) = " "
      strArray(i) = Mid(tmpstr, position1, position2 - position1)
      strArray(i) = Trim(strArray(i))
      pair = False
      switch = 0

    Else
      If switch = 0 Then
        If Mid(tmpstr, position, 1) = "," Or Mid(tmpstr, position, 1) = " " Then
          strArray(i) = Left(tmpstr, position)
          tokenlen = Len(strArray(i))
          strArray(i) = Trim(strArray(i))
          If Not (Len(strArray(i)) = 0) Then
            Mid(tmpstr, 1, tokenlen) = String(tokenlen, " ")
            ReDim Preserve strArray(LBound(strArray) To UBound(strArray) + 1)
            i = i + 1
          End If
        End If
      End If
    End If
    position = position + 1
  Loop
    strArray(i) = Trim(tmpstr)

  parseAgain = False
  position = 1

  For i = 0 To UBound(strArray)
    position = Len(strArray(i))
    If Mid(strArray(i), position, 1) = "," Then
      Mid(strArray(i), position, 1) = " "
      strArray(i) = Trim(strArray(i))
      If Len(strArray(i)) = 0 Then
        parseAgain = True
      End If
    End If
  Next i

  If parseAgain = True Then
    tmpstr = ""
    For i = 0 To UBound(strArray)
      If Not (Len(strArray(i))) = 0 Then
        tmpstr = tmpstr & "'" & strArray(i) & "'" & ","
      End If
    Next i
    ParseString (tmpstr), strArray, ","
  End If

  If ReturnQuoted = True Then
    For i = 0 To UBound(strArray)
     strArray(i) = "'" & strArray(i) & "'"
    Next i
  End If

End Sub

Public Function Subst(str As String, SearchChar As String, Optional ReplaceChar) As String

Dim complete As Boolean
Dim i As Long
Dim first As String
Dim last As String
Dim tmpstr As String
Dim position As Long

tmpstr = str

position = 1
complete = False

If IsMissing(ReplaceChar) Then
  Do Until complete = True
    position = InStr(position, tmpstr, SearchChar)
    If position = 0 Or position > Len(tmpstr) Then
      complete = True
    Else
      first = Before(tmpstr, SearchChar)
      last = After(tmpstr, SearchChar)
      tmpstr = first & last
    End If
  Loop
End If

Do Until complete = True
  position = InStr(position, tmpstr, SearchChar)
  If position = 0 Or position > Len(tmpstr) Then
    complete = True
  Else
    Mid(tmpstr, position, Len(ReplaceChar)) = ReplaceChar
    position = position + Len(ReplaceChar)
  End If
Loop

Subst = tmpstr

End Function

Public Function ReturnFilename2(strPathName As String) As String

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  lngIndex1 = InStrRev(strPathName, "\", , vbTextCompare)
  lngIndex2 = InStrRev(strPathName, "/", , vbTextCompare)

  If lngIndex1 = 0 And lngIndex2 = 0 Then
    ReturnFilename2 = strPathName
  Else
    If lngIndex1 = 0 Then
      ReturnFilename2 = Right(strPathName, Len(strPathName) - lngIndex2)
    Else
      ReturnFilename2 = Right(strPathName, Len(strPathName) - lngIndex1)
    End If
  End If

End Function

Public Function ReturnDir3(strPathName As String, Optional booPutTrailingBackslash As Boolean = True) As String

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  lngIndex1 = InStrRev(strPathName, "\", , vbTextCompare)
  lngIndex2 = InStrRev(strPathName, "/", , vbTextCompare)

  If lngIndex1 = 0 And lngIndex2 = 0 Then
    ReturnDir3 = strPathName
  Else
    If lngIndex1 = 0 Then
      ReturnDir3 = Left(strPathName, lngIndex2)
    Else
      ReturnDir3 = Left(strPathName, lngIndex1)
    End If
  End If

  If booPutTrailingBackslash Then
    If Right(ReturnDir3, 1) <> "\" Then
      ReturnDir3 = ReturnDir3 & "\"
    End If
  Else
    If Right(ReturnDir3, 1) = "\" Then
      ReturnDir3 = Left(ReturnDir3, Len(ReturnDir3) - 1)
    End If
  End If

End Function


