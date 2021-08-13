Attribute VB_Name = "aml_func_mod"
'             Environmental Systems Research Institute, Inc.
'Module Name: aml_func.bas
'Description: Used to perform AML-like functions for string manipipulation.
'             Also parses out tokens in a string to elements in an array.
'
'   Requires: ParseString and ParseStringR require that you Dim the named array arg in
'             the calling propgram: Must be a zero dimensioned string array:
'             Dim yourarray() as string

'    Methods: After   - Returns the substring to the right of the leftmost occurrence of
'                       searchStr
'             Before  - Returns the substring to the left of the leftmost occurrence of
'                       searchStr.
'        ExistFileDir - Returns True/False if file or Directory exists.
'             Extract - Returns an element from a list of elements
'             Index   - Returns the position of the leftmost occurrence of a specified
'                       string in a target string.
'             Keyword - Returns the position of a keyword within a list of keywords.
'             Search  - Returns the position of the first character of a search string
'                       in a target string.
'             Sort    - Returns a sorted a list of elements.
'             Subst   - Returns a string that has had one string subsituted for another
'             Substr  - Returns a substring that starts at a specified character position.
'             Token   - Allows tokens in a list to be manipulated.
'
'                       Count - the number of tokens in a list
'                       Find  - the position of a token in a list
'                       Move  - tokens in a list
'                       Insert- a new token into a list
'                       Delete- a token in a list
'                       Replace-one token in a list for another
'                       Switch -one token in a list for another
'
'        ParseString  - Populates string array with tokens in a string.
'        ParseStringR - Same as ParseString except blanks and commas are
'                       treated as delimiters.
'
'    History: DMA     - 03/04/97 - Original coding
'       Glenn Meister - 12/19/97 - Added ExistFileDir function
'
'   MODIFICATIONS:  JEFF JENNESS
'   GetFullFileString - ' CONVERTS 8.3 FILESTRING TO FULL TEXT
'           ReturnDir - GIVEN A STRING, RETURNS TEXT PRECEDING LAST "\" CHARACTER
'          ReturnDir2 - GIVEN A STRING, RETURNS TEXT PRECEDING LAST "\" CHARACTER, MORE EFFICENT THAN ReturnDir
'          ReturnDir3 - GIVEN A STRING, RETURNS TEXT PRECEDING LAST "\" CHARACTER WITH OPTION TO FORCE A TRAILING BACKSLASH
'      ReturnFilename - GIVEN A STRING, RETURNS TEXT FOLLOWING LAST "\" CHARACTER
'         ReturnFiles - GIVEN A STRING PATH, RETURNS COLLECTION CONTAINING THE NUMBER OF FILES THAT MET CRITERIA
'                       AND ARRAY OF STRING FILENAMES.  OPTIONS FOR INCLUDE READ-ONLY, SYSTEM, HIDDEN FILES.
'       ReturnFolders - GIVENA STRING PATH, RETURNS COLLECTION CONTAINING THE NUMBER OF FOLDERS THAT MET CRITERIA
'                       AND ARRAY OF FOLDER FILENAMES.
'       ClipExtension - GIVEN A STRING, RETURNS TEXT PRECEDING LAST "." CHARACTER
'        SetExtension - GIVEN A STRING AND EXTENSION, RETURNS TEXT PRECEDING LAST "." CHARACTER, PLUS "." AND NEW EXTENSION
'    GetExtensionText - GIVEN A STRING, RETURNS TEXT FOLLOWING LAST "." CHARACTER
'      FieldIsNumeric - GIVEN A FIELD, RETURNS BOOLEAN
'       FieldIsString - GIVEN A FIELD, RETURNS BOOLEAN
'         FieldIsDate - GIVEN A FIELD, RETURNS BOOLEAN
'        FieldIsShape - GIVEN A FIELD, RETURNS BOOLEAN
'        InsertCommas - GIVEN A NUMBER OR STRING, INSERTS COMMAS AND RETURNS STRING
'     BasicTrimAvenue - GIVEN A STRING, RETURNS A STRING WITH VALUES TRIMMED OFF EACH SIDE
'  MakeUniqueFilename - GIVEN A STRING REPRESENTING A FILEPATH, RETURNS A UNIQUE FILENAME WHICH MAY HAVE NUMBERS APPENDED TO IT
'      GetTheUserName - JUST RETURNS THE USER NAME; CAN BE USED TO FIND THE "MY DOCUMENTS" FOLDER
'    TempPathLocation - RETURNS PATHNAME TO TEMP DIRECTORY
'     CreateShapefile - GIVEN PATHNAME AND SHAPE TYPE, RETURNS A pFeatureClass.  ESRI CODE, MODIFIED BY JENNESS
'    CreateShapefile2 - GIVEN PATHNAME, SHAPE TYPE AND ARRAY OF FIELDS, RETURNS A pFeatureClass.
'    CreatedBASETable - GIVEN PATHNAME AND OPTIONAL FIELDS, RETURNS ITABLE
'        GetMxDocPath - GIVEN APPLICATION OBJECT, RETURNS PATHNAME OF MAP DOCUMENT
'         QuoteString - GIVEN A STRING, RETURNS A QUOTED VERSION OF THAT STRING
'    SubstituteString - GIVEN A STRING, A SEARCH STRING AND A SUBSTITUTE STRING, REPLACES ALL INSTANCES OF SEARCH TEXT WITH
'                       SUBSTITUTE TEXT IN THE ORIGINAL STRING
'ReturnArcGISInstallDir - RETURNS A STRING CONTAINING THE ARCGIS INSTALL LOCATION
'ReturnArcGISGeneralDir - RETURNS DIRECTORY STRING FOR INSTALL, LAST LOCATION, LAST BROWSED, LAST EXPORT AND LAST SAVE TO
' ReturnArcGISVersion - RETURNS ARCGIS VERSION NUMBER
'ReturnArcGISVersionAlt - MORE DIRECT METHOD OF GETTING ARCGIS VERSION; RETURNS A LONG WHERE 0 = 8.3, 1 = 9.0, 2 = 9.2, 3 = 9.3
'          FileExists - GIVEN A STRING PATH, RETURNS TRUE OR FALSE
' ReturnShapeTypeName - GIVEN A SHAPE TYPE ENUMERATION, RETURNS SHAPE NAME
' QueryForNewFilename - GIVEN A SAMPLE FILENAME, OPENS DIALOG AND QUERIES USER TO SPECIFY NEW FILENAME

Option Explicit

' FILENAME 8.3 FORMAT CONVERSION BELOW COPIED FROM http://forums.esri.com/Thread.asp?c=93&f=993&t=123512&mc=1#msgid400360
'   POSTED BY  Brett N.Meroney
'              GIS Programmer / Analyst
'              Integrated Laboratory Systems, Inc.
'              contractor to US-EPA Region VIII
'              Golden , CO
' NECESSARY TO CONVERT 8.3 FILENAMES TO USEABLE FILENAMES

Private Const strMAXPATH = 260
Private Const m_Quotation = """"

'
' Win32 Registry functions
'
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As Long) _
     As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal Reserved As Long, _
     ByVal lpClass As String, _
     ByVal dwOptions As Long, _
     ByVal samDesired As Long, _
     lpSecurityAttributes As Any, _
     phkResult As Long, _
     lpdwDisposition As Long) _
     As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
     (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Any, _
     lpcbData As Long) _
     As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
     (ByVal hKey As Long, _
     ByVal lpClass As String, _
     lpcbClass As Long, _
     lpReserved As Long, _
     lpcSubKeys As Long, _
     lpcbMaxSubKeyLen As Long, _
     lpcbMaxClassLen As Long, _
     lpcValues As Long, _
     lpcbMaxValueNameLen As Long, _
     lpcbMaxValueLen As Long, _
     lpcbSecurityDescriptor As Long, _
     lpftLastWriteTime As Any) _
     As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
     (ByVal hKey As Long, _
     ByVal dwIndex As Long, _
     ByVal lpName As String, _
     lpcbName As Long, _
     lpReserved As Long, _
     ByVal lpClass As String, _
     lpcbClass As Long, _
     lpftLastWriteTime As Any) _
     As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
     (ByVal hKey As Long, _
      ByVal dwIndex As Long, _
     ByVal lpValueName As String, _
     lpcbValueName As Long, _
     lpReserved As Long, _
     lpType As Long, _
     lpData As Any, _
     lpcbData As Long) _
     As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
     (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpData As Any, _
     ByVal cbData As Long) _
     As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
     (ByVal hKey As Long, _
      ByVal lpSubKey As String) _
     As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
     (ByVal hKey As Long, _
     ByVal lpValueName As String) _
     As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
     (ByVal hKey As Long) _
     As Long
'
' Constants for Windows 32-bit Registry API
'
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
'
' Reg result codes
'
Private Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Private Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
'
' Reg Create Type Values...
'
Private Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
'
' Reg Key Security Options
'
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const REG_SZ = 1                         ' Unicode nul terminated string
'

Private Declare Function GetLongPathName Lib "Kernel32" Alias _
    "GetLongPathNameA" (ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, ByVal cchBuffer As Long) _
    As Long

Public Enum enumArcGISFolderTypes
   enumLastBrowsedLocation
   enumLastExportToLocation
   enumLastLocation
   enumLastSaveToLocation
   enumArcGISInstallLocation
End Enum

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Const Quotation = """"

Public Function ReturnArcGISVersionAlt(pMxDoc As IMxDocument) As Long

 '  ReturnArcGISVersionAlt = "Unable to determine ArcGIS Version!"
   
   Dim pDocVersion As IDocumentVersion
   Set pDocVersion = pMxDoc
   
   ReturnArcGISVersionAlt = CLng(pDocVersion.DocumentVersion)

End Function
Public Function ReturnArcGISVersionAlt2(pMxDoc As IMxDocument, Optional strDecimalVersion As String) As Long

 '  ReturnArcGISVersionAlt = "Unable to determine ArcGIS Version!"
   ' 0 = 8.3, 1 = 9.0, 2 = 9.2, 3 = 9.3
   
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

' ORIGINAL FUNCTION:
' Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
   ' Section   Required. String expression containing the name of the section where the key setting is found.
   '           If omitted, key setting is assumed to be in default subkey.
   ' Key       Required. String expression containing the name of the key setting to return.
   ' Default   Optional. Expression containing the value to return if no value is set in the key setting.
   '           If omitted, default is assumed to be a zero-length string ("").
   ' ADAPTED BY JENNESS FROM Ask the VB Pro column, [Getting Started with VB", Spring 1998.]
   ' SET DEFAULT VALUE
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
   
   ' Assume failure and set return to Default
'   GetSetting = Default

   ' Open key
   nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey("ESRI", "ArcGIS", strSection), 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If strKey = "*" Then strKey = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer and get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL and return successful query!
               strDir = Left(Buffer, nBytes - 1)
               ReturnArcGISVersion = strDir
            End If
         End If
      End If
      Call RegCloseKey(hKey)
   End If

End Function

Public Function ReturnArcGISInstallDir() As String

' ORIGINAL FUNCTION:
' Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
   ' Section   Required. String expression containing the name of the section where the key setting is found.
   '           If omitted, key setting is assumed to be in default subkey.
   ' Key       Required. String expression containing the name of the key setting to return.
   ' Default   Optional. Expression containing the value to return if no value is set in the key setting.
   '           If omitted, default is assumed to be a zero-length string ("").
   ' ADAPTED BY JENNESS FROM Ask the VB Pro column, [Getting Started with VB", Spring 1998.]
   ' SET DEFAULT VALUE
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
   
   
   ' Assume failure and set return to Default
'   GetSetting = Default

   ' Open key
   nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey("ESRI", "ArcGIS", strSection), 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If strKey = "*" Then strKey = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer and get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL and return successful query!
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

' ORIGINAL FUNCTION:
' Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
   ' Section   Required. String expression containing the name of the section where the key setting is found.
   '           If omitted, key setting is assumed to be in default subkey.
   ' Key       Required. String expression containing the name of the key setting to return.
   ' Default   Optional. Expression containing the value to return if no value is set in the key setting.
   '           If omitted, default is assumed to be a zero-length string ("").
   ' ADAPTED BY JENNESS FROM Ask the VB Pro column, [Getting Started with VB", Spring 1998.]
   ' SET DEFAULT VALUE
   
   
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
   
  ' Assume failure and set return to Default
  '   GetSetting = Default

   ' Open key
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If strKey = "*" Then strKey = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer and get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, strKey, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL and return successful query!
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

   ' Build SubKey from known values
   SubKey = "Software\" & strCompany & "\" & strAppName
   If Len(Section) Then
      SubKey = SubKey & "\" & Section
   End If

End Function
Public Function ReturnFolders(ByVal DirPath As String, Optional ByVal FolderName As String) As Collection

  ' RETURNS A COLLECTION CONTAINING:
  '    A) Count of files that met criteria
  '    B) String Array of full folder filenames
  
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

  ' RETURNS A COLLECTION CONTAINING:
  '    A) Count of files that met criteria
  '    B) String Array of full filenames
  
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
'  If anIndex = 0 Then
'    returnCollection.Add 0
'  Else
'    returnCollection.Add (UBound(strFiles) - LBound(strFiles) + 1)
'  End If
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
  
  ' The document is always the last item
  GetMxDocPath = pTemplates.Item(lTempCount - 1)

End Function

Public Function CreatedBASETable(strFullName As String, Optional pFields As IFields) As ITable

' createDBF: simple function to create a DBASE file.
' note: the name of the DBASE file should not contain the .dbf extension
' ESRI Sample; modified by Jenness August 20 2007
  
  Dim strName As String
  Dim strFolder As String
  
  strFolder = aml_func_mod.ReturnDir(strFullName)
  strName = aml_func_mod.ReturnFilename(strFullName)
  If Right(strName, 4) = ".dbf" Then strName = Left(strName, Len(strName) - 4)
  
  ' Open the Workspace
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
  
  ' if a fields collection is not passed in then create one
  If pFields Is Nothing Then
    ' create the fields used by our object
    Set pFields = New Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 1
    
    'Create text Field
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


'  MsgBox pField.Name
'  MsgBox pField.Scale
'  MsgBox pField.Precision
'  MsgBox pField.Type

'  For lngIndex = 0 To pFields.FieldCount - 1
'    strString = strString + "-----------------" & "  Field Name = " & pFields.Field(lngIndex).Name & vbCrLf & _
'        "  Field Scale = " & pFields.Field(lngIndex).Scale & vbCrLf & "  Precision = " & pFields.Field(lngIndex).Precision & _
'        vbCrLf & "  Field Type = " & CStr(pFields.Field(lngIndex).Type) & vbCrLf
'  Next lngIndex
'  MsgBox "Problem with workspace? " & CStr(pFWS Is Nothing) & vbCrLf & "Filename = " & strFullName & vbCrLf & _
'        "Folder = " & strFolder & vbCrLf & "strName = " & strName & vbCrLf & pFields.FieldCount & " fields..." & vbCrLf & _
'        "File already exists? " & CBool(FileExists(strFullName)) & vbCrLf & strString
  
  Set CreatedBASETable = pFWS.CreateTable(strName, pFields, Nothing, Nothing, "")

End Function


Public Function CreateShapefile(sPath As String, sName As String, pSpatialReference As ISpatialReference, strShapeType As String) As IFeatureClass   ' Don't include filename!
  
  If Right(sPath, 4) = ".shp" Then sPath = ReturnDir(sPath)
  If Right(sName, 4) = ".shp" Then sName = Left(sName, Len(sName) - 4)
  
  ' SET GEOMETRY TYPE, AND EXIT IF NOT ONE OF STANDARD OPTIONS
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    Select Case strShapeType
      Case "Polygon", "polygon"
        .GeometryType = esriGeometryPolygon
      Case "Polyline", "polyline"
        .GeometryType = esriGeometryPolyline
      Case "Point", "point"
        .GeometryType = esriGeometryPoint
      Case "Multipoint", "multipoint", "MultiPoint"
        .GeometryType = esriGeometryMultipoint
      Case "Multipatch", "multipatch", "MultiPatch"
        .GeometryType = esriGeometryMultiPatch
      Case Else
        MsgBox "Invalid Shape Type [" & strShapeType & "]!  This function is only written to generate " & _
            "Point, Polyline, Polygon, Multipoint or Multipatch shapefiles...", vbCritical, "Invalid Shape Type:"
    End Select
'    Set .SpatialReference = New UnknownCoordinateSystem
    Set .SpatialReference = pSpatialReference
  End With
  
  ' Open the folder to contain the shapefile as a workspace
  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
'   If Not pWorkspaceFactory.IsWorkspace(sPath) Then
  If Not ExistFileDir(sPath) Then
    MsgBox "Unable to create Feature Class:" & vbCrLf & _
           sPath & " is not a valid workspace...", , "Failed to Create Feature Class:"
    Set CreateShapefile = Nothing
    Exit Function
  End If
  
  Set pFWS = pWorkspaceFactory.OpenFromFile(sPath, 0)
  
  ' Set up a simple fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields
  
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  ' Make the shape field
  ' it will need a geometry definition, with a spatial reference
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Shape"
  pFieldEdit.Type = esriFieldTypeGeometry
  
  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField

  ' Add an ID field
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
      .length = 8
      .Name = "Unique_ID"
      .Type = esriFieldTypeInteger
      .Precision = 0
  End With
  pFieldsEdit.AddField pField
  
  ' Create the shapefile
  ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".shp"
'    MsgBox sPath & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & sName & ".shp") <> "")
  Else
    strCheckString = sPath & "\" & sName & ".shp"
'    MsgBox sPath & "\" & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & "\" & sName & ".shp") <> "")
  End If
  
  booFileExists = (Dir(strCheckString) <> "")
  
  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreateShapefile = Nothing
    Exit Function
  End If
  
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(sName, pFields, Nothing, _
                                           Nothing, esriFTSimple, "Shape", "")
                                           
  Set CreateShapefile = pFeatClass

End Function

Public Function CreateShapefile2(sPath As String, sName As String, pSpatialReference As ISpatialReference, _
    strShapeType As String, pAddFields As esriSystem.IVariantArray) As IFeatureClass     ' Don't include filename!
  
  If Right(sPath, 4) = ".shp" Then sPath = ReturnDir(sPath)
  If Right(sName, 4) = ".shp" Then sName = Left(sName, Len(sName) - 4)
  
  ' SET GEOMETRY TYPE, AND EXIT IF NOT ONE OF STANDARD OPTIONS
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    Select Case strShapeType
      Case "Polygon", "polygon"
        .GeometryType = esriGeometryPolygon
      Case "Polyline", "polyline"
        .GeometryType = esriGeometryPolyline
      Case "Point", "point"
        .GeometryType = esriGeometryPoint
      Case "Multipoint", "multipoint", "MultiPoint"
        .GeometryType = esriGeometryMultipoint
      Case "Multipatch", "multipatch", "MultiPatch"
        .GeometryType = esriGeometryMultiPatch
      Case Else
        MsgBox "Invalid Shape Type [" & strShapeType & "]!  This function is only written to generate " & _
            "Point, Polyline, Polygon, Multipoint or Multipatch shapefiles...", vbCritical, "Invalid Shape Type:"
    End Select
'    Set .SpatialReference = New UnknownCoordinateSystem
    Set .SpatialReference = pSpatialReference
  End With
  
  ' Open the folder to contain the shapefile as a workspace
  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
'   If Not pWorkspaceFactory.IsWorkspace(sPath) Then
  If Not ExistFileDir(sPath) Then
    MsgBox "Unable to create Feature Class:" & vbCrLf & _
           sPath & " is not a valid workspace...", , "Failed to Create Feature Class:"
    Set CreateShapefile2 = Nothing
    Exit Function
  End If
  
  Set pFWS = pWorkspaceFactory.OpenFromFile(sPath, 0)
  
'   Set up a simple fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  ' Make the shape field
  ' it will need a geometry definition, with a spatial reference
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Shape"
  pFieldEdit.Type = esriFieldTypeGeometry

  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField
  
  ' ADD FIELDS
  Dim lngIndex As Long
  For lngIndex = 0 To pAddFields.Count - 1
    pFieldsEdit.AddField pAddFields.Element(lngIndex)
  Next lngIndex
    
  ' Create the shapefile
  ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".shp"
'    MsgBox sPath & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & sName & ".shp") <> "")
  Else
    strCheckString = sPath & "\" & sName & ".shp"
'    MsgBox sPath & "\" & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & "\" & sName & ".shp") <> "")
  End If
  
  booFileExists = (Dir(strCheckString) <> "")
  
  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreateShapefile2 = Nothing
    Exit Function
  End If
  
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(sName, pFields, Nothing, _
                                           Nothing, esriFTSimple, "Shape", "")
                                           
  Set CreateShapefile2 = pFeatClass

End Function


Public Function TempPathLocation() As String

  Dim sBuffer As String
  sBuffer = Space(strMAXPATH)
  If GetTempPath(strMAXPATH, sBuffer) <> 0 Then
    TempPathLocation = Left$(sBuffer, _
      InStr(sBuffer, vbNullChar) - 1)
  Else
    TempPathLocation = ""
  End If

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

Public Function InsertCommas(InputValue As Variant) As String
  
  Dim theString As String
  theString = CStr(InputValue)
  
  Dim theDecLocation As Long
  theDecLocation = InStr(theString, ".")
  
  Dim HasDecimal As Boolean
  HasDecimal = theDecLocation > 0
  
  Dim theLength As Long
  theLength = Len(theString)
  
  Dim theBaseNumber As String
  Dim theRemainder As String
  
  If HasDecimal Then
    theRemainder = Right(theString, theLength - theDecLocation)
    theBaseNumber = Left(theString, theDecLocation - 1)
  Else
    theRemainder = ""
    theBaseNumber = theString
  End If
  
  Dim theCount As Long
  theCount = Len(theBaseNumber)
  
  Dim theCommaString As String
  
  If theCount > 3 Then
    Dim anIndex As Long
    For anIndex = (theCount - 2) To 1 Step -3
      theCommaString = Mid(theBaseNumber, anIndex, 3) & "," & theCommaString
      If anIndex < 4 Then
        theCommaString = Left(theBaseNumber, anIndex - 1) & "," & theCommaString
      End If
    Next anIndex
    
    Do While Right(theCommaString, 1) = ","
      theCommaString = Left(theCommaString, Len(theCommaString) - 1)
    Loop
    Do While Left(theCommaString, 1) = ","
      theCommaString = Right(theCommaString, Len(theCommaString) - 1)
    Loop
  Else
    theCommaString = theBaseNumber
  End If
  
  If HasDecimal Then
    theCommaString = theCommaString & "." & theRemainder
  End If
  
  InsertCommas = theCommaString

End Function

Public Function ClipExtension2(strPathname As String) As String
  
  Dim lngLastDot As Long
  Dim lngLastSlash As Long
  Dim lngLastForwardSlash As Long
  
  lngLastDot = InStrRev(strPathname, ".")
  lngLastSlash = InStrRev(strPathname, "\")
  lngLastForwardSlash = InStrRev(strPathname, "/")
  
  Dim strSplit() As String
  Dim strFinalPath As String
  Dim lngIndex As Long
  
  If lngLastDot > 0 And lngLastDot > lngLastSlash And lngLastDot > lngLastForwardSlash Then
    strSplit = Split(strPathname, ".")
    strFinalPath = strSplit(0)
    For lngIndex = 1 To UBound(strSplit) - 1
      strFinalPath = strFinalPath & "." & strSplit(lngIndex)
    Next lngIndex
    ClipExtension2 = strFinalPath
  Else
    ClipExtension2 = strPathname
  End If
  
  Erase strSplit

End Function
Public Function ClipExtension(strPathname As String) As String
  
'  Dim lngLastDot As Long
'  Dim lngLastSlash As Long
'  Dim lngLastForwardSlash As Long
'
'  lngLastDot = InStrRev(strPathname, ".")
'  lngLastSlash = InStrRev(strPathname, "\")
'  lngLastForwardSlash = InStrRev(strPathname, "/")
'
'  Dim strSplit() As String
'  Dim strFinalPath As String
'  Dim lngIndex As Long
'
'  If lngLastDot > 0 And lngLastDot > lngLastSlash And lngLastDot > lngLastForwardSlash Then
'    strSplit = Split(strPathname, ".")
'    strFinalPath = strSplit(0)
'    For lngIndex = 1 To UBound(strSplit) - 1
'      strFinalPath = strFinalPath & "." & strSplit(lngIndex)
'    Next lngIndex
'    ClipExtension = strFinalPath
'  Else
'    ClipExtension = strPathname
'  End If
'
'  Erase strSplit
'
  Dim strDirPath As String
  Dim strDirTokens() As String

  aml_func_mod.ParseString strPathname, strDirTokens, "."
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

Public Function FieldIsNumeric(pTheField As iField) As Boolean
  
  Dim theFieldType As esriFieldType
  theFieldType = pTheField.Type
  
  FieldIsNumeric = _
    (theFieldType = esriFieldTypeSmallInteger) Or (theFieldType = esriFieldTypeDouble) Or (theFieldType = esriFieldTypeInteger) Or _
          (theFieldType = esriFieldTypeSingle)

End Function
Public Function FieldIsString(pTheField As iField) As Boolean
  
  Dim theFieldType As esriFieldType
  theFieldType = pTheField.Type
  
  FieldIsString = (theFieldType = esriFieldTypeString)

End Function
Public Function FieldIsDate(pTheField As iField) As Boolean
  
  Dim theFieldType As esriFieldType
  theFieldType = pTheField.Type
  
  FieldIsDate = (theFieldType = esriFieldTypeDate)

End Function
Public Function FieldIsShape(pTheField As iField) As Boolean
  
  Dim theFieldType As esriFieldType
  theFieldType = pTheField.Type
  
  FieldIsShape = (theFieldType = esriFieldTypeGeometry)

End Function

Public Function SetExtension(strPathname As String, strExtension As String) As String
  
  Dim theClippedPath As String
  SetExtension = ClipExtension(strPathname) & "." & strExtension

End Function

Public Function GetExtensionText(strPathname As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String
  
  aml_func_mod.ParseString strPathname, strDirTokens, "."
  If UBound(strDirTokens) = 0 Then
    GetExtensionText = ""
  Else
    GetExtensionText = strDirTokens(UBound(strDirTokens))
  End If

End Function

Public Function GetFullFileString(str83Type As String) As String
  
  ' ADAPTED FROM BRETT MERONEY'S POST ABOVE
  
  Dim lLen As Long
  Dim sBuffer As String
  
  sBuffer = String$(strMAXPATH, 0)
  lLen = GetLongPathName(str83Type, sBuffer, Len(sBuffer))
  If lLen > 0 And err.Number = 0 Then
    GetFullFileString = Left$(sBuffer, lLen)
  Else
    GetFullFileString = str83Type
  End If

End Function

Public Function ReturnDir(strPathname As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String
  
  If InStr(1, strPathname, "\") = 0 Then
    ReturnDir = ""
  Else
    
    aml_func_mod.ParseString strPathname, strDirTokens, "\"
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


Public Function ReturnFilename(strPathname As String) As String

  Dim strDirPath As String
  Dim strDirTokens() As String
  
  If InStr(1, strPathname, "\") = 0 Then
    ReturnFilename = strPathname
  Else
  
    aml_func_mod.ParseString strPathname, strDirTokens, "\"
    ReturnFilename = strDirTokens(UBound(strDirTokens))
  End If

End Function
Public Sub ParseString(str As String, strArray() As String, Delim As String)

' Populates a named string array with elements in a string. Each array element
' contains one word. Multiple words ' within single quotes ' are treated as
' one word. NOTE: Use parseStringR if you want both commas and blanks to be
' treated as delimiting characters.
'
' Before calling this Sub you must declare your array in the calling program
' As String, with no bounds.
'
' Dim myarray() As String
' mystring = "ARC YES, POLY NO, TICS YES"
' parseString (mystring), myarray, ","
' Returns:
' array(0) = ARC YES
' array(1) = POLY NO
' array(2) = TICS YES

' parseString(mystring),myarray," "
' Returns:
' array(0) = ARC
' array(1) = YES,
' array(2) = POLY
' array(3) = NO,
' array(4) = TICS
' array(5) = YES,

' mystring = "'Universe,Medium','Helvetica,Bold','Times,Medium'"
' parsestring(mystring),myarray,","
' Returns:
' array(0) = Universe,Medium
' array(1) = Helvetica,Bold
' array(2) = Times,Medium

' Dim counters

Dim i As Long
Dim tokenlen As Long
Dim tmpstr As String
Dim position As Long
Dim length As Long

'Dim variables to keep track of embedded quotes

Dim switch As Long
Dim position1 As Long
Dim position2 As Long
Dim pair As Long

On Error Resume Next

' If string contains no elements raise error
  If Trim(Subst(str, Delim)) = "" Then
    err.Raise vbObjectError + 1, "aml_func.ParseString", _
    "StringPassed"
    Exit Sub
  End If

' intialize array. Warning: This will overwrite any data elements currently
' stored in this named array

  ReDim strArray(0)

'intializer counters and tracking variables

  pair = False
  switch = 0
  length = Len(str)
  position = 1
  i = 0
  tmpstr = str

'check each character in the array. If it is a quote, store if it is first or last
' 0 = havent read one yet
' 1 = read first single quote
' 2 = read second single quote

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

' if last char read was last in a pair of quotes, store contents between first and last
' in current array element and reset tracking variables

    If pair = True Then
      Mid(tmpstr, position1, 1) = " "
      Mid(tmpstr, position2, 1) = " "
      strArray(i) = Mid(tmpstr, position1, position2 - position1)
      strArray(i) = Trim(strArray(i))
      pair = False
      switch = 0
      
' check to see if we are reading till the next single quote. If switch = 0, we are not
' if not check if the next character is a delimiter. if it is store everything to the left
' replace everything to the left of original str with blanks so we can safely use LEFT
' function then trim the blanks

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

' we have populated our array, now remove the delimiters from each element

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

' Returns the substring of Str to the right of the leftmost
' occurrence of the searchStr.
Dim position As Long
Dim length As Long

  position = InStr(str, SearchStr)
  length = Len(SearchStr)
  If Not (position = 0) Then
   After = Mid(str, position + length)
  End If

End Function

Public Function Before(str As String, SearchStr As String) As String

' Returns the substring of Str to the left of the leftmost
' occurrence of the searchStr.
Dim position As Long
Dim length As Long

  position = InStr(str, SearchStr)
  length = Len(SearchStr)
  If Not (position = 0) Then
   Before = Mid(str, 1, position - 1)
  End If

End Function

Function ExistFileDir(sTest As String) As Boolean

'Checks for the existance of a File or Directory
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

Public Function Extract(ElemNum As Long, ElemList As String) As String

' extracts an element from a list of elements

Dim strArray() As String

  ParseStringR (ElemList), strArray
  If ElemNum > UBound(strArray) + 1 Then
   Exit Function
  End If
  
  If ElemNum = 0 Then
   Exit Function
  End If
  Extract = strArray(ElemNum - 1)

End Function

Public Function Index(str As String, SearchStr As String) As Long

' Returns the position of the leftmost occurrence of searcStr in str.

  Index = InStr(str, SearchStr)

End Function

Public Sub ParseStringR(str As String, strArray() As String, Optional ReturnQuoted)

' Populates a named string arrary with elements in a string. Each array element
' contains one word. Multiple words ' within single quotes ' are treated as
' one word. Treats both blanks and commas as delimters NOTE: Use parseString
' to specify a specific delimiting character.

' ReturnQuoted - indicates if elements are to be returned quoted.
' FALSE - DEFAULT return elements unquoted
' TRUE - return elements quoted

' Before calling this function you must declare your array in the calling program
' As String, with no bounds.
'
' Dim myarray() As String
' mystring = "ARC YES, POLY NO, TICS YES"
' parseStringR(mystring),myarray

' Returns:
' array(0) = ARC
' array(1) = YES
' array(2) = POLY
' array(3) = NO
' array(4) = TICS
' array(5) = YES

' mystring = "'Universe,Medium','Helvetica,Bold','Times,Medium'"

' parsestringR(mystring),myarray
' Returns:
' array(0) = Universe,Medium
' array(1) = Helvetica,Bold
' array(2) = Times,Medium

' parsestring(mystring,myarray,TRUE)
' Returns:
' array(0) = 'Universe,Medium'
' array(1) = 'Helvetica,Bold'
' array(2) = 'Times,Medium'

' Dim counters

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

' If string contains no elements raise error
  If Trim(Subst(str, ",")) = "" Then
    err.Raise vbObjectError + 1, "aml_func.ParseStringR", _
    "StringPassed"
    Exit Sub
  End If

' intialize counters and tracking variables
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

' check each character in the array. If it is a quote, store if it is first or last
' 0 = havent read one yet
' 1 = read first single quote
' 2 = just read second single quote

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
     
' if last char read was last in a pair of single quotes, store contents between first
' and last in current array element and reset tracking variables
 
    If pair = True Then
      Mid(tmpstr, position1, 1) = " "
      Mid(tmpstr, position2, 1) = " "
      strArray(i) = Mid(tmpstr, position1, position2 - position1)
      strArray(i) = Trim(strArray(i))
      pair = False
      switch = 0
    
' check to see if we are reading till the next single quote. If switch = 0, we are not
' if not check if the next character is a delimiter. if it is store everything to the left
' replace everything to the left of original str with blanks so we can safely use LEFT
' function then trim the blanks
  
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

' we have populated our array, now remove the delimiters from each element
' set parseAgain flag if there are any blank elements

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
  
' now remove any blank elements
  
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

Public Function Keyword(str As String, SearchStr As String) As Long

' Returns the position of a string within a list of keywords.
' converts Str and searchStr to upper case before comparing
'  0  if keyword not found
' -1  if string is ambiguous - mutiple occurances of same keyword
'  n  position of keyword in string

Dim strArray() As String
Dim i As Long
Dim keywordCnt As Long

  ParseStringR (str), strArray
  keywordCnt = 0
    For i = 0 To UBound(strArray)
    If UCase(SearchStr) = UCase(strArray(i)) Then
      Keyword = i + 1
      keywordCnt = keywordCnt + 1
    End If
    Next i
  
  If keywordCnt > 1 Then
    Keyword = -1
  End If

End Function

Public Function Search(str, SearchStr) As Long

' Returns the position of the first character in Str
' which occurs in searchStr.

Dim strArray() As String
Dim i As Long
Dim Index As Long
Dim firstchar
Dim InString As Boolean

  Index = 1
  ReDim strArray(Len(str))

  For i = 0 To Len(SearchStr) - 1
    firstchar = Mid(SearchStr, i + 1, 1)
    Index = InStr(str, firstchar)
    strArray(Index) = i + 1
  Next i
  
  InString = False
  
  For i = 1 To UBound(strArray)
    If Not (strArray(i)) = "" Then
    InString = True
    Exit For
    Else
    End If
  Next i
  
  If InString = False Then
    Search = 0
  Else
    Search = i
  End If

End Function

Public Function Sort(str As String, Optional SortOption, Optional SortType) As String

' Returns a string of sorted elements

  If IsMissing(SortOption) Then
    SortOption = "-ASCEND"
  ElseIf Not (UCase(SortOption) = "-DESCEND") Then
    SortOption = "-ASCEND"
  End If
  
  If IsMissing(SortType) Then
    SortType = "-CHARACTER"
  ElseIf Not (UCase(SortType) = "-NUMERIC") Then
    SortType = "-CHARACTER"
  End If
  
  If (UCase(SortType)) = "-NUMERIC" Then
    Call Sort_Num(str, SortOption)
  Else
    Call Sort_Char(str, SortOption, True)
  End If
    Sort = str

End Function

Private Function Sort_Num(str As String, SortOption) As String

' Sort function - performs a numerical sort
' Ref: Selectionsort Chapter8 of VB Algorithms; Rod Stephens

Dim i As Long
Dim j As Long
Dim min As Long
Dim max As Long
Dim best_value As String
Dim best_j As Long
Dim sortArray() As String
Dim sorted As String

  ParseStringR (str), sortArray
  
  min = LBound(sortArray)
  max = UBound(sortArray)
  
  For i = min To max - 1
    best_value = sortArray(i)
    best_j = i
    
    For j = i + 1 To max
      If Val(sortArray(j)) < Val(best_value) Then
        best_value = sortArray(j)
        best_j = j
      End If
    Next j
      
    sortArray(best_j) = sortArray(i)
    sortArray(i) = best_value
  Next i
    
  If UCase(SortOption) = "-DESCEND" Then
    For i = max To min Step -1
      sorted = sorted & sortArray(i) & ","
    Next i
  Else
    For i = min To max
      sorted = sorted & sortArray(i) & ","
    Next i
  End If
  
  Mid(sorted, Len(sorted), 1) = " "
  str = sorted
  Sort_Num = sorted

End Function

Private Function Sort_Char(str As String, SortOption, Optional ReturnQuoted) As String

' Sort function - performs a character sort
' Ref: Selectionsort Chapter8 of VB Algorithms; Rod Stephens

Dim i As Long
Dim j As Long
Dim min As Long
Dim max As Long
Dim best_value As String
Dim best_j As Long
Dim sortArray() As String
Dim sorted As String

If IsMissing(ReturnQuoted) Then
  ReturnQuoted = False
End If
If Not (ReturnQuoted = False) Then
  ReturnQuoted = True
End If

ParseStringR (str), sortArray, ReturnQuoted

min = LBound(sortArray)
max = UBound(sortArray)

  For i = min To max - 1
    best_value = sortArray(i)
    best_j = i
    
    For j = i + 1 To max
      If sortArray(j) < best_value Then
      best_value = sortArray(j)
      best_j = j
      End If
    Next j
    
    sortArray(best_j) = sortArray(i)
    sortArray(i) = best_value
  Next i
  
  If UCase(SortOption) = "-DESCEND" Then
    For i = max To min Step -1
      sorted = sorted & sortArray(i) & " "
    Next i
  Else
    For i = min To max
      sorted = sorted & sortArray(i) & " "
    Next i
  End If
  Mid(sorted, Len(sorted), 1) = " "
  
  str = sorted
  Sort_Char = sorted

End Function

Public Function Subst(str As String, SearchChar As String, Optional ReplaceChar) As String

' Replaces all occurances of specified char in string.

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

Public Function Substr(str As String, position As Long, Optional NumChars) As String

'extracts a substring starting at a specified character position.

If IsMissing(NumChars) Then
  If position = 0 Or position > Len(str) Then
    Substr = ""
  Else
    Substr = Mid(str, position)
  End If
Else
  If position = 0 Or position > Len(str) Then
    Substr = ""
  Else
   Substr = Mid(str, position, NumChars)
  End If
End If

End Function

Public Function Token(ElemList As String, Arg As String, ParamArray OtherArgs()) As Variant

' Performs various functions for string manipulation

Dim strArray() As String
Dim i As Long
Dim temp As String
Dim from_elem  As Long
Dim to_elem As Long
Dim start_elem As Long
Dim insertStr As String
Dim DELETE As Long
Dim SearchStr As String

' Parse ElemList out to strarray
' Select TOKEN argument and perform function

  ParseStringR (ElemList), strArray
  Arg = Subst(Arg, "-")
  
  Select Case UCase(Arg)

  ' Count - returns the number of tokens in a list
    Case "COUNT"
      Token = UBound(strArray) + 1
    
  ' Find <token> - returns the position of a token in a list
    Case "FIND"
      SearchStr = OtherArgs(0)
      Token = 0
      For i = 0 To UBound(strArray)
        If UCase(SearchStr) = UCase(strArray(i)) Then
          Token = i + 1
        End If
      Next i
    
  ' Move <from_position> <to_position> - moves a token in the list
    Case "MOVE"
      from_elem = OtherArgs(0) - 1
      to_elem = OtherArgs(1)
      temp = strArray(to_elem)
      strArray(to_elem) = strArray(from_elem)
      For i = from_elem To to_elem - 1
        strArray(i) = strArray(i + 1)
        Next i
         strArray(i) = temp
      For i = 0 To UBound(strArray) - 1
        Token = Token & strArray(i) & ","
      Next i
    
  ' Insert <position> - inserts a new token at <position> in the list
    Case "INSERT"
      ReDim Preserve strArray(UBound(strArray) + 1)
      start_elem = OtherArgs(0) - 1
      insertStr = OtherArgs(1)
      For i = UBound(strArray) To start_elem Step -1
        strArray(i) = strArray(i - 1)
        Next i
        strArray(start_elem) = insertStr
        For i = 0 To UBound(strArray) - 1
        Token = Token & strArray(i) & ","
      Next i
    
  ' Delete <position> - removes the token at <position> from the list.
    Case "DELETE"
      DELETE = OtherArgs(0) - 1
      For i = DELETE To UBound(strArray) - 1
        strArray(i) = strArray(i + 1)
      Next i
      For i = 0 To UBound(strArray) - 1
        Token = Token & strArray(i) & ","
      Next i
    
  ' Replace <position> <new_string> - replaces the token at <position> with the
  ' <new_string>.
    Case "REPLACE"
      strArray(OtherArgs(1) - 1) = OtherArgs(0)
      For i = 0 To UBound(strArray)
        Token = Token & strArray(i) & ","
      Next i
    
  ' Switch <position_1> <position_2> - moves token at <position_1> to <position_2> and moves
  ' token at <position_2> to <position_1>.
    Case "SWITCH"
      from_elem = OtherArgs(0) - 1
      to_elem = OtherArgs(1) - 1
      temp = strArray(to_elem)
      strArray(to_elem) = strArray(from_elem)
      strArray(from_elem) = temp
      For i = 0 To UBound(strArray)
        Token = Token & strArray(i) & ","
      Next i
    
    Case Else
  End Select

End Function


Public Function PathIsDirectory(strPath As String) As Boolean

  On Error GoTo ErrHandler:
  PathIsDirectory = GetAttr(strPath) = vbDirectory
  Exit Function
ErrHandler:
  PathIsDirectory = False
End Function


Public Function ReturnDir2(strPathname As String) As String

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  lngIndex1 = InStrRev(strPathname, "\", , vbTextCompare)
  lngIndex2 = InStrRev(strPathname, "/", , vbTextCompare)
  
  If lngIndex1 = 0 And lngIndex2 = 0 Then
    ReturnDir2 = strPathname
  Else
    If lngIndex1 = 0 Then
      ReturnDir2 = Left(strPathname, lngIndex2)
    Else
      ReturnDir2 = Left(strPathname, lngIndex1)
    End If
  End If

End Function


Public Function ReturnFilename2(strPathname As String) As String

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  lngIndex1 = InStrRev(strPathname, "\", , vbTextCompare)
  lngIndex2 = InStrRev(strPathname, "/", , vbTextCompare)
  
  If lngIndex1 = 0 And lngIndex2 = 0 Then
    ReturnFilename2 = strPathname
  Else
    If lngIndex1 = 0 Then
      ReturnFilename2 = Right(strPathname, Len(strPathname) - lngIndex2)
    Else
      ReturnFilename2 = Right(strPathname, Len(strPathname) - lngIndex1)
    End If
  End If

End Function
Public Function ReturnDir3(strPathname As String, Optional booPutTrailingBackslash As Boolean = True) As String

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  lngIndex1 = InStrRev(strPathname, "\", , vbTextCompare)
  lngIndex2 = InStrRev(strPathname, "/", , vbTextCompare)
  
  If lngIndex1 = 0 And lngIndex2 = 0 Then
    ReturnDir3 = strPathname
  Else
    If lngIndex1 = 0 Then
      ReturnDir3 = Left(strPathname, lngIndex2)
    Else
      ReturnDir3 = Left(strPathname, lngIndex1)
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


