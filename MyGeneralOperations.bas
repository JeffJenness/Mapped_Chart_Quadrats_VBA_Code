Attribute VB_Name = "MyGeneralOperations"
Option Explicit

Public Enum JenDatasetTypes
  ENUM_Shapefile = 1
  ENUM_FileGDB = 2
  ENUM_PersonalGDB = 4
  ENUM_Coverage = 8
  ENUM_SDC_FeatureClass = 16
  ENUM_File_Raster = 32
End Enum

Public Enum Jen_ElementEnvPoint
  ENUM_Upper_Left = 1
  ENUM_Upper_Right = 2
  ENUM_Lower_Left = 3
  ENUM_Lower_Right = 4
  ENUM_Upper_Center = 5
  ENUM_Lower_Center = 6
  ENUM_Center_Left = 7
  ENUM_Center_Right = 8
  ENUM_Center_Center = 9
  ENUM_By_Percentages = 10
End Enum

Public Function ReplaceBadChars(strName As String, Optional booReplacePeriod As Boolean = False, _
    Optional booReplaceBackSlash As Boolean = False, Optional booReplaceSpace As Boolean = True, _
    Optional booStartWithLetter As Boolean = False) As String

  If strName = "" Then
    ReplaceBadChars = "z"
  Else
    Dim strAcceptable As String
    strAcceptable = "abcdefghijklmnopqrstuvwxyz0123456789_"
    If Not booReplacePeriod Then strAcceptable = strAcceptable & "."
    If Not booReplaceBackSlash Then strAcceptable = strAcceptable & "\"
    If Not booReplaceSpace Then strAcceptable = strAcceptable & " "

    Dim strOutput As String
    strOutput = strName
    Dim lngIndex As Long
    Dim strChar As String
    For lngIndex = 1 To Len(strOutput)
      strChar = Mid(strOutput, lngIndex, 1)
      If InStr(1, strAcceptable, strChar, vbTextCompare) = 0 Then
        strOutput = Replace(strOutput, strChar, "_")
      End If
    Next lngIndex
    If booStartWithLetter Then
      Dim strLetters As String
      strLetters = "abcdefghijklmnopqrstuvwxyz"
      If InStr(1, strLetters, Left(strOutput, 1), vbTextCompare) = 0 Then
        strOutput = "z" & strOutput
      End If
    End If
    ReplaceBadChars = strOutput
  End If

End Function

Public Function CheckIfFeatureClassExists(pWS As IFeatureWorkspace, strName As String) As Boolean

    Dim pFClass As IFeatureClass
    On Error Resume Next
    Set pFClass = pWS.OpenFeatureClass(strName)
    If pFClass Is Nothing Then
        CheckIfFeatureClassExists = False
    Else
        CheckIfFeatureClassExists = True
    End If
    If CheckIfFeatureClassExists = False Then
      Dim pTable As ITable
      Set pTable = pWS.OpenTable(strName)
      If pTable Is Nothing Then
          CheckIfFeatureClassExists = False
      Else
          CheckIfFeatureClassExists = True
      End If
    End If

  GoTo ClearMemory
ClearMemory:
  Set pFClass = Nothing
  Set pTable = Nothing

End Function

Public Function MakeUniquedBASEName(strFilename As String) As String

  If Not ExistFileDir(strFilename) Then
    MakeUniquedBASEName = strFilename
    Exit Function
  Else

    Dim theCounter As Long
    theCounter = 1

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

    theBaseName = theFilename

    Do While ExistFileDir(theFilename & theExtension)
      theCounter = theCounter + 1
      theFilename = theBaseName & "_" & CStr(theCounter)
    Loop

    MakeUniquedBASEName = theFilename & theExtension

  End If

End Function

Public Function CreateGDBTable(featWorkspace As IFeatureWorkspace, _
                                            Name As String, _
                                            Optional pAddFields As esriSystem.IVariantArray, _
                                            Optional pCLSID As UID, _
                                            Optional pCLSEXT As UID, _
                                            Optional ConfigWord As String = "" _
                                            ) As ITable

  Set CreateGDBTable = Nothing
  If featWorkspace Is Nothing Then Exit Function
  If Name = "" Then Exit Function

  If (pCLSID Is Nothing) Or IsMissing(pCLSID) Then
    Set pCLSID = Nothing
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Object"
  End If

  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit

  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Object_ID"
  pFieldEdit.Type = esriFieldTypeOID
  pFieldsEdit.AddField pField

  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If

  If (pCLSEXT Is Nothing) Or IsMissing(pCLSEXT) Then
    Set pCLSEXT = Nothing
  End If

  Set CreateGDBTable = featWorkspace.CreateTable(Name, pFields, pCLSID, pCLSEXT, "")

  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing

End Function

Public Function CreateGDBFeatureClass2(featWorkspace As IFeatureWorkspace, _
                                            Name As String, _
                                            featType As esriFeatureType, _
                                            pSpRef As ISpatialReference, _
                                            Optional geomType As esriGeometryType = esriGeometryPoint, _
                                            Optional pAddFields As esriSystem.IVariantArray, _
                                            Optional pCLSID As UID, _
                                            Optional pCLSEXT As UID, _
                                            Optional ConfigWord As String = "", _
                                            Optional booForceUniqueIDField As Boolean = True, _
                                            Optional lngCategory As JenDatasetTypes, _
                                            Optional pExtent As IEnvelope, _
                                            Optional lngOriginalRecCount As Long = -9999, _
                                            Optional booHasZ As Boolean, _
                                            Optional booHasM As Boolean) As IFeatureClass

  Set CreateGDBFeatureClass2 = Nothing
  If featWorkspace Is Nothing Then Exit Function
  If Name = "" Then Exit Function

  If (pCLSID Is Nothing) Or IsMissing(pCLSID) Then
    Set pCLSID = Nothing
    Set pCLSID = New UID

    Select Case featType
      Case esriFTSimple
        pCLSID.Value = "esriGeoDatabase.Feature"
        If geomType = esriGeometryLine Then geomType = esriGeometryPolyline
      Case esriFTSimpleJunction
        geomType = esriGeometryPoint
        pCLSID.Value = "esriGeoDatabase.SimpleJunctionFeature"
      Case esriFTComplexJunction
        pCLSID.Value = "esriGeoDatabase.ComplexJunctionFeature"
      Case esriFTSimpleEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.SimpleEdgeFeature"
      Case esriFTComplexEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.ComplexEdgeFeature"
      Case esriFTAnnotation
        Exit Function
    End Select
  End If

  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit

  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef

  Dim dblIndex0 As Double
  Dim dblIndex1 As Double
  Dim dblIndex2 As Double
  dblIndex0 = 100
  dblIndex1 = 0
  dblIndex2 = 0

  If Not pExtent Is Nothing And lngOriginalRecCount > 0 Then
    If Not pExtent.IsEmpty Then
      Dim dblOptimalGridCells As Double
      dblOptimalGridCells = lngOriginalRecCount / 200
      Dim dblX As Double
      Dim dblY As Double
      dblX = pExtent.Width
      dblY = pExtent.Height
      dblIndex0 = Sqr(dblX * dblY / dblOptimalGridCells)
      If dblIndex0 > dblX Then
        dblIndex0 = dblX
      End If
      If dblIndex0 > dblY Then
        dblIndex0 = dblY
      End If

      dblIndex1 = 0
      dblIndex2 = 0
    End If
  End If

  If Not CheckSpRefDomain(pSpRef) Then
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
  End If

  With pGeomDefEdit
    .GeometryType = geomType
    If lngCategory = ENUM_FileGDB Then       ' FILE GDB
      .GridCount = 1
      .GridSize(0) = dblIndex0
    Else
      .GridCount = 1
      .GridSize(0) = dblIndex0
    End If
    .HasM = booHasM
    .HasZ = booHasZ
    Set .SpatialReference = pSpRef
  End With

  Set pField = New Field
  Set pFieldEdit = pField

  pFieldEdit.Name = "Shape"
  pFieldEdit.AliasName = "geometry"
  pFieldEdit.Type = esriFieldTypeGeometry
  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField

    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Object_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 11 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If

    Set pField = New Field
    Set pFieldEdit = pField
    pFieldEdit.Name = strUniqueName
    pFieldEdit.AliasName = "object identifier"
    pFieldEdit.Type = esriFieldTypeOID
    pFieldsEdit.AddField pField

  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If

  If (pCLSEXT Is Nothing) Or IsMissing(pCLSEXT) Then
    Set pCLSEXT = Nothing
  End If

  Dim strShapeFld As String
  Dim j As Integer
  For j = 0 To pFields.FieldCount - 1
    If pFields.Field(j).Type = esriFieldTypeGeometry Then
      strShapeFld = pFields.Field(j).Name
    End If
  Next

  Set CreateGDBFeatureClass2 = featWorkspace.CreateFeatureClass(Name, pFields, pCLSID, _
                             pCLSEXT, featType, strShapeFld, ConfigWord)

  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSpRefRes = Nothing
  Set pCheckField = Nothing

End Function

Private Function CheckSpRefDomain(pSpRef As ISpatialReference) As Boolean
  On Error GoTo ErrHandler

  Dim dXmax As Double
  Dim dYmax As Double
  Dim dXmin As Double
  Dim dYmin As Double

  pSpRef.GetDomain dXmin, dXmax, dYmin, dYmax
  CheckSpRefDomain = True

  Exit Function
ErrHandler:
  CheckSpRefDomain = False

End Function

Public Function CreateShapefileFeatureClass2(sPath As String, sName As String, pSpatialReference As ISpatialReference, _
    pGeomType As esriGeometryType, Optional pAddFields As esriSystem.IVariantArray, _
    Optional booForceUniqueIDField As Boolean = True, Optional booHasZ As Boolean = False, _
    Optional booHasM As Boolean = False) As IFeatureClass                                                  ' Don't include filename!

  If Right(sPath, 4) = ".shp" Then sPath = ReturnDir(sPath)
  If Right(sName, 4) = ".shp" Then sName = Left(sName, Len(sName) - 4)

  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    .GeometryType = pGeomType
    .HasM = booHasM
    .HasZ = booHasZ
    Set .SpatialReference = pSpatialReference
  End With

  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory

  If Not ExistFileDir(sPath) Then
    MsgBox "Unable to create Feature Class:" & vbCrLf & _
           sPath & " is not a valid workspace...", , "Failed to Create Feature Class:"
    Set CreateShapefileFeatureClass2 = Nothing
    Exit Function
  End If

  Set pFWS = pWorkspaceFactory.OpenFromFile(sPath, 0)

  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit

  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Shape"
  pFieldEdit.Type = esriFieldTypeGeometry

  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField

  If booForceUniqueIDField Then
    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Unique_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 10 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If

    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
        .Precision = 8
        .Name = strUniqueName
        .Type = esriFieldTypeInteger
    End With
    pFieldsEdit.AddField pField
  End If

  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".shp"
  Else
    strCheckString = sPath & "\" & sName & ".shp"
  End If

  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreateShapefileFeatureClass2 = Nothing
    Exit Function
  End If

  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(sName, pFields, Nothing, _
                                           Nothing, esriFTSimple, "Shape", "")

  Set CreateShapefileFeatureClass2 = pFeatClass

  GoTo ClearMemory

ClearMemory:
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pFWS = Nothing
  Set pWorkspaceFactory = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pCheckField = Nothing
  Set pFeatClass = Nothing
End Function

Public Function ReturnDBASEFieldName(strName As String, pFieldSet As Variant) As String

  Dim lngIndex As Long

  Dim strAcceptable As String
  strAcceptable = "abcdefghijklmnopqrstuvwxyz1234567890_"
  Dim strCharacters As String
  strCharacters = "abcdefghijklmnopqrstuvwxyz"

  Dim pField As iField

  Dim strNewName As String
  Dim strChar As String

  If strName = "" Then strName = "z_Field"
  If StrComp(strName, "date", vbTextCompare) = 0 Then strName = "z_Date"
  If StrComp(strName, "day", vbTextCompare) = 0 Then strName = "z_Day"
  If StrComp(strName, "month", vbTextCompare) = 0 Then strName = "z_Month"
  If StrComp(strName, "table", vbTextCompare) = 0 Then strName = "z_Table"
  If StrComp(strName, "text", vbTextCompare) = 0 Then strName = "z_Text"
  If StrComp(strName, "user", vbTextCompare) = 0 Then strName = "z_User"
  If StrComp(strName, "when", vbTextCompare) = 0 Then strName = "z_When"
  If StrComp(strName, "where", vbTextCompare) = 0 Then strName = "z_Where"
  If StrComp(strName, "year", vbTextCompare) = 0 Then strName = "z_Year"
  If StrComp(strName, "zone", vbTextCompare) = 0 Then strName = "z_Zone"
  If StrComp(strName, "Shape_Length", vbTextCompare) = 0 Then strName = "zShape_Len"
  If StrComp(strName, "Shape_Area", vbTextCompare) = 0 Then strName = "zShapeArea"
  If StrComp(strName, "OID", vbTextCompare) = 0 Then strName = "z_OID"
  If StrComp(strName, "FID", vbTextCompare) = 0 Then strName = "z_FID"
  If StrComp(strName, "Object_ID", vbTextCompare) = 0 Then strName = "z_ObjectID"
  If StrComp(strName, "Shape", vbTextCompare) = 0 Then strName = "z_Shape"
  If StrComp(strName, "ObjectID", vbTextCompare) = 0 Then strName = "z_ObjectID"

  strName = Left(strName, 10)

  If Not (InStr(1, strCharacters, Left(strName, 1), vbTextCompare) > 0) Then
    strName = Left("z" & strName, 10)
  End If

  strNewName = ""
  For lngIndex = 1 To Len(strName)
    strChar = Mid(strName, lngIndex, 1)
    If Not (InStr(1, strAcceptable, strChar, vbTextCompare) > 0) Then
      strNewName = strNewName & "_"
    Else
      strNewName = strNewName & strChar
    End If
  Next lngIndex
  strName = strNewName

  Dim strTestName As String

  Dim pFieldArray As esriSystem.IStringArray
  Set pFieldArray = New esriSystem.strArray
  If Not pFieldSet Is Nothing Then
    If TypeOf pFieldSet Is IFields Then

      Dim pFields As IFields
      Set pFields = pFieldSet
      If pFields.FieldCount > 0 Then
        For lngIndex = 0 To pFields.FieldCount - 1
          strTestName = pFields.Field(lngIndex).Name
          pFieldArray.Add strTestName
        Next lngIndex
      End If

    ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then

      Dim pVar As Variant
      Dim pVarArray As esriSystem.IVariantArray
      Set pVarArray = pFieldSet
      If pVarArray.Count > 0 Then
        For lngIndex = 0 To pVarArray.Count - 1
          If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
            Set pField = pVarArray.Element(lngIndex)
            strTestName = pField.Name
          Else
            strTestName = pVarArray.Element(lngIndex)
          End If
          pFieldArray.Add strTestName
        Next lngIndex
      End If

    ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then

      Set pFieldArray = pFieldSet

    End If
  End If

  If pFieldArray.Count > 0 Then
    Dim lngCounter As Long
    Dim strBaseName As String
    strBaseName = strName
    Dim booFoundConflict As Boolean
    booFoundConflict = True
    Do Until booFoundConflict = False
      booFoundConflict = False
      For lngIndex = 0 To pFieldArray.Count - 1
        strTestName = pFieldArray.Element(lngIndex)
        If StrComp(strName, strTestName, vbTextCompare) = 0 Then
          booFoundConflict = True
          Exit For
        End If
      Next lngIndex
      If booFoundConflict Then
        lngCounter = lngCounter + 1
        strName = Left(strBaseName, 10 - Len(CStr(lngCounter))) & CStr(lngCounter)
      End If
    Loop
  End If

  ReturnDBASEFieldName = strName

  GoTo ClearMemory

ClearMemory:
  Set pField = Nothing
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing

End Function

Public Function ReturnLayerByName(strName As String, pMap As IMap) As ILayer

  Set ReturnLayerByName = Nothing
  Dim pLayers As IEnumLayer
  Set pLayers = pMap.Layers(Nothing, True)
  Dim pLayer As ILayer
  pLayers.Reset
  Set pLayer = pLayers.Next
  Do Until pLayer Is Nothing
    If StrComp(pLayer.Name, strName, vbTextCompare) = 0 Then
      Set ReturnLayerByName = pLayer
      Exit Do
    End If
    Set pLayer = pLayers.Next
  Loop

  GoTo ClearMemory

ClearMemory:
  Set pLayers = Nothing
  Set pLayer = Nothing

End Function

Public Sub Graphic_MakeFromGeometry(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String, _
    Optional pSymbol As ISymbol, Optional booAddToLayout As Boolean = False)

  Dim pGContainer As IGraphicsContainer
  If booAddToLayout Then
    Set pGContainer = pMxDoc.PageLayout
  Else
    Set pGContainer = pMxDoc.FocusMap
  End If

  Dim pElement As IElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties

  Dim pMarkerElement As IMarkerElement
  Dim pFillElement As IFillShapeElement
  Dim pLineElement As ILineElement

  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone

  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType

  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
      Set pMarkerElement = pElement
    Case 2:
      Set pElement = New GroupElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
      Set pLineElement = pElement
    Case 4, 11:
      Set pElement = New PolygonElement
      Set pFillElement = pElement
    Case 5:
      Set pElement = New RectangleElement
      Set pFillElement = pElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select

  Dim pGroupElement As IGroupElement2
  Dim pSubElement As IElement
  Dim lngIndex As Long
  Dim pPtColl As IPointCollection
  Dim pPt As IPoint

  If pGeometryType = 2 Then
    Set pGroupElement = New GroupElement
    Set pSubElement = New MarkerElement
    Set pPtColl = pNewGeometry
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      If Not pSymbol Is Nothing Then
        Set pMarkerElement = pSubElement
        pMarkerElement.Symbol = pSymbol
      End If
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName

    pGContainer.AddElement pGroupElement, 0
  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName

    If Not pSymbol Is Nothing Then
      Select Case pGeometryType
        Case 1:
          pMarkerElement.Symbol = pSymbol
        Case 3, 6, 13, 14, 15, 16:
          pLineElement.Symbol = pSymbol
        Case 4, 11:
          pFillElement.Symbol = pSymbol
        Case 5:
          pFillElement.Symbol = pSymbol
      End Select
    End If

    pGContainer.AddElement pElement, 0
  End If

  pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  GoTo ClearMemory

ClearMemory:
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pMarkerElement = Nothing
  Set pFillElement = Nothing
  Set pLineElement = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing

End Sub

Public Function ReturnGraphicsByName(ByRef pMxDoc As IMxDocument, strName As String, _
      Optional AsElements As Boolean) As IArray

  Dim pGraphicsContainer As IGraphicsContainer

  Set pGraphicsContainer = pMxDoc.FocusMap

  pGraphicsContainer.Reset

  Dim pElement As IElement
  Dim pElementProperties As IElementProperties

  Set pElement = pGraphicsContainer.Next

  Dim pArray As IArray
  Set pArray = New esriSystem.Array
  Dim pGeometry As IGeometry
  Dim pClone As IClone

  While Not pElement Is Nothing
    Set pElementProperties = pElement
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If AsElements Then
        pArray.Add pElement
      Else
        Set pGeometry = pElement.Geometry
        Set pClone = pGeometry
        pArray.Add pClone.Clone     ' ONLY RETURN A COPY OF THE GEOMETRY; DON'T WANT TO MODIFY ACTUAL GRAPHIC HERE
      End If
    End If
    Set pElement = pGraphicsContainer.Next

  Wend
  Set ReturnGraphicsByName = pArray

  GoTo ClearMemory
ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pArray = Nothing
  Set pGeometry = Nothing
  Set pClone = Nothing
End Function

Public Sub DeleteGraphicsByName(ByRef pMxDoc As IMxDocument, strName As String, _
    Optional booDeleteFromLayout As Boolean = False)

  Dim pGraphicsContainer As IGraphicsContainer
  Dim pActiveView As IActiveView

  If booDeleteFromLayout Then
    Set pGraphicsContainer = pMxDoc.PageLayout
  Else
    Set pGraphicsContainer = pMxDoc.FocusMap
  End If
  Set pActiveView = pMxDoc.ActiveView
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  Dim pEnvelope As IEnvelope

  pGraphicsContainer.Reset

  Set pElement = pGraphicsContainer.Next

  Dim pDeleteArray As esriSystem.IVariantArray
  Set pDeleteArray = New esriSystem.varArray

  While Not pElement Is Nothing
    Set pElementProperties = pElement

    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If (pEnvelope Is Nothing) Then
        Set pEnvelope = pElement.Geometry.Envelope
      Else
        pEnvelope.Union pElement.Geometry.Envelope
      End If
      pDeleteArray.Add pElement
    End If
    Set pElement = pGraphicsContainer.Next

  Wend

  Dim lngIndex As Long
  If pDeleteArray.Count > 0 Then
    For lngIndex = 0 To pDeleteArray.Count - 1
      Set pElement = pDeleteArray.Element(lngIndex)
      pGraphicsContainer.DeleteElement pElement
    Next lngIndex
  End If

  If (Not pEnvelope Is Nothing) Then
    pActiveView.PartialRefresh esriViewGraphics + esriViewGraphicSelection + esriViewGeography, Nothing, pEnvelope
  End If

  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pActiveView = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pEnvelope = Nothing
  Set pDeleteArray = Nothing
End Sub

Public Function ReturnTimeElapsedFromMilliseconds(lngMilliseconds As Long) As String

  Dim theElapsedTime As Double
  Dim theNumDays As Double
  Dim theNumHours As Double
  Dim theNumMinutes As Double
  Dim theNumSeconds As Double

  theElapsedTime = lngMilliseconds / 1000

  theNumDays = Int(theElapsedTime / 86400)
  theNumHours = Int((theElapsedTime Mod 86400) / 3600)
  theNumMinutes = Int((theElapsedTime Mod 3600) / 60)
  theNumSeconds = theElapsedTime Mod 60

  Dim theDayString As String
  Dim theHourString As String
  Dim theMinString As String
  Dim theSecString As String

  If theNumDays = 1 Then
    theDayString = " day"
  Else
    theDayString = " days"
  End If

  If theNumHours = 1 Then
    theHourString = " hour"
  Else
    theHourString = " hours"
  End If

  If theNumMinutes = 1 Then
    theMinString = " minute"
  Else
    theMinString = " minutes"
  End If

  If theNumSeconds = 1 Then
    theSecString = " second..."
  Else
    theSecString = " seconds..."
  End If

  Dim theElapsedTimeString As String
  theElapsedTimeString = ""
  If theNumDays > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumDays & theDayString & ", " & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumHours > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumMinutes > 0 Then
    theElapsedTimeString = theElapsedTimeString & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  Else
    theElapsedTimeString = theElapsedTimeString & theNumSeconds & theSecString
  End If

  ReturnTimeElapsedFromMilliseconds = theElapsedTimeString

End Function

Public Function CheckCollectionForKey(colCollection As Collection, strKey As String) As Boolean
  On Error GoTo ErrHandler

  CheckCollectionForKey = True
  Dim lngVarType As Long
  lngVarType = VarType(colCollection.Item(strKey))

  Exit Function
ErrHandler:
  CheckCollectionForKey = False

End Function

Public Function CreateNestedFoldersByPath(ByVal completeDirectory As String) As Long

   Dim r As Long
   Dim SA As SECURITY_ATTRIBUTES
   Dim drivePart As String
   Dim newDirectory  As String
   Dim Item As String
   Dim sfolders() As String
   Dim pos As Long
   Dim x As Long

   If Right$(completeDirectory, 1) <> "\" Then
      completeDirectory = completeDirectory & "\"
   End If

   pos = InStr(completeDirectory, ":")

   If pos Then
      drivePart = GetPart(completeDirectory, "\")
   Else: drivePart = ""
   End If

   Do Until completeDirectory = ""

     Item = GetPart(completeDirectory, "\")

     ReDim Preserve sfolders(0 To x) As String

     If x = 0 Then Item = drivePart & Item
     sfolders(x) = Item

     x = x + 1

   Loop

   x = -1

   Do

      x = x + 1

      newDirectory = newDirectory & sfolders(x)

      SA.nLength = LenB(SA)

      Call CreateDirectory(newDirectory, SA)

   Loop Until x = UBound(sfolders)

   CreateNestedFoldersByPath = x + 1

  GoTo ClearMemory
ClearMemory:
  Erase sfolders

End Function

Public Function GetPart(startStrg As String, delimiter As String) As String

  Dim C As Integer
  Dim Item As String

  C = 1

  Do

    If Mid$(startStrg, C, 1) = delimiter Then

      Item = Mid$(startStrg, 1, C)
      startStrg = Mid$(startStrg, C + 1, Len(startStrg))
      GetPart = Item
      Exit Function

    End If

    C = C + 1

  Loop

End Function

Public Function ReturnTimeStamp() As String

  ReturnTimeStamp = CStr(Format(Now, "yyyymmdd_HhNnSs"))

End Function

Public Function ReturnQuerySpecialCharacters(pDataset As IDataset, Optional strPrefix As String, _
    Optional strSuffix As String, Optional strWildcardSingleMatch As String, _
    Optional strWildlcardManyMatch As String, Optional strSQLEscapePrefix As String, _
    Optional strSQLEscapeSuffix As String) As Boolean
  On Error GoTo ErrHandler

  ReturnQuerySpecialCharacters = False

  Dim pWS As IWorkspace
  If TypeOf pDataset Is IWorkspace Then
    Set pWS = pDataset
  Else
    Set pWS = pDataset.Workspace
  End If

  Dim pSQLSyntax As ISQLSyntax
  Set pSQLSyntax = pWS
  strPrefix = pSQLSyntax.GetSpecialCharacter(esriSQL_DelimitedIdentifierPrefix)
  strSuffix = pSQLSyntax.GetSpecialCharacter(esriSQL_DelimitedIdentifierSuffix)
  strWildcardSingleMatch = pSQLSyntax.GetSpecialCharacter(esriSQL_WildcardSingleMatch)
  strWildlcardManyMatch = pSQLSyntax.GetSpecialCharacter(esriSQL_WildcardManyMatch)
  strSQLEscapePrefix = pSQLSyntax.GetSpecialCharacter(esriSQL_EscapeKeyPrefix)
  strSQLEscapeSuffix = pSQLSyntax.GetSpecialCharacter(esriSQL_EscapeKeySuffix)

  ReturnQuerySpecialCharacters = True

  GoTo ClearMemory
  Exit Function
ErrHandler:

  ReturnQuerySpecialCharacters = False

ClearMemory:
  Set pWS = Nothing
  Set pSQLSyntax = Nothing

End Function

Public Function ReturnTitleCase(ByVal strString As String) As String

  Do Until InStr(1, strString, "  ") = 0
    strString = Replace(strString, "  ", " ")
  Loop

  Dim strSplit() As String
  strSplit = Split(strString, " ")
  Dim lngIndex As Long
  Dim strWord As String
  For lngIndex = 0 To UBound(strSplit)
    strWord = Trim(strSplit(lngIndex))
    If Len(strWord) = 1 Then
      ReturnTitleCase = ReturnTitleCase & UCase(strWord) & " "
    Else
      ReturnTitleCase = ReturnTitleCase & UCase(Left(strWord, 1)) & _
        LCase(Right(strWord, Len(strWord) - 1)) & " "
    End If
  Next lngIndex
  ReturnTitleCase = Left(ReturnTitleCase, Len(ReturnTitleCase) - 1)

  GoTo ClearMemory
ClearMemory:
  Erase strSplit
End Function

Public Function ReadTextFile(strFilename As String) As String

  If Dir(strFilename) <> "" Then

    Dim lngFileNumber As Long
    lngFileNumber = FreeFile(0)

    Dim strFileText As String

    Open strFilename For Binary As #lngFileNumber

    strFileText = Space$(LOF(lngFileNumber))
    Get #lngFileNumber, , strFileText

    Close #lngFileNumber

    ReadTextFile = strFileText
  Else
    ReadTextFile = ""
  End If

End Function

Public Function WriteTextFile(strFilename As String, strText As String, Optional booForceOverwrite As Boolean = False, _
    Optional booAppend As Boolean = False) As Boolean

  Dim lngFileNumber As Long

  If Dir(strFilename) = "" Or booForceOverwrite Then
    lngFileNumber = FreeFile(0)

    Open strFilename For Output As #lngFileNumber

    Print #lngFileNumber, strText
    Close #lngFileNumber
    WriteTextFile = True

  ElseIf booAppend Then
    lngFileNumber = FreeFile(0)

    Open strFilename For Append As #lngFileNumber

    Print #lngFileNumber, strText
    Close #lngFileNumber
    WriteTextFile = True

  Else
    Dim lngVBResult As VbMsgBoxResult

    lngVBResult = MsgBox("File Already Exists!" & vbCrLf & vbCrLf & "  --> " & strFilename & vbCrLf & vbCrLf & _
        "Click 'OK' to overwrite the file, or 'CANCEL' to quit...", vbOKCancel, "File Exists:")
    If lngVBResult = vbOK Then
      Kill strFilename
      If Dir(strFilename) <> "" Then
        MsgBox "Unable to delete " & strFilename & vbCrLf & vbCrLf & _
          "It may be open in another application.  Please delete this file manually or save the text to a new filename.", , "Unable to Delete File:"
        WriteTextFile = False
      Else

        lngFileNumber = FreeFile(0)

        Open strFilename For Output As #lngFileNumber

        Print #lngFileNumber, strText
        Close #lngFileNumber
        WriteTextFile = True
      End If
    Else
      WriteTextFile = False
    End If
  End If

End Function

Public Function ReturnFilesFromNestedFolders2(ByVal strDir As String, strAnyTextInName As String) As esriSystem.IStringArray

  Set ReturnFilesFromNestedFolders2 = New esriSystem.strArray

  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"

  Dim strOriginalDir As String
  strOriginalDir = strDir

  Dim booFoundSubFolders As Boolean

  Dim pPathArray As esriSystem.IStringArray
  Set pPathArray = New esriSystem.strArray

  Dim pFinalArray As esriSystem.IStringArray
  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection

  pFinalArray.Add strDir
  pCheckColl.Add True, strDir

  Dim strDirName As String

  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     If strDirName <> "." And strDirName <> ".." Then
        If IsFolder_FalseIfCrash((strDir & strDirName)) Then
           pPathArray.Add strDir & strDirName & "\"
           pFinalArray.Add strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop

  booFoundSubFolders = pPathArray.Count > 0

  Dim strSubFolder As String

  Dim booFoundSubHere As Boolean
  Dim pSubArray As esriSystem.IStringArray

  Dim lngIndex As Long

  Do While booFoundSubFolders
    booFoundSubFolders = False
    Set pSubArray = New esriSystem.strArray

    For lngIndex = 0 To pPathArray.Count - 1
      strSubFolder = pPathArray.Element(lngIndex)

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         If strDirName <> "." And strDirName <> ".." Then
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
               pSubArray.Add strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 pFinalArray.Add strSubFolder & strDirName & "\"
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop

      If Not booFoundSubHere Then
        pSubArray.Add strSubFolder
      End If

    Next lngIndex

    If booFoundSubFolders Then
      Set pPathArray = pSubArray
    End If

  Loop

  Dim strFolders() As String
  ReDim strFolders(pFinalArray.Count - 1)

  For lngIndex = 0 To pFinalArray.Count - 1
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
  Next lngIndex

  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1

  Dim lngCounter As Long
  lngCounter = 0

  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pFilenames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pFilenames = New esriSystem.strArray
  Dim strDirAndFile As String

  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)
    strDirName = Dir(strDir, vbNormal)   ' Retrieve the first entry.
    lngCounter = lngCounter + 1

    Do While strDirName <> ""   ' Start the loop.
       If strDirName <> "." And strDirName <> ".." Then
          strDirAndFile = strDir & strDirName

          If IsNormal_FalseIfCrash(strDirAndFile) Then
            If InStr(1, strDirName, strAnyTextInName, vbTextCompare) > 0 Then
              pFilenames.Add strDirAndFile
            End If
          End If
       End If
       strDirName = Dir   ' Get next entry.
    Loop
  Next lngIndex

  Set ReturnFilesFromNestedFolders2 = pFilenames

  GoTo ClearMemory

ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing

End Function

Public Function Get_Element_Or_Envelope_Point(pElementOrEnvelope As IUnknown, lngAnchorPoint As Jen_ElementEnvPoint, _
  Optional dblXPercent As Double = 0.5, Optional dblYPercent As Double = 0.5, Optional pActiveView As IActiveView) As IPoint

  Dim theOutput As IPoint
  Set theOutput = Nothing
  Dim pEnv As IEnvelope

  If TypeOf pElementOrEnvelope Is IElement Then
    Dim pTemp As IElement
    Set pTemp = pElementOrEnvelope
    If pActiveView Is Nothing Then
      Set pEnv = pTemp.Geometry.Envelope
    Else
      Set pEnv = New Envelope
      pTemp.QueryBounds pActiveView.ScreenDisplay, pEnv
    End If
  Else
    Set pEnv = pElementOrEnvelope
  End If

  Select Case lngAnchorPoint
    Case ENUM_Upper_Left  ' UPPER LEFT CORNER
        Set theOutput = pEnv.UpperLeft
    Case ENUM_Upper_Right  ' UPPER RIGHT CORNER
        Set theOutput = pEnv.UpperRight
    Case ENUM_Lower_Left ' LOWER LEFT CORNER
        Set theOutput = pEnv.LowerLeft
    Case ENUM_Lower_Right  ' LOWER RIGHT CORNER
        Set theOutput = pEnv.LowerRight
    Case ENUM_Upper_Center  ' VERTICAL TOP, HORIZONTAL CENTER
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, pEnv.YMax
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Lower_Center  ' VERTICAL BOTTOM, HORIZONTAL CENTER
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Center_Left  ' VERTICAL CENTER, HORIZONTAL LEFT
        Set theOutput = New Point
        theOutput.PutCoords pEnv.XMin, ((pEnv.YMax - pEnv.YMin) / 2) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Center_Right  ' VERTICAL CENTER, HORIZONTAL RIGHT
        Set theOutput = New Point
        theOutput.PutCoords pEnv.XMax, ((pEnv.YMax - pEnv.YMin) / 2) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Center_Center  ' VERTICAL CENTER, HORIZONTAL CENTER
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, ((pEnv.YMax - pEnv.YMin) / 2) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_By_Percentages
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) * dblXPercent) + pEnv.XMin, ((pEnv.YMax - pEnv.YMin) * dblYPercent) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case Else
        MsgBox "position not supported: " & CLng(lngAnchorPoint)
  End Select

  Set Get_Element_Or_Envelope_Point = theOutput

  GoTo ClearMemory

ClearMemory:
  Set theOutput = Nothing
  Set pEnv = Nothing
  Set pTemp = Nothing

End Function ' Get_Element_Or_Envelope_Point

Public Function ReturnGraphicsByNameFromLayout(pMxDoc As IMxDocument, strName As String, _
      Optional AsElements As Boolean) As IArray

  Dim pGraphicsContainer As IGraphicsContainer

  Set pGraphicsContainer = pMxDoc.PageLayout

  pGraphicsContainer.Reset

  Dim pElement As IElement
  Dim pElementProperties As IElementProperties

  Set pElement = pGraphicsContainer.Next

  Dim pArray As IArray
  Set pArray = New esriSystem.Array
  Dim pGeometry As IGeometry
  Dim pClone As IClone

  While Not pElement Is Nothing
    Set pElementProperties = pElement
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If AsElements Then
        pArray.Add pElement
      Else
        Set pGeometry = pElement.Geometry
        Set pClone = pGeometry
        pArray.Add pClone.Clone     ' ONLY RETURN A COPY OF THE GEOMETRY; DON'T WANT TO MODIFY ACTUAL GRAPHIC HERE
      End If
    End If
    Set pElement = pGraphicsContainer.Next

  Wend
  Set ReturnGraphicsByNameFromLayout = pArray

  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pArray = Nothing
  Set pGeometry = Nothing
  Set pClone = Nothing

End Function

Public Function CreateSpatialReferenceNAD83() As ISpatialReference

  Dim pNAD83 As IGeographicCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pNAD83 = pSpatRefFact.CreateGeographicCoordinateSystem(esriSRGeoCS_NAD1983)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pNAD83
  pSpRefRes.ConstructFromHorizon

  Set CreateSpatialReferenceNAD83 = pNAD83

  Set pNAD83 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

  GoTo ClearMemory
ClearMemory:
  Set pNAD83 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function

Public Function ReturnFoldersFromNestedFolders(ByVal strDir As String, strPartialText As String) As esriSystem.IStringArray

  Set ReturnFoldersFromNestedFolders = New esriSystem.strArray

  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"

  Dim strOriginalDir As String
  strOriginalDir = strDir

  Dim booFoundSubFolders As Boolean

  Dim pPathArray As esriSystem.IStringArray
  Set pPathArray = New esriSystem.strArray

  Dim pFinalArray As esriSystem.IStringArray
  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection

  pFinalArray.Add strDir
  pCheckColl.Add True, strDir

  Dim strDirName As String

  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     If strDirName <> "." And strDirName <> ".." Then
        If IsFolder_FalseIfCrash(strDir & strDirName) Then
           pPathArray.Add strDir & strDirName & "\"
           pFinalArray.Add strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop

  booFoundSubFolders = pPathArray.Count > 0

  Dim strSubFolder As String

  Dim booFoundSubHere As Boolean
  Dim pSubArray As esriSystem.IStringArray

  Dim lngIndex As Long

  Do While booFoundSubFolders
    booFoundSubFolders = False
    Set pSubArray = New esriSystem.strArray

    For lngIndex = 0 To pPathArray.Count - 1
      strSubFolder = pPathArray.Element(lngIndex)

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         If strDirName <> "." And strDirName <> ".." Then
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
               pSubArray.Add strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 pFinalArray.Add strSubFolder & strDirName & "\"
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop

      If Not booFoundSubHere Then
        pSubArray.Add strSubFolder
      End If

    Next lngIndex

    If booFoundSubFolders Then
      Set pPathArray = pSubArray
    End If

  Loop

  Dim strFolders() As String
  ReDim strFolders(pFinalArray.Count - 1)

  For lngIndex = 0 To pFinalArray.Count - 1
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
  Next lngIndex

  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1

  Dim lngCounter As Long
  lngCounter = 0

  Debug.Print

  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pFilenames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pFilenames = New esriSystem.strArray

  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)

    If InStr(1, strDir, strPartialText, vbTextCompare) > 0 Then
      pFilenames.Add strDir
    End If
  Next lngIndex

  Set ReturnFoldersFromNestedFolders = pFilenames

  GoTo ClearMemory

ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing

End Function

Public Function SpacesInFrontOfText(strText As String, lngTotalLength As Long) As String

  Dim lngCurrentLength As Long
  lngCurrentLength = Len(strText)

  If lngCurrentLength >= lngTotalLength Then
    SpacesInFrontOfText = strText
  Else
    SpacesInFrontOfText = Space(lngTotalLength - lngCurrentLength) & strText
  End If

End Function

Public Function CreateInMemoryFeatureClass3(pGeometryArray As esriSystem.IArray, _
    Optional pValueArray As esriSystem.IVariantArray, Optional pTemplateFields As esriSystem.IVariantArray, _
    Optional pApp As IApplication, Optional strStatusMessage As String = "", _
    Optional lngFlushCount As Long = 500) As IFeatureClass

    Dim pSBar As IStatusBar
    Dim pPro As IStepProgressor
    Dim dateRunningTime As Date
    Dim strHeader As String
    Dim lngTotalCount As Long

    If Not pApp Is Nothing Then
      Set pSBar = pApp.StatusBar
      Set pPro = pSBar.ProgressBar
      dateRunningTime = Now
      strHeader = strStatusMessage
      lngTotalCount = pGeometryArray.Count
      pSBar.ShowProgressBar strStatusMessage, 0, lngTotalCount, 1, True
      pPro.position = 0
    End If

    Dim pGeom As IGeometry
    Set pGeom = pGeometryArray.Element(0)

    Dim pSpRef As ISpatialReference
    Set pSpRef = pGeom.SpatialReference
    Dim pClone As IClone

    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New InMemoryWorkspaceFactory

    Dim pName As IName
    Set pName = pWSF.Create("", "inmemory", Nothing, 0)
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pName.Open

    Dim pFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit

    Set pFields = New Fields
    Set pFieldsEdit = pFields

    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef

    With pGeomDefEdit
      .GeometryType = pGeom.GeometryType
      .GridCount = 1
      .GridSize(0) = 0
      If pGeom.GeometryType = esriGeometryPoint Then
        .AvgNumPoints = 1
      Else
        .AvgNumPoints = 5
      End If
      .HasM = False
      .HasZ = False
      Set .SpatialReference = pGeom.SpatialReference
    End With

    Set pField = New Field
    Set pFieldEdit = pField

    pFieldEdit.Name = "Shape"
    pFieldEdit.AliasName = "geometry"
    pFieldEdit.Type = esriFieldTypeGeometry
    Set pFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField

    Dim booAddAttribute As Boolean
    booAddAttribute = Not pValueArray Is Nothing And Not pTemplateFields Is Nothing
    Dim varVal As Variant

    Dim lngIndex As Long
    Dim pSubArray As esriSystem.IVariantArray
    Dim pTemplateField As iField

    Dim pTempField As iField
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        Set pClone = pTemplateField
        Set pTempField = pClone.Clone
        pFieldsEdit.AddField pTempField
      Next lngIndex
      ReDim lngIDIndex(pTemplateFields.Count - 1)
    Else
      Set pField = New Field
      Set pFieldEdit = pField
      With pFieldEdit
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
      End With
      pFieldsEdit.AddField pField
      ReDim lngIDIndex(0)
    End If

    Dim pCLSID As UID
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Feature"

    Dim pInMemFC As IFeatureClass
    Set pInMemFC = pFWS.CreateFeatureClass("In_Memory", pFields, _
                             pCLSID, Nothing, esriFTSimple, _
                             "Shape", "")

    Dim lngAttIndex As Long
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        lngIDIndex(lngIndex) = pInMemFC.FindField(pTemplateField.Name)
      Next lngIndex
    Else
      lngIDIndex(0) = pInMemFC.FindField("Unique_ID")
    End If

    Dim pOutFeatBuffer As IFeatureBuffer
    Set pOutFeatBuffer = pInMemFC.CreateFeatureBuffer
    Dim pOutFCursor As IFeatureCursor
    Set pOutFCursor = pInMemFC.Insert(True)
    Dim lngIndex2 As Long

    For lngIndex = 0 To pGeometryArray.Count - 1
      Set pClone = pGeometryArray.Element(lngIndex)
      Set pGeom = pClone.Clone

      Set pOutFeatBuffer.Shape = pGeom

      If booAddAttribute Then
        Set pSubArray = pValueArray.Element(lngIndex)
        For lngIndex2 = 0 To pSubArray.Count - 1
          varVal = pSubArray.Element(lngIndex2)
          pOutFeatBuffer.Value(lngIDIndex(lngIndex2)) = varVal

        Next lngIndex2
      Else
        pOutFeatBuffer.Value(lngIDIndex(0)) = lngIndex + 1
      End If
      pOutFCursor.InsertFeature pOutFeatBuffer

      If lngIndex Mod lngFlushCount = 0 Then pOutFCursor.Flush
      If Not pApp Is Nothing Then
        pPro.Step
      End If
    Next lngIndex
    pOutFCursor.Flush

    If Not pApp Is Nothing Then
      pPro.position = 1
      pSBar.HideProgressBar
    End If
    Set CreateInMemoryFeatureClass3 = pInMemFC

  GoTo ClearMemory
ClearMemory:
  Set pSBar = Nothing
  Set pPro = Nothing
  Set pGeom = Nothing
  Set pSpRef = Nothing
  Set pClone = Nothing
  Set pSpRefRes = Nothing
  Set pWSF = Nothing
  Set pName = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  varVal = Null
  Set pSubArray = Nothing
  Set pTemplateField = Nothing
  Erase lngIDIndex
  Set pTempField = Nothing
  Set pCLSID = Nothing
  Set pInMemFC = Nothing
  Set pOutFeatBuffer = Nothing
  Set pOutFCursor = Nothing

End Function

Public Function IsFolder_FalseIfCrash(strPath As String) As Boolean
  On Error GoTo ErrHandle

  IsFolder_FalseIfCrash = (GetAttr(strPath) And vbDirectory) = vbDirectory

  Exit Function
ErrHandle:
  IsFolder_FalseIfCrash = False

End Function

Public Function IsNormal_FalseIfCrash(strPath As String) As Boolean
  On Error GoTo ErrHandle

  IsNormal_FalseIfCrash = (GetAttr(strPath) And vbNormal) = vbNormal

  Exit Function
ErrHandle:
  IsNormal_FalseIfCrash = False

End Function

Public Function CreateGeneralProjectedSpatialReference(lngFactoryID As Long) As ISpatialReference

  Dim pGeneralGeo As IProjectedCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pGeneralGeo = pSpatRefFact.CreateProjectedCoordinateSystem(lngFactoryID)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pGeneralGeo
  pSpRefRes.ConstructFromHorizon

  Set CreateGeneralProjectedSpatialReference = pGeneralGeo

  Set pGeneralGeo = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

  GoTo ClearMemory
ClearMemory:
  Set pGeneralGeo = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function

Public Function Graphic_ReturnElementFromGeometry3(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String, _
    Optional pSymbol As ISymbol, Optional booElementIsForLayout As Boolean = True, _
    Optional booAlsoAddElementToMapDoc As Boolean = False) As IElement

  Dim pMxDocument As esriArcMapUI.IMxDocument
  Dim pActiveView As esriCarto.IActiveView

  Dim pGContainer As IGraphicsContainer
  If booElementIsForLayout Then
    Set pGContainer = pMxDoc.PageLayout
  Else
    Set pGContainer = pMxDoc.FocusMap
  End If

  Dim pElement As IElement
  Dim pPolygonElement As IPolygonElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties

  Dim pMarkerElement As IMarkerElement
  Dim pFillElement As IFillShapeElement
  Dim pLineElement As ILineElement

  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone

  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType

  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
      Set pMarkerElement = pElement
    Case 2:
      Set pElement = New GroupElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
      Set pLineElement = pElement
    Case 4, 11:
      Set pElement = New PolygonElement
      Set pFillElement = pElement
    Case 5:
      Set pElement = New RectangleElement
      Set pFillElement = pElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select

  If pGeometryType = 2 Then
    Dim pGroupElement As IGroupElement2
    Set pGroupElement = New GroupElement
    Dim pSubElement As IElement
    Set pSubElement = New MarkerElement
    Dim lngIndex As Long
    Dim pPtColl As IPointCollection
    Set pPtColl = pNewGeometry
    Dim pPt As IPoint
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName

    If booAlsoAddElementToMapDoc Then pGContainer.AddElement pGroupElement, 0

    Set Graphic_ReturnElementFromGeometry3 = pGroupElement
  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName

    If Not pSymbol Is Nothing Then
      Select Case pGeometryType
        Case 1:
          pMarkerElement.Symbol = pSymbol
        Case 3, 6, 13, 14, 15, 16:
          pLineElement.Symbol = pSymbol
        Case 4, 11:
          pFillElement.Symbol = pSymbol
        Case 5:
          pFillElement.Symbol = pSymbol
      End Select
    End If

    If booAlsoAddElementToMapDoc Then pGContainer.AddElement pElement, 0

    Set Graphic_ReturnElementFromGeometry3 = pElement

  End If

  GoTo ClearMemory
ClearMemory:
  Set pMxDocument = Nothing
  Set pActiveView = Nothing
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pPolygonElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pMarkerElement = Nothing
  Set pFillElement = Nothing
  Set pLineElement = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing

End Function

Public Function CreateFieldAttributeIndex(pTable As ITable, strFieldName As String, _
    Optional strFailReason As String) As Boolean

  Dim pDataset As IDataset
  Dim strDatasetName As String
  Set pDataset = pTable
  strDatasetName = pDataset.BrowseName

  CreateFieldAttributeIndex = True
  strFailReason = ""

  Dim pIndexes As IIndexes
  Set pIndexes = pTable.Indexes

  Dim pEnumIndex As IEnumIndex
  Set pEnumIndex = pIndexes.FindIndexesByFieldName(strFieldName)

  Dim pFields As IFields
  Dim pField As iField
  Dim pFieldsEdit As IFieldsEdit
  Dim pIndex As IIndex
  Dim pIndexEdit As IIndexEdit

  Set pIndex = pEnumIndex.Next

  If pIndex Is Nothing Then

    If pTable.FindField(strFieldName) = -1 Then
      CreateFieldAttributeIndex = False
      strFailReason = "No Field with name '" & strFieldName & "' [" & strDatasetName & "]"
      Debug.Print "  --> " & strFailReason
    Else
      Set pField = pTable.Fields.Field(pTable.FindField(strFieldName))
      Set pFields = New Fields
      Set pFieldsEdit = pFields
      pFieldsEdit.FieldCount = 1
      Set pFieldsEdit.Field(0) = pField

      Set pIndex = New Index
      Set pIndexEdit = pIndex

      Set pIndexEdit.Fields = pFields
      pIndexEdit.IsAscending = False
      pIndexEdit.IsUnique = False
      pIndexEdit.Name = strFieldName

      pTable.AddIndex pIndex

      Debug.Print "  --> Built Index for '" & strFieldName & "' [" & strDatasetName & "]"

    End If
  Else
    CreateFieldAttributeIndex = False
    strFailReason = "Field '" & strFieldName & "' already indexed"
  End If

ClearMemory:
  Set pIndexes = Nothing
  Set pEnumIndex = Nothing
  Set pFields = Nothing
  Set pField = Nothing
  Set pFieldsEdit = Nothing
  Set pIndex = Nothing
  Set pIndexEdit = Nothing

End Function

Function IsDimmed(Arr As Variant) As Boolean
  On Error GoTo ReturnFalse
  IsDimmed = UBound(Arr) >= LBound(Arr)
  Exit Function
ReturnFalse:
  IsDimmed = False
End Function

Public Function ReturnValidFGDBFieldName2(strName As String, pFieldSet As IUnknown) As String

  Dim strAlphaNumeric As String
  strAlphaNumeric = "0123456789_abcdefghijklmnopqrstuvwxyz"

  Dim strChar As String
  Dim lngIndex As Long
  Dim strNewName As String

  For lngIndex = 1 To Len(strName)
    strChar = Mid(strName, lngIndex, 1)
    If InStr(1, strAlphaNumeric, strChar, vbTextCompare) > 0 Then
      strNewName = strNewName & strChar
    Else
      strNewName = strNewName & "_"
    End If
  Next lngIndex
  strChar = Left(strNewName, 1)
  If IsNumeric(strChar) Then
    strNewName = "z" & strNewName
  End If

  Dim strReserved() As String
  ReDim strReserved(30)
  strReserved(0) = "Add"
  strReserved(1) = "ALTER"
  strReserved(2) = "AND"
  strReserved(3) = "AS"
  strReserved(4) = "Asc"
  strReserved(5) = "BETWEEN"
  strReserved(6) = "BY"
  strReserved(7) = "Column"
  strReserved(8) = "Create"
  strReserved(9) = "Date"
  strReserved(10) = "DELETE"
  strReserved(11) = "DESC"
  strReserved(12) = "DROP"
  strReserved(13) = "Exists"
  strReserved(14) = "FOR"
  strReserved(15) = "From"
  strReserved(16) = "IN"
  strReserved(17) = "Insert"
  strReserved(18) = "INTO"
  strReserved(19) = "IS"
  strReserved(20) = "LIKE"
  strReserved(21) = "NOT"
  strReserved(22) = "Null"
  strReserved(23) = "OR"
  strReserved(24) = "Order"
  strReserved(25) = "SELECT"
  strReserved(26) = "SET"
  strReserved(27) = "Table"
  strReserved(28) = "Update"
  strReserved(29) = "Values"
  strReserved(30) = "WHERE"

  For lngIndex = 0 To UBound(strReserved)
    If StrComp(strNewName, strReserved(lngIndex), vbTextCompare) = 0 Then
      strNewName = strNewName & "_"
    End If
  Next lngIndex

  Dim lngMaxLength As Long
  lngMaxLength = 64
  strNewName = Left(strNewName, lngMaxLength)

  Dim strTestName As String

  Dim pFieldArray As esriSystem.IStringArray
  Set pFieldArray = New esriSystem.strArray

  If Not pFieldSet Is Nothing Then

    If TypeOf pFieldSet Is IFields Then   ' <--------------------

      Dim pFields As IFields
      Set pFields = pFieldSet
      If pFields.FieldCount > 0 Then
        For lngIndex = 0 To pFields.FieldCount - 1
          strTestName = pFields.Field(lngIndex).Name
          pFieldArray.Add strTestName
        Next lngIndex
      End If

    ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then

      Dim pVar As Variant
      Dim pVarArray As esriSystem.IVariantArray
      Dim pField As iField
      Set pVarArray = pFieldSet
      If pVarArray.Count > 0 Then
        For lngIndex = 0 To pVarArray.Count - 1
          If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
            Set pField = pVarArray.Element(lngIndex)
            strTestName = pField.Name
          Else
            strTestName = pVarArray.Element(lngIndex)
          End If
          pFieldArray.Add strTestName
        Next lngIndex
      End If

    ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then

      Set pFieldArray = pFieldSet

    End If
  End If

  If pFieldArray.Count > 0 Then
    Dim lngCounter As Long
    Dim strBaseName As String
    strBaseName = strNewName
    Dim booFoundConflict As Boolean
    booFoundConflict = True
    Do Until booFoundConflict = False
      booFoundConflict = False
      For lngIndex = 0 To pFieldArray.Count - 1
        strTestName = pFieldArray.Element(lngIndex)
        If StrComp(strNewName, strTestName, vbTextCompare) = 0 Then
          booFoundConflict = True
          Exit For
        End If
      Next lngIndex
      If booFoundConflict Then
        lngCounter = lngCounter + 1
        strNewName = Left(strBaseName, lngMaxLength - Len(CStr(lngCounter))) & CStr(lngCounter)
      End If
    Loop
  End If

  ReturnValidFGDBFieldName2 = strNewName

  GoTo ClearMemory
ClearMemory:
  Erase strReserved
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing
  Set pField = Nothing

End Function

Public Function ReturnAcceptableFieldName2(strOrigName As String, pFieldSet As IUnknown, Optional booRestrictToDBase As Boolean = False, _
    Optional booRestrictToPersonalGDB As Boolean = False, Optional booRestrictToCoverage As Boolean = False, _
    Optional booRestrictToFGDB As Boolean = False) As String

    Dim strName As String
    strName = strOrigName

  If booRestrictToDBase Then
    ReturnAcceptableFieldName2 = ReturnDBASEFieldName(strName, pFieldSet)
  ElseIf booRestrictToFGDB Then
    ReturnAcceptableFieldName2 = ReturnValidFGDBFieldName2(strName, pFieldSet)

  Else
    Dim lngIndex As Long

    Dim strAcceptable As String
    strAcceptable = "abcdefghijklmnopqrstuvwxyz1234567890_"
    Dim strCharacters As String
    strCharacters = "abcdefghijklmnopqrstuvwxyz"

    Dim lngMaxLength As Long
    If booRestrictToCoverage Then
      lngMaxLength = 16
    ElseIf booRestrictToPersonalGDB Then
      lngMaxLength = 52
    ElseIf booRestrictToFGDB Then
      lngMaxLength = 64
    Else
      lngMaxLength = 64
    End If

    Dim pField As iField

    Dim strNewName As String
    Dim strChar As String

    If strName = "" Then strName = "z_Field"
    If StrComp(strName, "date", vbTextCompare) = 0 Then strName = "z_Date"
    If StrComp(strName, "day", vbTextCompare) = 0 Then strName = "z_Day"
    If StrComp(strName, "month", vbTextCompare) = 0 Then strName = "z_Month"
    If StrComp(strName, "table", vbTextCompare) = 0 Then strName = "z_Table"
    If StrComp(strName, "text", vbTextCompare) = 0 Then strName = "z_Text"
    If StrComp(strName, "user", vbTextCompare) = 0 Then strName = "z_User"
    If StrComp(strName, "when", vbTextCompare) = 0 Then strName = "z_When"
    If StrComp(strName, "where", vbTextCompare) = 0 Then strName = "z_Where"
    If StrComp(strName, "year", vbTextCompare) = 0 Then strName = "z_Year"
    If StrComp(strName, "zone", vbTextCompare) = 0 Then strName = "z_Zone"
    If StrComp(strName, "Shape_Length", vbTextCompare) = 0 Then strName = "z_Shape_Length"
    If StrComp(strName, "Shape_Area", vbTextCompare) = 0 Then strName = "z_Shape_Area"
    If StrComp(strName, "OID", vbTextCompare) = 0 Then strName = "z_OID"
    If StrComp(strName, "FID", vbTextCompare) = 0 Then strName = "z_FID"
    If StrComp(strName, "Object_ID", vbTextCompare) = 0 Then strName = "z_Object_ID"
    If StrComp(strName, "Shape", vbTextCompare) = 0 Then strName = "z_Shape"
    If StrComp(strName, "ObjectID", vbTextCompare) = 0 Then strName = "z_ObjectID"

    strName = Left(strName, lngMaxLength)

    If Not (InStr(1, strCharacters, Left(strName, 1), vbTextCompare) > 0) Then
      strName = Left("z" & strName, lngMaxLength)
    End If

    strNewName = ""
    For lngIndex = 1 To Len(strName)
      strChar = Mid(strName, lngIndex, 1)
      If Not (InStr(1, strAcceptable, strChar, vbTextCompare) > 0) Then
        strNewName = strNewName & "_"
      Else
        strNewName = strNewName & strChar
      End If
    Next lngIndex
    strName = strNewName

    Dim strTestName As String

    Dim pFieldArray As esriSystem.IStringArray
    Set pFieldArray = New esriSystem.strArray

    If Not pFieldSet Is Nothing Then

      If TypeOf pFieldSet Is IFields Then   ' <--------------------

        Dim pFields As IFields
        Set pFields = pFieldSet
        If pFields.FieldCount > 0 Then
          For lngIndex = 0 To pFields.FieldCount - 1
            strTestName = pFields.Field(lngIndex).Name
            pFieldArray.Add strTestName
          Next lngIndex
        End If

      ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then

        Dim pVar As Variant
        Dim pVarArray As esriSystem.IVariantArray
        Set pVarArray = pFieldSet
        If pVarArray.Count > 0 Then
          For lngIndex = 0 To pVarArray.Count - 1
            If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
              Set pField = pVarArray.Element(lngIndex)
              strTestName = pField.Name
            Else
              strTestName = pVarArray.Element(lngIndex)
            End If
            pFieldArray.Add strTestName
          Next lngIndex
        End If

      ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then

        Set pFieldArray = pFieldSet

      End If
    End If

    If pFieldArray.Count > 0 Then
      Dim lngCounter As Long
      Dim strBaseName As String
      strBaseName = strName
      Dim booFoundConflict As Boolean
      booFoundConflict = True
      Do Until booFoundConflict = False
        booFoundConflict = False
        For lngIndex = 0 To pFieldArray.Count - 1
          strTestName = pFieldArray.Element(lngIndex)
          If StrComp(strName, strTestName, vbTextCompare) = 0 Then
            booFoundConflict = True
            Exit For
          End If
        Next lngIndex
        If booFoundConflict Then
          lngCounter = lngCounter + 1
          strName = Left(strBaseName, lngMaxLength - Len(CStr(lngCounter))) & CStr(lngCounter)
        End If
      Loop
    End If

    ReturnAcceptableFieldName2 = strName

  End If
  GoTo ClearMemory
ClearMemory:
  Set pField = Nothing
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing

End Function

Public Function CreateInMemoryFeatureClass_Empty(pTemplateFields As esriSystem.IVariantArray, strName As String, _
    pSpRef As ISpatialReference, lngGeometryType As esriGeometryType, booHasM As Boolean, booHasZ As Boolean) As IFeatureClass

    Dim pClone As IClone
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New InMemoryWorkspaceFactory

    Dim pName As IName
    Set pName = pWSF.Create("", "inmemory", Nothing, 0)
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pName.Open

    Dim pFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit

    Set pFields = New Fields
    Set pFieldsEdit = pFields

    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef

    With pGeomDefEdit
      .GeometryType = lngGeometryType
      .GridCount = 1
      .GridSize(0) = 0
      If lngGeometryType = esriGeometryPoint Then
        .AvgNumPoints = 1
      Else
        .AvgNumPoints = 5
      End If
      .HasM = booHasM
      .HasZ = booHasZ
      Set .SpatialReference = pSpRef
    End With

    Set pField = New Field
    Set pFieldEdit = pField

    pFieldEdit.Name = "Shape"
    pFieldEdit.AliasName = "geometry"
    pFieldEdit.Type = esriFieldTypeGeometry
    Set pFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField

    Dim booAddAttribute As Boolean
    booAddAttribute = Not pTemplateFields Is Nothing
    Dim varVal As Variant

    Dim lngIndex As Long
    Dim pTemplateField As iField

    Dim pTempField As iField
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        Set pClone = pTemplateField
        Set pTempField = pClone.Clone
        pFieldsEdit.AddField pTempField
      Next lngIndex
      ReDim lngIDIndex(pTemplateFields.Count - 1)
    Else
      Set pField = New Field
      Set pFieldEdit = pField
      With pFieldEdit
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
      End With
      pFieldsEdit.AddField pField
      ReDim lngIDIndex(0)
    End If

    Dim pCLSID As UID
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Feature"

    Dim pInMemFC As IFeatureClass
    Set pInMemFC = pFWS.CreateFeatureClass(strName, pFields, _
                             pCLSID, Nothing, esriFTSimple, _
                             "Shape", "")

    Set CreateInMemoryFeatureClass_Empty = pInMemFC

  GoTo ClearMemory
ClearMemory:
  Set pClone = Nothing
  Set pSpRefRes = Nothing
  Set pWSF = Nothing
  Set pName = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  varVal = Null
  Set pTemplateField = Nothing
  Set pTempField = Nothing
  Set pCLSID = Nothing
  Set pInMemFC = Nothing

End Function

Public Function ReturnEmptyFClassWithSameSchema(pFClass As IFeatureClass, pWS_NothingForInMemory As IWorkspace, _
    varFieldIndexArray() As Variant, strName As String, booHasFields As Boolean, _
    Optional lngForceGeometryType As esriGeometryType = esriGeometryAny) As IFeatureClass

  Dim pFields As IFields
  Set pFields = pFClass.Fields
  Dim booIsShapefile As Boolean
  Dim booIsAccess As Boolean
  Dim booIsFGDB As Boolean
  Dim booIsInMem As Boolean
  Dim lngCategory As JenDatasetTypes

  If lngForceGeometryType = esriGeometryAny Then
    lngForceGeometryType = pFClass.ShapeType
  End If

  If pWS_NothingForInMemory Is Nothing Then
    booIsInMem = True
  Else
    booIsInMem = False
    booIsShapefile = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "Esri Shapefile Workspace Factory"
    booIsAccess = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "Esri Access Workspace Factory"
    booIsFGDB = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "File GeoDatabase Workspace Factory"
  End If

  If booIsAccess Then
    lngCategory = ENUM_PersonalGDB
  ElseIf booIsFGDB Then
    lngCategory = ENUM_FileGDB
  End If

  Dim lngIndex As Long
  Dim pSrcField As iField
  Dim pNewField As iField
  Dim pNewFieldEdit As IFieldEdit
  Dim pClone As IClone

  Dim pNewFieldArray As esriSystem.IVariantArray
  Set pNewFieldArray = New esriSystem.varArray

  Dim lngCounter As Long
  lngCounter = -1
  Dim varReturnArray() As Variant

  For lngIndex = 0 To pFields.FieldCount - 1
    Set pSrcField = pFields.Field(lngIndex)
    If Not pSrcField.Type = esriFieldTypeOID And pSrcField.Type <> esriFieldTypeGeometry And _
        StrComp(Left(pSrcField.Name, 6), "Shape_", vbTextCompare) <> 0 Then
      Set pClone = pSrcField
      Set pNewField = pClone.Clone
      Set pNewFieldEdit = pNewField
      With pNewFieldEdit
        If booIsShapefile Then
          .Name = MyGeneralOperations.ReturnAcceptableFieldName2(pSrcField.Name, pNewFieldArray, booIsShapefile, booIsAccess, False, booIsFGDB)
          .IsNullable = False
          If pSrcField.Type = esriFieldTypeDouble Then
            .Precision = 16
            .Scale = 6
          End If
        Else
          .IsNullable = True
        End If
      End With
      pNewFieldArray.Add pNewField

      lngCounter = lngCounter + 1
      ReDim Preserve varReturnArray(3, lngCounter)
      varReturnArray(0, lngCounter) = pSrcField.Name
      varReturnArray(1, lngCounter) = lngIndex
      varReturnArray(2, lngCounter) = pNewField.Name

    End If
  Next lngIndex

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Set pDataset = pFClass
  Set pGeoDataset = pFClass
  Dim pGeomDef As IGeometryDef
  Set pGeomDef = pFClass.Fields.Field(pFClass.FindField(pFClass.ShapeFieldName)).GeometryDef

  Dim pNewFClass As IFeatureClass

  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  pEnv.PutCoords -5, -5, 5, 5
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)

  If booIsInMem Then
    Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass_Empty(pNewFieldArray, strName, pSpRef, _
        lngForceGeometryType, pGeomDef.HasM, pGeomDef.HasZ)
  ElseIf booIsFGDB Or booIsAccess Then
    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass2(pWS_NothingForInMemory, strName, esriFTSimple, pSpRef, _
        lngForceGeometryType, pNewFieldArray, , , , False, lngCategory, pEnv, , pGeomDef.HasZ, pGeomDef.HasM)
  ElseIf booIsShapefile Then
    Set pNewFClass = MyGeneralOperations.CreateShapefileFeatureClass2(pWS_NothingForInMemory.PathName, strName, _
        pSpRef, lngForceGeometryType, pNewFieldArray, False, pGeomDef.HasZ, pGeomDef.HasM)
  Else
    MsgBox "No code written for this workspace type!"
    GoTo ClearMemory
  End If

  booHasFields = lngCounter > -1
  If booHasFields Then
    For lngIndex = 0 To lngCounter
      varReturnArray(3, lngIndex) = pNewFClass.FindField(CStr(varReturnArray(2, lngIndex)))
    Next lngIndex
  End If
  varFieldIndexArray = varReturnArray
  Set ReturnEmptyFClassWithSameSchema = pNewFClass

  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pSrcField = Nothing
  Set pNewField = Nothing
  Set pNewFieldEdit = Nothing
  Set pClone = Nothing
  Set pNewFieldArray = Nothing
  Erase varReturnArray
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pGeomDef = Nothing
  Set pNewFClass = Nothing

End Function

Public Function ExportToCSV(pFLayerOrStandaloneTable As IUnknown, strFilename As String, _
    booOverwrite As Boolean, booForceUniqueFilename As Boolean, booOnlySelected As Boolean, _
    booSucceeded As Boolean, Optional varFieldNames As Variant = Null, _
    Optional pApp As IApplication = Nothing, Optional booSkipOID As Boolean = True) As String

  Dim booShowProgress As Boolean
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  booShowProgress = Not pApp Is Nothing

  booSucceeded = True
  Dim lngCount As Long
  Dim lngWriteInterval As Long
  Dim lngCounter As Long

  Dim pStTable As IStandaloneTable
  Dim pRowSel As ITableSelection
  Dim pTable As ITable
  Dim pCursor As ICursor
  Dim pRow As IRow

  Dim pFLayer As IFeatureLayer
  Dim pFeatSel As IFeatureSelection
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature

  Dim pSelSet As ISelectionSet
  Dim strDataName As String
  Dim pFields As IFields

  If Not (TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureClass Or _
          TypeOf pFLayerOrStandaloneTable Is ITable) Then
    booSucceeded = False
    ExportToCSV = "Invalid Data Type:  This function requires either a feature class or table"
    GoTo ClearMemory
  End If

  Dim booIsTable As Boolean
  Dim pDataset As IDataset

  If TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Then
    booIsTable = True
    Set pStTable = pFLayerOrStandaloneTable
    strDataName = pStTable.Name
    Set pTable = pStTable.Table
    Set pFields = pTable.Fields
    If booOnlySelected Then
      Set pRowSel = pStTable
      Set pSelSet = pRowSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV = "No rows selected in table"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pCursor
    Else
      Set pCursor = pTable.Search(Nothing, False)
      lngCount = pTable.RowCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is ITable Then
    booIsTable = True
    Set pTable = pFLayerOrStandaloneTable
    Set pDataset = pTable
    strDataName = pDataset.BrowseName
    Set pFields = pTable.Fields
    Set pCursor = pTable.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  ElseIf TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Then
    booIsTable = False
    Set pFLayer = pFLayerOrStandaloneTable
    strDataName = pFLayer.Name
    Set pFClass = pFLayer.FeatureClass
    Set pFields = pFClass.Fields
    If booOnlySelected Then
      Set pFeatSel = pFLayer
      Set pSelSet = pFeatSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV = "No features selected in feature class"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pFCursor
    Else
      Set pFCursor = pFClass.Search(Nothing, False)
      lngCount = pFClass.FeatureCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is IFeatureClass Then
    booIsTable = True
    Set pFClass = pFLayerOrStandaloneTable
    Set pDataset = pFClass
    strDataName = pDataset.BrowseName
    Set pTable = pFClass
    Set pFields = pTable.Fields
    Set pCursor = pTable.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  End If

  If aml_func_mod.ExistFileDir(strFilename) Then
    If booForceUniqueFilename Then
      strFilename = MakeUniqueFilename(strFilename)
    ElseIf booOverwrite Then
      Kill strFilename
    Else
      ExportToCSV = "File Exists"
      booSucceeded = False
      GoTo ClearMemory
    End If
  End If

  Dim lngFieldCounter As Long
  Dim varFieldIndexes() As Variant
  Dim lngIndex As Long
  Dim pField As iField
  Dim strFieldName As String
  Dim lngFieldIndex As Long
  Dim strReport As String
  Dim lngFileNumber As Long
  Dim lngFieldType As esriFieldType

  lngFieldCounter = -1
  If IsNull(varFieldNames) Then
    For lngIndex = 0 To pTable.Fields.FieldCount - 1
      Set pField = pFields.Field(lngIndex)
      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry And _
              (pField.Type <> esriFieldTypeOID Or Not booSkipOID) Then

        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.AliasName
        varFieldIndexes(1, lngFieldCounter) = lngIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type

        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","

      End If
    Next lngIndex
  Else
    For lngIndex = 0 To UBound(varFieldNames)
      strFieldName = varFieldNames(lngIndex)
      lngFieldIndex = pFields.FindField(strFieldName)
      If lngFieldIndex = -1 Then lngFieldIndex = pFields.FindFieldByAliasName(strFieldName)
      If lngFieldIndex = -1 Then
        ExportToCSV = "Unable to find Attribute Field '" & strFieldName & "'"
        booSucceeded = False
        GoTo ClearMemory
      End If
      Set pField = pFields.Field(lngFieldIndex)

      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry Then
        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.Name
        varFieldIndexes(1, lngFieldCounter) = lngFieldIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type

        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","
      End If
    Next
  End If

  If lngFieldCounter = -1 Then
    ExportToCSV = "No Exportable Fields Found"
    booSucceeded = False
    GoTo ClearMemory
  End If

  If Right(strReport, 1) = "," Then strReport = Left(strReport, Len(strReport) - 1)
  lngFileNumber = FreeFile(0)
  Open strFilename For Output As #lngFileNumber
  Print #lngFileNumber, strReport
  Close #lngFileNumber

  If lngCount < 1000 Then
    lngWriteInterval = 1
  ElseIf lngCount > 1000000 Then
    lngWriteInterval = 100
  Else
    lngWriteInterval = lngCount / 1000
  End If

  If booShowProgress Then
    Set pSBar = pApp.StatusBar
    Set pProg = pSBar.ProgressBar
    pSBar.ShowProgressBar "Exporting '" & strDataName & "' to " & strFilename & "...", 0, lngCount, 1, True
    pProg.position = 0
  End If

  strReport = ""
  If booIsTable Then
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      lngCounter = lngCounter + 1
      If booShowProgress Then pProg.Step
      For lngIndex = 0 To UBound(varFieldIndexes, 2)
        lngFieldIndex = varFieldIndexes(1, lngIndex)
        lngFieldType = varFieldIndexes(2, lngIndex)
        If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
          If IsNull(pRow.Value(lngFieldIndex)) Then
            strReport = strReport & """"","
          Else
            strReport = strReport & aml_func_mod.QuoteString(pRow.Value(lngFieldIndex)) & ","
          End If
        ElseIf lngFieldType = esriFieldTypeDouble Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000000000") & ","
        ElseIf lngFieldType = esriFieldTypeSingle Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000") & ","
        ElseIf lngFieldType = esriFieldTypeDate Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
        ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0") & ","
        Else
          strReport = strReport & CStr(pRow.Value(lngFieldIndex)) & ","
        End If
      Next lngIndex

      strReport = Left(strReport, Len(strReport) - 1)
      If lngCounter >= lngWriteInterval Then
        lngCounter = 0
        lngFileNumber = FreeFile(0)
        Open strFilename For Append As #lngFileNumber
        Print #lngFileNumber, strReport
        Close #lngFileNumber
        strReport = ""
        DoEvents
      Else
        strReport = strReport & vbCrLf
      End If

      Set pRow = pCursor.NextRow
    Loop
  Else
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      If booShowProgress Then pProg.Step
      lngCounter = lngCounter + 1
      For lngIndex = 0 To UBound(varFieldIndexes, 2)
        lngFieldIndex = varFieldIndexes(1, lngIndex)
        lngFieldType = varFieldIndexes(2, lngIndex)
        If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
          strReport = strReport & aml_func_mod.QuoteString(pFeature.Value(lngFieldIndex)) & ","
        ElseIf lngFieldType = esriFieldTypeDouble Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000000000") & ","
        ElseIf lngFieldType = esriFieldTypeSingle Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000") & ","
        ElseIf lngFieldType = esriFieldTypeDate Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
        ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0") & ","
        Else
          strReport = strReport & CStr(pFeature.Value(lngFieldIndex)) & ","
        End If
      Next lngIndex

      strReport = Left(strReport, Len(strReport) - 1)
      If lngCounter >= lngWriteInterval Then
        lngCounter = 0
        lngFileNumber = FreeFile(0)
        Open strFilename For Append As #lngFileNumber
        Print #lngFileNumber, strReport
        Close #lngFileNumber
        strReport = ""
        DoEvents
      Else
        strReport = strReport & vbCrLf
      End If

      Set pFeature = pFCursor.NextFeature
    Loop
  End If

  If strReport <> "" Then
    lngFileNumber = FreeFile(0)
    Open strFilename For Append As #lngFileNumber
    Print #lngFileNumber, strReport
    Close #lngFileNumber
  End If

  If booShowProgress Then
    pSBar.HideProgressBar
    pProg.position = 0
  End If

  booSucceeded = True
  ExportToCSV = "Succeeded"

  GoTo ClearMemory
  Exit Function

ClearMemory:
  Set pStTable = Nothing
  Set pRowSel = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pFLayer = Nothing
  Set pFeatSel = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSelSet = Nothing
  Erase varFieldIndexes
  Set pField = Nothing

End Function

Public Function ExportToCSV_SpecialCases(pFLayerOrStandaloneTable As IUnknown, strFilename As String, _
    booOverwrite As Boolean, booForceUniqueFilename As Boolean, booOnlySelected As Boolean, _
    booSucceeded As Boolean, Optional varFieldNames As Variant = Null, _
    Optional pApp As IApplication = Nothing, Optional booSkipOID As Boolean = True, _
    Optional booCreateAreaAndCentroidFields = False, Optional pPlotLocColl As Collection) As String

  Dim booShowProgress As Boolean
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  booShowProgress = Not pApp Is Nothing

  booSucceeded = True
  Dim lngCount As Long
  Dim lngWriteInterval As Long
  Dim lngCounter As Long

  Dim pStTable As IStandaloneTable
  Dim pRowSel As ITableSelection
  Dim pTable As ITable
  Dim pCursor As ICursor
  Dim pRow As IRow

  Dim pFLayer As IFeatureLayer
  Dim pFeatSel As IFeatureSelection
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature

  Dim pSelSet As ISelectionSet
  Dim strDataName As String
  Dim pFields As IFields
  Dim dblEasting As Double
  Dim dblNorthing As Double
  Dim lngQuadratIDIndex As Long
  Dim strQuad As String

  If TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Or _
          TypeOf pFLayerOrStandaloneTable Is ITable And Not TypeOf pFLayerOrStandaloneTable Is IFeatureClass Then booCreateAreaAndCentroidFields = False

  If Not (TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureClass Or _
          TypeOf pFLayerOrStandaloneTable Is ITable) Then
    booSucceeded = False
    ExportToCSV_SpecialCases = "Invalid Data Type:  This function requires either a feature class or table"
    GoTo ClearMemory
  End If

  Dim booIsTable As Boolean
  Dim pDataset As IDataset

  If TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Then
    booIsTable = False
    Set pFLayer = pFLayerOrStandaloneTable
    strDataName = pFLayer.Name
    Set pFClass = pFLayer.FeatureClass
    Set pFields = pFClass.Fields
    If booOnlySelected Then
      Set pFeatSel = pFLayer
      Set pSelSet = pFeatSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV_SpecialCases = "No features selected in feature class"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pFCursor
    Else
      Set pFCursor = pFClass.Search(Nothing, False)
      lngCount = pFClass.FeatureCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is IFeatureClass Then
    booIsTable = False
    Set pFClass = pFLayerOrStandaloneTable
    Set pDataset = pFClass
    strDataName = pDataset.BrowseName
    Set pTable = pFClass
    Set pFields = pTable.Fields
    Set pFCursor = pFClass.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  ElseIf TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Then
    booIsTable = True
    Set pStTable = pFLayerOrStandaloneTable
    strDataName = pStTable.Name
    Set pTable = pStTable.Table
    Set pFields = pTable.Fields
    If booOnlySelected Then
      Set pRowSel = pStTable
      Set pSelSet = pRowSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV_SpecialCases = "No rows selected in table"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pCursor
    Else
      Set pCursor = pTable.Search(Nothing, False)
      lngCount = pTable.RowCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is ITable Then
    booIsTable = True
    Set pTable = pFLayerOrStandaloneTable
    Set pDataset = pTable
    strDataName = pDataset.BrowseName
    Set pFields = pTable.Fields
    Set pCursor = pTable.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  End If

  If aml_func_mod.ExistFileDir(strFilename) Then
    If booForceUniqueFilename Then
      strFilename = MakeUniqueFilename(strFilename)
    ElseIf booOverwrite Then
      Kill strFilename
    Else
      ExportToCSV_SpecialCases = "File Exists"
      booSucceeded = False
      GoTo ClearMemory
    End If
  End If

  Dim lngFieldCounter As Long
  Dim varFieldIndexes() As Variant
  Dim lngIndex As Long
  Dim pField As iField
  Dim strFieldName As String
  Dim lngFieldIndex As Long
  Dim strReport As String
  Dim lngFileNumber As Long
  Dim lngFieldType As esriFieldType

  lngFieldCounter = -1
  If IsNull(varFieldNames) Then
    For lngIndex = 0 To pTable.Fields.FieldCount - 1
      Set pField = pFields.Field(lngIndex)
      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry And _
              (pField.Type <> esriFieldTypeOID Or Not booSkipOID) Then

        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.AliasName
        varFieldIndexes(1, lngFieldCounter) = lngIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type

        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","

      End If
    Next lngIndex
  Else
    For lngIndex = 0 To UBound(varFieldNames)
      strFieldName = varFieldNames(lngIndex)
      lngFieldIndex = pFields.FindField(strFieldName)
      If lngFieldIndex = -1 Then lngFieldIndex = pFields.FindFieldByAliasName(strFieldName)
      If lngFieldIndex = -1 Then
        ExportToCSV_SpecialCases = "Unable to find Attribute Field '" & strFieldName & "'"
        booSucceeded = False
        GoTo ClearMemory
      End If
      Set pField = pFields.Field(lngFieldIndex)

      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry Then
        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.Name
        varFieldIndexes(1, lngFieldCounter) = lngFieldIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type

        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","
      End If
    Next
  End If

  If booCreateAreaAndCentroidFields Then strReport = strReport & _
      """Easting_NAD83_UTM_Zone_12"",""Northing_NAD83_UTM_Zone_12""," & _
      """X_Coord_In_Quadrat"",""Y_Coord_In_Quadrat"",""Sq_Cm"""

  If lngFieldCounter = -1 Then
    ExportToCSV_SpecialCases = "No Exportable Fields Found"
    booSucceeded = False
    GoTo ClearMemory
  End If

  If Right(strReport, 1) = "," Then strReport = Left(strReport, Len(strReport) - 1)
  lngFileNumber = FreeFile(0)
  Open strFilename For Output As #lngFileNumber
  Print #lngFileNumber, strReport
  Close #lngFileNumber

  If lngCount < 1000 Then
    lngWriteInterval = 1
  ElseIf lngCount > 1000000 Then
    lngWriteInterval = 100
  Else
    lngWriteInterval = lngCount / 1000
  End If

  Dim pPoly As IPolygon
  Dim pArea As IArea

  If booShowProgress Then
    Set pSBar = pApp.StatusBar
    Set pProg = pSBar.ProgressBar
    pSBar.ShowProgressBar "Exporting '" & strDataName & "' to " & strFilename & "...", 0, lngCount, 1, True
    pProg.position = 0
  End If

  Dim booInclude As Boolean
  Dim varArray() As Variant

  strReport = ""
  If booIsTable Then
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      lngCounter = lngCounter + 1
      If booShowProgress Then pProg.Step
      For lngIndex = 0 To UBound(varFieldIndexes, 2)
        lngFieldIndex = varFieldIndexes(1, lngIndex)
        lngFieldType = varFieldIndexes(2, lngIndex)
        If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
          strReport = strReport & aml_func_mod.QuoteString(pRow.Value(lngFieldIndex)) & ","
        ElseIf lngFieldType = esriFieldTypeDouble Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000000000") & ","
        ElseIf lngFieldType = esriFieldTypeSingle Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000") & ","
        ElseIf lngFieldType = esriFieldTypeDate Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
        ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0") & ","
        Else
          strReport = strReport & CStr(pRow.Value(lngFieldIndex)) & ","
        End If
      Next lngIndex

      strReport = Left(strReport, Len(strReport) - 1)
      If lngCounter >= lngWriteInterval Then
        lngCounter = 0
        lngFileNumber = FreeFile(0)
        Open strFilename For Append As #lngFileNumber
        Print #lngFileNumber, strReport
        Close #lngFileNumber
        strReport = ""
        DoEvents
      Else
        strReport = strReport & vbCrLf
      End If

      Set pRow = pCursor.NextRow
    Loop
  Else
    lngQuadratIDIndex = pFCursor.FindField("Quadrat")
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      If booShowProgress Then pProg.Step
      lngCounter = lngCounter + 1

      booInclude = True
      If booCreateAreaAndCentroidFields Then
        strQuad = Trim(CStr(pFeature.Value(lngQuadratIDIndex)))
        strQuad = Replace(strQuad, "*", "")
        strQuad = Trim(strQuad)
        If strQuad = "6" Then
          DoEvents
        End If

        varArray = pPlotLocColl.Item(strQuad)
        dblEasting = varArray(0)
        dblNorthing = varArray(1)

        Set pPoly = pFeature.ShapeCopy
        If pPoly.IsEmpty Then
          booInclude = False
        Else
          Set pArea = pPoly
        End If
      End If

      If booInclude Then
        For lngIndex = 0 To UBound(varFieldIndexes, 2)
          lngFieldIndex = varFieldIndexes(1, lngIndex)
          lngFieldType = varFieldIndexes(2, lngIndex)
          If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
            strReport = strReport & aml_func_mod.QuoteString(pFeature.Value(lngFieldIndex)) & ","
          ElseIf lngFieldType = esriFieldTypeDouble Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000000000") & ","
          ElseIf lngFieldType = esriFieldTypeSingle Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000") & ","
          ElseIf lngFieldType = esriFieldTypeDate Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
          ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0") & ","
          Else
            strReport = strReport & CStr(pFeature.Value(lngFieldIndex)) & ","
          End If
        Next lngIndex

        If booCreateAreaAndCentroidFields Then

          strReport = strReport & Format(pArea.Centroid.x, "0.000000") & "," & _
              Format(pArea.Centroid.Y, "0.000000") & "," & Format(pArea.Centroid.x - dblEasting, "0.000000") & "," & _
              Format(pArea.Centroid.Y - dblNorthing + 1, "0.000000") & "," & Format(pArea.Area * 10000, "0.000000") & ","
        End If

        strReport = Left(strReport, Len(strReport) - 1)
        If lngCounter >= lngWriteInterval Then
          lngCounter = 0
          lngFileNumber = FreeFile(0)
          Open strFilename For Append As #lngFileNumber
          Print #lngFileNumber, strReport
          Close #lngFileNumber
          strReport = ""
          DoEvents
        Else
          strReport = strReport & vbCrLf
        End If
      End If

      Set pFeature = pFCursor.NextFeature
    Loop
  End If

  If strReport <> "" Then
    lngFileNumber = FreeFile(0)
    Open strFilename For Append As #lngFileNumber
    Print #lngFileNumber, strReport
    Close #lngFileNumber
  End If

  If booShowProgress Then
    pSBar.HideProgressBar
    pProg.position = 0
  End If

  booSucceeded = True
  ExportToCSV_SpecialCases = "Succeeded"

  GoTo ClearMemory
  Exit Function

ClearMemory:
  Set pStTable = Nothing
  Set pRowSel = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pFLayer = Nothing
  Set pFeatSel = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSelSet = Nothing
  Erase varFieldIndexes
  Set pField = Nothing

End Function

Public Function CreateOrReturnFileGeodatabase(strPath As String) As IWorkspace

  Dim pWSName As IWorkspaceName
  Dim pName As IName
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory

  If pWSFact.IsWorkspace(strPath) Then
    Set pWS = pWSFact.OpenFromFile(strPath, 0)
  ElseIf pWSFact.IsWorkspace(strPath & ".gdb") Then
    Set pWS = pWSFact.OpenFromFile(strPath & ".gdb", 0)
  Else
    Set pWSName = pWSFact.Create(aml_func_mod.ReturnDir3(strPath, False), _
        aml_func_mod.ReturnFilename2(strPath), Nothing, 0)
    Set pName = pWSName
    Set pWS = pName.Open
  End If

  Set CreateOrReturnFileGeodatabase = pWS

      GoTo ClearMemory

ClearMemory:
  Set pWSName = Nothing
  Set pName = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing

End Function

Public Function ReturnWorkspaceFactoryType(strClassID As String) As String

  Dim pColl As New Collection
  pColl.Add "Esri Access Workspace Factory", "{DD48C96A-D92A-11D1-AA81-00C04FA33A15}"
  pColl.Add "Workspace factory used to create workspace objects for ArcInfo coverages and Info tables", "{1D887452-D9F2-11D1-AA81-00C04FA33A15}"
  pColl.Add "Esri Cad Workspace Factory", "{9E2C27CE-62C6-11D2-9AED-00C04FA33299}"
  pColl.Add "Excel Workspace Factory", "{30F6F271-852B-4EE8-BD2D-099F51D6B238}"
  pColl.Add "FeatureService workspace factory", "{C81194E7-4DAA-418B-8C83-2942E65D2B8C}"
  pColl.Add "File GeoDatabase Workspace Factory", "{71FE75F0-EA0C-4406-873E-B7D53748AE7E}"
  pColl.Add "GeoRSS workspace factory", "{894AF6A1-157A-47BA-BDEC-3CF98D4CCE74}"
  pColl.Add "InMemory workspace factory", "{7F2BC55C-B902-43D0-A566-AA47EA9FDA2C}"
  pColl.Add "Esri LasDataset workspace-factory component", "{7AB01D9A-FDFE-4DFB-9209-86603EE9AEC6}"
  pColl.Add "OleDB Workspace Factory", "{59158055-3171-11D2-AA94-00C04FA37849}"
  pColl.Add "Esri PC ARC/INFO Workspace Factory", "{6DE812D2-9AB6-11D2-B0D7-0000F8780820}"
  pColl.Add "Provides access to members that control creation of raster workspaces", "{4C91D963-3390-11D2-8D25-0000F8780535}"
  pColl.Add "Esri SDC workspace factory", "{34DAE34F-DBE2-409C-8F85-DDBB46138011}"
  pColl.Add "Esri SDE Workspace Factory", "{D9B4FA40-D6D9-11D1-AA81-00C04FA33A15}"
  pColl.Add "Esri Shapefile Workspace Factory", "{A06ADB96-D95C-11D1-AA81-00C04FA33A15}"
  pColl.Add "Sql workspace factory", "{5297187B-FD2B-4A5F-8991-EB3F6F1CA502}"
  pColl.Add "Text File Workspace Factory", "{72CE59EC-0BE8-11D4-AE03-00C04FA33A15}"
  pColl.Add "Esri TIN workspace factory is used to access TINs on disk", "{AD4E89D9-00A5-11D2-B1CA-00C04F8EDEFF}"
  pColl.Add "Workspace Factory used to open toolbox workspaces", "{E9231B31-2A34-4729-8DE2-12CF39674B1B}"
  pColl.Add "Esri VPF Workspace Factory", "{397847F9-C865-11D3-9B56-00C04FA33299}"

  If CheckCollectionForKey(pColl, strClassID) Then
    ReturnWorkspaceFactoryType = pColl.Item(strClassID)
  Else
    ReturnWorkspaceFactoryType = "<- Unknown Workspace Factory Type ->"
  End If

  Set pColl = Nothing

End Function

Public Function CreateNewField(strFieldName As String, lngType As esriFieldType, _
    Optional strAlias As String = "", Optional lngLength As Long = -999, _
    Optional dblPrecision As Double = -999, Optional dblScale As Double = -999) As iField

  Dim pReturn As iField
  Dim pFieldEdit As IFieldEdit

  Set pReturn = New Field
  Set pFieldEdit = pReturn
  With pFieldEdit
    .Name = strFieldName
    If strAlias = "" Then .AliasName = strFieldName
    .Type = lngType
    If lngType = esriFieldTypeString Then
      If lngLength = -999 Then
        .length = 255
      Else
        .length = lngLength
      End If
    End If
    If lngType = esriFieldTypeDouble Then
      If dblPrecision <> -999 Then .Precision = dblPrecision
      If dblScale <> -999 Then .Scale = dblScale
    End If
  End With

  Set CreateNewField = pReturn
  Set pReturn = Nothing
  Set pFieldEdit = Nothing

End Function


