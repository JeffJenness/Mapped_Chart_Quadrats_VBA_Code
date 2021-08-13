Attribute VB_Name = "More_Margaret_Functions"
Option Explicit

Public Sub RecreateSubsetsOfConvertedDatasets()

  Dim strFieldToSplitBy As String
  strFieldToSplitBy = "Site"
  
  ' DON'T BRING IN ANY EMPTY GEOMETRIES OR "No Point Species" OR "No Polygon Species"
  ' NO, THIS HAS CHANGED!  WE WANT EMPTY GEOMETRIES IF SURVEY DONE BUT NO SPECIES OBSERVED.  THIS WILL HELP
  ' USER DISTINGUISH BETWEEN NO ACTUAL PLANTS VS. NO SURVEY CONDUCTED.
  Debug.Print "----------------------------------------"
  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strRecreatedModifiedRoot As String
  
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot, , , strRecreatedModifiedRoot)
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewShapefileWS As IWorkspace
  Dim pNewFeatShapefileWS As IFeatureWorkspace
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Set pNewFGDBWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strRecreatedModifiedRoot & "\Combined_by_Site")
  Set pNewFeatFGDBWS = pNewFGDBWS
  
  If Not aml_func_mod.ExistFileDir(strRecreatedModifiedRoot & "\Shapefiles") Then
    MyGeneralOperations.CreateNestedFoldersByPath (strRecreatedModifiedRoot & "\Shapefiles")
  End If
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pNewShapefileWS = pWSFact.OpenFromFile(strRecreatedModifiedRoot & "\Shapefiles", 0)
  Set pNewFeatShapefileWS = pNewShapefileWS

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose
    
  ' FIRST GET LIST OF NEW NAMES
  Dim pCollOfFClasses As New Collection    ' FOR EACH ID VALUE, WILL CONTAIN VARIANT ARRAY OF [FClass Name, pFClass, pInsertCursor, pInsertBuffer]
  Dim pCollOfShapefiles As New Collection  ' FOR EACH ID VALUE, WILL CONTAIN VARIANT ARRAY OF [FClass Name, pFClass, pInsertCursor, pInsertBuffer]
  Dim varItems() As Variant
  Dim strNewFClassName As String
  Dim strNewShapefileName As String
  Dim varGDBFieldIndexes() As Variant
  Dim varShapefileFieldIndexes() As Variant
  Dim pDataset As IDataset
  Dim strArrayOfGDBNames() As String
  Dim strArrayOfShapefileNames() As String
  
  ' SOURCE DATA
  Dim pCoverFClass As IFeatureClass
  Dim pCoverFCursor As IFeatureCursor
  Dim pCoverFeature As IFeature
  
  Dim pDensityFClass As IFeatureClass
  Dim pDensityFCursor As IFeatureCursor
  Dim pDensityFeature As IFeature
  
  ' NEW COMPREHENSIVE DATA
  ' ...GEODATABASE
  ' ......COVER
  Dim pNewCoverGDBFClass As IFeatureClass
  Dim pNewGDBCoverFCursor As IFeatureCursor
  Dim varNewGDBCoverFieldIndexes() As Variant
  Dim pNewGDBCoverFBuffer As IFeatureBuffer
  ' ......DENSITY
  Dim pNewDensityGDBFClass As IFeatureClass
  Dim pNewGDBDensityFCursor As IFeatureCursor
  Dim varNewGDBDensityFieldIndexes() As Variant
  Dim pNewGDBDensityFBuffer As IFeatureBuffer
  ' ...SHAPEFILE
  ' ......COVER
  Dim pNewCoverShapefile As IFeatureClass
  Dim pNewShpCoverFCursor As IFeatureCursor
  Dim varNewShpCoverFieldIndexes() As Variant
  Dim pNewShpCoverFBuffer As IFeatureBuffer
  ' ......DENSITY
  Dim pNewDensityShapefile As IFeatureClass
  Dim pNewShpDensityFCursor As IFeatureCursor
  Dim varNewShpDensityFieldIndexes() As Variant
  Dim pNewShpDensityFBuffer As IFeatureBuffer
  
  ' NEW SITE DATA:  DON'T NEED TO SEPARATE INTO COVER VS. DENSITY
  ' ...GEODATABASE
  Dim pNewGeneralGDBFClass As IFeatureClass
  Dim pNewGDBGeneralFCursor As IFeatureCursor
  Dim varNewGDBGeneralFieldIndexes() As Variant
  Dim pNewGDBGeneralFBuffer As IFeatureBuffer
  ' ...SHAPEFILE
  Dim pNewGeneralShpFClass As IFeatureClass
  Dim pNewShpGeneralFCursor As IFeatureCursor
  Dim varNewShpGeneralFieldIndexes() As Variant
  Dim pNewShpGeneralFBuffer As IFeatureBuffer
  
  Dim lngIDIndex As Long
  Dim strIDVal As String
  Dim lngSpeciesIndex As Long
  Dim strSpecies As String
  Dim varVal As Variant
  Dim booCheckIfShouldContinue As Boolean
  
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  
  lngCount = pCoverFClass.FeatureCount(Nothing) + pDensityFClass.FeatureCount(Nothing)
  lngCounter = 0
  pSBar.ShowProgressBar "Pass 1:  Transferring Cover features...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pCoverFCursor = pCoverFClass.Search(Nothing, False)
  Set pCoverFeature = pCoverFCursor.NextFeature
  lngIDIndex = pCoverFClass.FindField(strFieldToSplitBy)
  lngSpeciesIndex = pCoverFClass.FindField("Species")
  
  FillVariousFClassObjects pCollOfFClasses, "Cover_All", pNewFeatFGDBWS, pCoverFClass, pNewCoverGDBFClass, _
      pNewGDBCoverFCursor, pNewGDBCoverFBuffer, strNewFClassName, varNewGDBCoverFieldIndexes, strArrayOfGDBNames, _
      strAbstract, strBaseString, strPurpose, pMxDoc
      
  FillVariousFClassObjects pCollOfShapefiles, "Cover_All", pNewFeatShapefileWS, pCoverFClass, pNewCoverShapefile, _
      pNewShpCoverFCursor, pNewShpCoverFBuffer, strNewFClassName, varNewShpCoverFieldIndexes, strArrayOfShapefileNames, _
      strAbstract, strBaseString, strPurpose, pMxDoc
  
  Do Until pCoverFeature Is Nothing
    UpdateCount lngCounter, pProg, lngCount
    booCheckIfShouldContinue = CheckGeometryAndSpecies(pCoverFeature, lngSpeciesIndex, strSpecies)
    If booCheckIfShouldContinue Then
      
      strIDVal = pCoverFeature.Value(lngIDIndex) & "_Cover"
      FillVariousFClassObjects pCollOfFClasses, strIDVal, pNewFeatFGDBWS, pCoverFClass, pNewGeneralGDBFClass, _
          pNewGDBGeneralFCursor, pNewGDBGeneralFBuffer, strNewFClassName, varNewGDBGeneralFieldIndexes, strArrayOfGDBNames, _
          strAbstract, strBaseString, strPurpose, pMxDoc
      FillVariousFClassObjects pCollOfShapefiles, strIDVal, pNewFeatShapefileWS, pCoverFClass, pNewGeneralShpFClass, _
          pNewShpGeneralFCursor, pNewShpGeneralFBuffer, strNewFClassName, varNewShpGeneralFieldIndexes, strArrayOfShapefileNames, _
          strAbstract, strBaseString, strPurpose, pMxDoc
      
      ' WRITE VALUES
      WriteValues pCoverFeature, pNewGDBCoverFCursor, pNewGDBCoverFBuffer, varNewGDBCoverFieldIndexes, False
      WriteValues pCoverFeature, pNewShpCoverFCursor, pNewShpCoverFBuffer, varNewShpCoverFieldIndexes, True
      
      WriteValues pCoverFeature, pNewGDBGeneralFCursor, pNewGDBGeneralFBuffer, varNewGDBGeneralFieldIndexes, False
      WriteValues pCoverFeature, pNewShpGeneralFCursor, pNewShpGeneralFBuffer, varNewShpGeneralFieldIndexes, True
      
      If lngCounter Mod 2000 = 0 Then
        FlushAllDatasets pCollOfFClasses, pCollOfShapefiles, strArrayOfGDBNames, strArrayOfShapefileNames
      End If
      
    End If
    Set pCoverFeature = pCoverFCursor.NextFeature
  Loop
  FlushAllDatasets pCollOfFClasses, pCollOfShapefiles, strArrayOfGDBNames, strArrayOfShapefileNames
  
'    Dim pCollOfFClasses As New Collection    ' FOR EACH ID VALUE, WILL CONTAIN VARIANT ARRAY OF [FClass Name, pFClass, pInsertCursor, pInsertBuffer]

  Set pDensityFCursor = pDensityFClass.Search(Nothing, False)
  Set pDensityFeature = pDensityFCursor.NextFeature
  lngIDIndex = pDensityFClass.FindField(strFieldToSplitBy)
  lngSpeciesIndex = pDensityFClass.FindField("Species")
  
  FillVariousFClassObjects pCollOfFClasses, "Density_All", pNewFeatFGDBWS, pDensityFClass, pNewDensityGDBFClass, _
      pNewGDBDensityFCursor, pNewGDBDensityFBuffer, strNewFClassName, varNewGDBDensityFieldIndexes, strArrayOfGDBNames, _
      strAbstract, strBaseString, strPurpose, pMxDoc
      
  FillVariousFClassObjects pCollOfShapefiles, "Density_All", pNewFeatShapefileWS, pDensityFClass, pNewDensityShapefile, _
      pNewShpDensityFCursor, pNewShpDensityFBuffer, strNewFClassName, varNewShpDensityFieldIndexes, strArrayOfShapefileNames, _
      strAbstract, strBaseString, strPurpose, pMxDoc
      
  Do Until pDensityFeature Is Nothing
    UpdateCount lngCounter, pProg, lngCount
    booCheckIfShouldContinue = CheckGeometryAndSpecies(pDensityFeature, lngSpeciesIndex, strSpecies)
    If booCheckIfShouldContinue Then
      
      strIDVal = pDensityFeature.Value(lngIDIndex) & "_Density"
      FillVariousFClassObjects pCollOfFClasses, strIDVal, pNewFeatFGDBWS, pDensityFClass, pNewGeneralGDBFClass, _
          pNewGDBGeneralFCursor, pNewGDBGeneralFBuffer, strNewFClassName, varNewGDBGeneralFieldIndexes, strArrayOfGDBNames, _
          strAbstract, strBaseString, strPurpose, pMxDoc
      FillVariousFClassObjects pCollOfShapefiles, strIDVal, pNewFeatShapefileWS, pDensityFClass, pNewGeneralShpFClass, _
          pNewShpGeneralFCursor, pNewShpGeneralFBuffer, strNewFClassName, varNewShpGeneralFieldIndexes, strArrayOfShapefileNames, _
          strAbstract, strBaseString, strPurpose, pMxDoc
      
      ' WRITE VALUES
      WriteValues pDensityFeature, pNewGDBDensityFCursor, pNewGDBDensityFBuffer, varNewGDBDensityFieldIndexes, False
      WriteValues pDensityFeature, pNewShpDensityFCursor, pNewShpDensityFBuffer, varNewShpDensityFieldIndexes, True
      
      WriteValues pDensityFeature, pNewGDBGeneralFCursor, pNewGDBGeneralFBuffer, varNewGDBGeneralFieldIndexes, False
      WriteValues pDensityFeature, pNewShpGeneralFCursor, pNewShpGeneralFBuffer, varNewShpGeneralFieldIndexes, True
      
      If lngCounter Mod 2000 = 0 Then
        FlushAllDatasets pCollOfFClasses, pCollOfShapefiles, strArrayOfGDBNames, strArrayOfShapefileNames
      End If
      
    End If
  
    Set pDensityFeature = pDensityFCursor.NextFeature
  Loop
  FlushAllDatasets pCollOfFClasses, pCollOfShapefiles, strArrayOfGDBNames, strArrayOfShapefileNames
  CreateAttributeFieldIndexes pCollOfFClasses, pCollOfShapefiles, strArrayOfGDBNames, strArrayOfShapefileNames
  
  pProg.position = 0
  pSBar.HideProgressBar
  
  Debug.Print "Done..."
  Debug.Print "Completed at " & Format(Now, "long time")
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewShapefileWS = Nothing
  Set pNewFeatShapefileWS = Nothing
  Set pCollOfFClasses = Nothing
  Set pCollOfShapefiles = Nothing
  Erase varItems
  Erase varGDBFieldIndexes
  Erase varShapefileFieldIndexes
  Set pDataset = Nothing
  Set pCoverFClass = Nothing
  Set pCoverFCursor = Nothing
  Set pCoverFeature = Nothing
  Set pDensityFClass = Nothing
  Set pDensityFCursor = Nothing
  Set pDensityFeature = Nothing
  Set pNewCoverGDBFClass = Nothing
  Set pNewGDBCoverFCursor = Nothing
  Erase varNewGDBCoverFieldIndexes
  Set pNewGDBCoverFBuffer = Nothing
  Set pNewDensityGDBFClass = Nothing
  Set pNewGDBDensityFCursor = Nothing
  Erase varNewGDBDensityFieldIndexes
  Set pNewGDBDensityFBuffer = Nothing
  Set pNewCoverShapefile = Nothing
  Set pNewShpCoverFCursor = Nothing
  Erase varNewShpCoverFieldIndexes
  Set pNewShpCoverFBuffer = Nothing
  Set pNewDensityShapefile = Nothing
  Set pNewShpDensityFCursor = Nothing
  Erase varNewShpDensityFieldIndexes
  Set pNewShpDensityFBuffer = Nothing
  Set pNewGeneralGDBFClass = Nothing
  Set pNewGDBGeneralFCursor = Nothing
  Erase varNewGDBGeneralFieldIndexes
  Set pNewGDBGeneralFBuffer = Nothing
  Set pNewGeneralShpFClass = Nothing
  Set pNewShpGeneralFCursor = Nothing
  Erase varNewShpGeneralFieldIndexes
  Set pNewShpGeneralFBuffer = Nothing
  varVal = Null




  
End Sub

Public Sub CreateAttributeFieldIndexes(pCollOfGDBDatasets As Collection, pCollOfShapefiles As Collection, _
    strArrayOfGDBNames() As String, strArrayOfShapefileNames)
    
  Dim lngIndex As Long
  Dim strName As String
  Dim varData() As Variant
  Dim pFClass As IFeatureClass
  
  For lngIndex = 0 To UBound(strArrayOfGDBNames)
    strName = strArrayOfGDBNames(lngIndex)
    If MyGeneralOperations.CheckCollectionForKey(pCollOfGDBDatasets, strName) Then
      varData = pCollOfGDBDatasets.Item(strName)
      Set pFClass = varData(1)
      CreateIndexesForSpecificFClass pFClass
      DoEvents
    End If
  Next lngIndex
  For lngIndex = 0 To UBound(strArrayOfShapefileNames)
    strName = strArrayOfShapefileNames(lngIndex)
    If MyGeneralOperations.CheckCollectionForKey(pCollOfShapefiles, strName) Then
      varData = pCollOfShapefiles.Item(strName)
      Set pFClass = varData(1)
      CreateIndexesForSpecificFClass pFClass
      DoEvents
    End If
  Next lngIndex
  
ClearMemory:
  Erase varData
  Set pFClass = Nothing

End Sub

Public Sub CreateIndexesForSpecificFClass(pFClass As IFeatureClass)

  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "z_Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pFClass, "Plot")

End Sub

Public Sub FlushAllDatasets(pCollOfGDBDatasets As Collection, pCollOfShapefiles As Collection, _
    strArrayOfGDBNames() As String, strArrayOfShapefileNames)
    
  Dim lngIndex As Long
  Dim strName As String
  Dim varData() As Variant
  Dim pFCursor As IFeatureCursor
  
  For lngIndex = 0 To UBound(strArrayOfGDBNames)
    strName = strArrayOfGDBNames(lngIndex)
    If MyGeneralOperations.CheckCollectionForKey(pCollOfGDBDatasets, strName) Then
      varData = pCollOfGDBDatasets.Item(strName)
      Set pFCursor = varData(2)
      pFCursor.Flush
    End If
  Next lngIndex
  For lngIndex = 0 To UBound(strArrayOfShapefileNames)
    strName = strArrayOfShapefileNames(lngIndex)
    If MyGeneralOperations.CheckCollectionForKey(pCollOfShapefiles, strName) Then
      varData = pCollOfShapefiles.Item(strName)
      Set pFCursor = varData(2)
      pFCursor.Flush
    End If
  Next lngIndex
  
ClearMemory:
  Erase varData
  Set pFCursor = Nothing



End Sub

Public Sub WriteValues(pSrcFeature As IFeature, pDestFCursor As IFeatureCursor, pDestFBuffer As IFeatureBuffer, _
      varFieldIndexArray() As Variant, booIsShapefile As Boolean)
      
  ' varFieldIndexArray WILL HAVE 4 COLUMNS AND ANY NUMBER OR ROWS.
  ' COLUMN 0 = SOURCE FIELD NAME
  ' COLUMN 1 = SOURCE FIELD INDEX
  ' COLUMN 2 = NEW FIELD NAME
  ' COLUMN 3 = NEW FIELD INDEX
  
  Dim lngIndex As Long
  Set pDestFBuffer.Shape = pSrcFeature.ShapeCopy
  
  Dim varSrcVal As Variant
  Dim lngVarType As Long
  Dim lngDestIndex As Long
  
  For lngIndex = 0 To UBound(varFieldIndexArray, 2)
    varSrcVal = pSrcFeature.Value(varFieldIndexArray(1, lngIndex))
    lngDestIndex = varFieldIndexArray(3, lngIndex)
    If booIsShapefile Then
      If IsNull(varSrcVal) Then
        If pDestFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
          varSrcVal = -999
        ElseIf pDestFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
          varSrcVal = -999
        ElseIf pDestFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
          varSrcVal = ""
        End If
      End If
    End If
    pDestFBuffer.Value(lngDestIndex) = varSrcVal
  Next lngIndex
  pDestFCursor.InsertFeature pDestFBuffer
  
End Sub

Public Sub FillVariousFClassObjects(pCollOfDatasets As Collection, strIDVal As String, pNewWS As IFeatureWorkspace, _
      pSourceFClass As IFeatureClass, pNewFClass As IFeatureClass, pNewFCursor As IFeatureCursor, _
      pNewFBuffer As IFeatureBuffer, strNewFClassName As String, varFieldIndexes() As Variant, _
      strArrayOfNames() As String, strAbstract As String, strBaseString As String, strPurpose As String, _
      pMxDoc As IMxDocument)
  
  Dim varItems() As Variant
  Dim pDataset As IDataset
  Dim lngArrayCounter As Long
  
  If IsDimmed(strArrayOfNames) Then
    lngArrayCounter = UBound(strArrayOfNames)
  Else
    lngArrayCounter = -1
  End If
  
  If MyGeneralOperations.CheckCollectionForKey(pCollOfDatasets, strIDVal) Then
    varItems = pCollOfDatasets.Item(strIDVal)
    strNewFClassName = varItems(0)
    Set pNewFClass = varItems(1)
    Set pNewFCursor = varItems(2)
    Set pNewFBuffer = varItems(3)
  Else
    ReDim varItems(3)
    strNewFClassName = ReplaceBadChars(strIDVal, True, True, True, True)
    Do Until InStr(1, strNewFClassName, "__", vbTextCompare) = 0
      strNewFClassName = Replace(strNewFClassName, "__", "_")
    Loop
    
    lngArrayCounter = lngArrayCounter + 1
    ReDim Preserve strArrayOfNames(lngArrayCounter)
    strArrayOfNames(lngArrayCounter) = strIDVal
    
    If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewFClassName) Then
      Set pDataset = pNewWS.OpenFeatureClass(strNewFClassName)
      pDataset.DELETE
    End If
    
    Set pNewFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSourceFClass, _
          pNewWS, varFieldIndexes, strNewFClassName, True)
    Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewFClass, strAbstract, strPurpose)
    
    Set pNewFCursor = pNewFClass.Insert(True)
    Set pNewFBuffer = pNewFClass.CreateFeatureBuffer
    varItems(0) = strNewFClassName
    varItems(1) = pNewFClass
    varItems(2) = pNewFCursor
    varItems(3) = pNewFBuffer
    pCollOfDatasets.Add varItems, strIDVal
  End If
  
ClearMemory:
  Erase varItems
  Set pDataset = Nothing

End Sub

Public Function CheckGeometryAndSpecies(pFeature As IFeature, lngSpeciesIndex As Long, strSpecies As String) As Boolean
  
  CheckGeometryAndSpecies = True
  strSpecies = ""
    
  If pFeature.ShapeCopy.IsEmpty Then
    CheckGeometryAndSpecies = False
  ElseIf IsNull(pFeature.Value(lngSpeciesIndex)) Then
    CheckGeometryAndSpecies = False
  Else
    strSpecies = pFeature.Value(lngSpeciesIndex)
    If strSpecies = "No Point Species" Or strSpecies = "No Polygon Species" Or Trim(strSpecies = "") Then
      CheckGeometryAndSpecies = False
    End If
  End If
    
End Function

Public Sub UpdateCount(lngCounter As Long, pProg As IStepProgressor, lngCount As Long)

  lngCounter = lngCounter + 1
  pProg.Step
  If lngCounter Mod 1000 = 0 Then
    pProg.Message = "Pass 1:  Transferring Cover features [" & Format(lngCounter, "#,##0") & " of " & _
        Format(lngCount, "#,##0") & "]..."
    DoEvents
  End If

End Sub

Public Sub FillMetadataItems(strAbstract As String, strBaseString As String, strPurpose As String)

  strBaseString = strBaseString & "Margaret M. Moore[1] (margaret.moore@nau.edu) *" & vbNewLine
  strBaseString = strBaseString & "Jeffrey S. Jenness[2] (jeffj@jennessent.com) *" & vbNewLine
  strBaseString = strBaseString & "Daniel C. Laughlin[3] (daniel.laughlin@uwyo.edu)" & vbNewLine
  strBaseString = strBaseString & "Robert T. Strahan[4] (strahanr@sou.edu)" & vbNewLine
  strBaseString = strBaseString & "Jonathan D. Bakker[5] (jdbakker@uw.edu)" & vbNewLine
  strBaseString = strBaseString & "Helen E. Dowling[6] (ldowling@pheasantsforever.org)   " & vbNewLine
  strBaseString = strBaseString & "Judith D. Springer[7] (judy.springer@nau.edu)" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine
  strBaseString = strBaseString & "[1] School of Forestry, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "[2] Jenness Enterprises, GIS Analysis and Application Design, Flagstaff, AZ  86004 USA" & vbNewLine
  strBaseString = strBaseString & "[3] Department of Botany, University of Wyoming, Laramie, WY 82072 USA" & vbNewLine
  strBaseString = strBaseString & "[4] Southern Oregon University, Ashland, OR 97520 USA" & vbNewLine
  strBaseString = strBaseString & "[5] School of Environmental and Forest Sciences, University of Washington, Seattle, WA  98195 USA" & vbNewLine
  strBaseString = strBaseString & "[6] Pheasants Forever, Waterville, WA 98858 USA" & vbNewLine
  strBaseString = strBaseString & "[7] Ecological Restoration Institute, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "*Corresponding authors: Margaret M. Moore [E-mail: Margaret.Moore@nau.edu] and " & _
      "Jeffrey S. Jenness [E-mail: jeffj@jennessent.com]" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine

  strAbstract = "This dataset consists of 98 permanent 1-m2 quadrats located on ponderosa pine" & _
      "–bunchgrass ecosystems in or near Flagstaff, Arizona, USA.  Individual plants in " & _
      "these quadrats were identified and mapped annually from 2002-2020.  The temporal and spatial data provide " & _
      "unique opportunities to examine the effects of climate and " & _
      "land-use variables on plant demography, population and community processes.  The original chart quadrats were " & _
      "established between 1912 and 1927 to determine the effects of livestock grazing on herbaceous plants and pine " & _
      "seedlings.  We provide the following data and data formats: " & vbNewLine
      strAbstract = strAbstract & _
      "(1) Digitized maps in shapefile and file geodatabase format" & vbNewLine & _
      "(2) Shapefiles of each individual species observed on each quadrat in each year, organized into separate subfolders" & _
      " by Species, Site and Quadrat.  These single-species shapefiles are formatted to input directly into Integral Projection" & _
      " Modeling (IPM) analysis in R." & vbNewLine & _
      "(3) Tabular representation of centroid or point location (x, y coordinates) for species mapped as points" & vbNewLine & _
      "(4) Tabular representation of basal cover for species mapped as polygons" & vbNewLine & _
      "(5) Species list including synonymy of names and plant growth forms" & vbNewLine & _
      "(6) Inventory of the years each quadrat was sampled" & vbNewLine & _
      "(7) Counts of each species recorded at each site and quadrat" & _
      "(8) Tree density and basal area records for overstory plots that surround each quadrat" & vbNewLine & _
      "(9) Quadrat centerpoint coordinates in UTM Zone 12 and Latitude/Longitude coordinates, both in North American Datum of 1983" & vbNewLine & _
      "(10) TIFF and PDF maps of all sites and years (n = 1,523 maps)"
      strAbstract = strAbstract & vbCrLf & vbCrLf & strBaseString

  strPurpose = "An analysis of cover and density of southwestern ponderosa pine-bunchgrass plants mapped multiple times " & _
      "between 2002 and 2020 in permanent quadrats."
 
End Sub

Public Sub RepairOverlappingPolygons()
  
  ' FIRST GET LISTS OF SPECIES AND PLOTS
  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot)
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityQuadratIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverQuadratIndex As Long
  Dim lngCoverSiteIndex As Long
  Dim lngCoverPlotIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityQuadratIndex = pDensityFClass.FindField("Quadrat")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
  
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverQuadratIndex = pCoverFClass.FindField("Quadrat")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
  
  Dim pSiteLookup As New Collection
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strSpecies As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllSpecies() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneSpecies As New Collection
  
  lngAllSpeciesIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    
    strSite = pFeature.Value(lngDensityQuadratIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
      
      pSiteLookup.Add pFeature.Value(lngDensitySiteIndex) & ", " & pFeature.Value(lngDensityPlotIndex), strSite
      
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    
    strSite = pFeature.Value(lngCoverQuadratIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
      
      pSiteLookup.Add pFeature.Value(lngCoverSiteIndex) & ", " & pFeature.Value(lngCoverPlotIndex), strSite
      
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
'  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  
  ' NOW GO THROUGH EACH UNIQUE COMBINATION OF PLOT, YEAR AND SPECIES.
  ' SELECT ALL CASES WHERE THAT SPECIES EXISTS IN THAT PLOT AND YEAR. BY GOING THROUGH
  Dim pQueryFilt As IQueryFilter
  Set pQueryFilt = New QueryFilter
  Dim lngYear As Long
  Dim strPrefix As String
  Dim strSuffix As String
  Dim strQuadrat As String
  Dim pSpFilt As ISpatialFilter
  Set pSpFilt = New SpatialFilter
  pSpFilt.SpatialRel = esriSpatialRelIntersects
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pCoverFClass, strPrefix, strSuffix)
  Dim lngTotalCount As Long
  Dim strSitePlot As String
  
  Dim strCheckReport As String
      
  pSBar.ShowProgressBar "Pass 2:  Working through sites and years...", 0, 18 * (UBound(strAllSites) + 1), 1, True
  pProg.position = 0
  
  lngCounter = 0
  Dim pFCursor2 As IFeatureCursor
  Dim pFeature2 As IFeature
  Dim lngCoverCount As Long
  Dim lngDensityCount As Long
  Dim strCoverQueryString As String
  Dim strDensityQueryString As String
  Dim pDoneQueryStrings As New Collection
  Dim strQueryLine As String
  Dim pPoly1 As IPolygon
  Dim pPoly2 As IPolygon
  Dim dblDist As Double
  Dim booFoundIntersect As Boolean
  Dim lngIntersectCounter As Long
  
  Dim lngCoverOIDArray() As Long
  Dim lngDensityOIDArray() As Long
  Dim lngCoverArrayIndex As Long
  Dim lngDensityArrayIndex As Long
  Dim pBufferPoly As IPolygon
  
  For lngYear = 2002 To 2020
    DoEvents
    For lngIndex = 0 To UBound(strAllSites)
      If lngIndex Mod 10 = 0 Then
        DoEvents
      End If
      pProg.Step
      
      strQuadrat = strAllSites(lngIndex)
      pQueryFilt.WhereClause = strPrefix & "Quadrat" & strSuffix & " = '" & strQuadrat & "' AND " & _
          strPrefix & "Year" & strSuffix & " = '" & Format(lngYear, "0") & "'"
      
      ' MAKE LIST OF SPECIES OBSERVED ON THIS QUADRAT THIS YEAR
      
'      Debug.Print "Quadrat = " & strQuadrat & "; Year = " & Format(lngYear, "0") & vbCrLf & _
'          "  --> Cover Count = " & Format(pCoverFClass.FeatureCount(pQueryFilt), "#,##0") & vbCrLf & _
'          "  --> Density Count = " & Format(pDensityFClass.FeatureCount(pQueryFilt), "#,##0")
                  
      lngAllSpeciesIndex = -1
      Erase strAllSpecies
      Set pDoneSpecies = New Collection
        
      Set pFCursor = pDensityFClass.Search(pQueryFilt, False)
      Set pFeature = pFCursor.NextFeature
      Do Until pFeature Is Nothing
        strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
        If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
        If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
          pDoneSpecies.Add True, strSpecies
          lngAllSpeciesIndex = lngAllSpeciesIndex + 1
          ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
          strAllSpecies(lngAllSpeciesIndex) = strSpecies
        End If
        
        Set pFeature = pFCursor.NextFeature
      Loop
        
      Set pFCursor = pCoverFClass.Search(pQueryFilt, False)
      Set pFeature = pFCursor.NextFeature
      Do Until pFeature Is Nothing
        strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
        If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
        If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
          pDoneSpecies.Add True, strSpecies
          lngAllSpeciesIndex = lngAllSpeciesIndex + 1
          ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
          strAllSpecies(lngAllSpeciesIndex) = strSpecies
        End If
        
        Set pFeature = pFCursor.NextFeature
      Loop
      
      If lngAllSpeciesIndex > -1 Then
        DoEvents
      End If
      
'      Debug.Print "Quadrat = " & strQuadrat & "; Year = " & Format(lngYear, "0") & vbCrLf & _
'          "  --> Cover Count = " & Format(pCoverFClass.FeatureCount(pQueryFilt), "#,##0") & vbCrLf & _
'          "  --> Density Count = " & Format(pDensityFClass.FeatureCount(pQueryFilt), "#,##0") & vbCrLf & _
'          "  --> Found " & Format(lngAllSpeciesIndex + 1, "0") & " unique species"
      
      ' NOW HAVE LIST OF ALL SPECIES OBSERVED ON THIS QUADRAT THIS YEAR.
      
      If lngAllSpeciesIndex > -1 Then
        For lngIndex2 = 0 To lngAllSpeciesIndex
          strSpecies = strAllSpecies(lngIndex2)
          If strSpecies <> "No Polygon Species" Then
            pQueryFilt.WhereClause = strPrefix & "Quadrat" & strSuffix & " = '" & strQuadrat & "' AND " & _
                strPrefix & "Year" & strSuffix & " = '" & Format(lngYear, "0") & "' AND " & _
                strPrefix & "Species" & strSuffix & " = '" & strSpecies & "'"
            pSpFilt.WhereClause = pQueryFilt.WhereClause
            
            ' WORK THOUGH EACH INSTANCE OF A PARTICULAR SPECIES IN THIS QUADRAT IN THIS YEAR.
            ' CHECK TO SEE HOW MANY COVER AND DENSITY FEATURES WITH THIS SPECIES INTERSECT THIS INSTANCE.  SHOULD BE EXACTLY ONE (ITSELF).
            ' FIRST COVER
            lngCoverArrayIndex = -1
            lngDensityArrayIndex = -1
            Erase lngCoverOIDArray
            Erase lngDensityOIDArray
            
            Set pFCursor = pCoverFClass.Search(pQueryFilt, False)
            Set pFeature = pFCursor.NextFeature
            Do Until pFeature Is Nothing
              
              If pFeature.OID = 27974 Then
                DoEvents
              End If
              
              Set pPoly1 = pFeature.ShapeCopy
              If Not pPoly1.IsEmpty Then
                Set pBufferPoly = ExpandSmallPolygonByBufferDist(pPoly1)
              Else
                Set pBufferPoly = pPoly1
              End If
              
'              Set pSpFilt.Geometry = pFeature.ShapeCopy
              Set pSpFilt.Geometry = pBufferPoly
              lngCoverCount = pCoverFClass.FeatureCount(pSpFilt)
              lngDensityCount = pDensityFClass.FeatureCount(pSpFilt)
              lngTotalCount = lngCoverCount + lngDensityCount
              
              If lngTotalCount <> 1 Then
                strCoverQueryString = ""
                strDensityQueryString = ""
                strQueryLine = ""
                booFoundIntersect = False
                lngIntersectCounter = 0
                
                strSitePlot = pSiteLookup.Item(strQuadrat)
                strQueryLine = strQuadrat & " (" & strSitePlot & "), " & Format(lngYear, "0") & vbCrLf & _
                    "  --> Species = " & strSpecies & vbCrLf & _
                    "  --> Found " & Format(lngTotalCount, "#,##0") & " overlapping features" & vbCrLf
                
                If lngCoverCount = 0 And lngDensityCount = 0 Then
                  If lngCoverCount = 0 Then
                    strCoverQueryString = strPrefix & pCoverFClass.OIDFieldName & strSuffix & _
                        " = " & Format(pFeature.OID, "0")
                    strQueryLine = strQueryLine & "  --> EMPTY GEOMETRY!!!  Cover Query String: " & strCoverQueryString & vbCrLf
                    booFoundIntersect = True  ' JUST TO FORCE EVENT TO BE WRITTEN TO REPORT
                    lngIntersectCounter = 2 ' JUST TO FORCE EVENT TO BE WRITTEN TO REPORT
                  End If
                End If
                
                If lngCoverCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Cover Polygon OID values = "
                  Set pFCursor2 = pCoverFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                      lngCoverArrayIndex = lngCoverArrayIndex + 1
                      ReDim Preserve lngCoverOIDArray(lngCoverArrayIndex)
                      lngCoverOIDArray(lngCoverArrayIndex) = pFeature2.OID
                    
                      lngIntersectCounter = lngIntersectCounter + 1
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strCoverQueryString = strCoverQueryString & _
                          strPrefix & pCoverFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      
                      booFoundIntersect = True
                    
                    End If
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  
                  If lngCoverArrayIndex > -1 Then
                    strCoverQueryString = Left(strCoverQueryString, Len(strCoverQueryString) - 4)
                    strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                        "  --> Cover Query String: " & strCoverQueryString & vbCrLf
                  End If
                End If
                
                If lngDensityCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Density Polygon OID values = "
                  Set pFCursor2 = pDensityFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                    
                      lngDensityArrayIndex = lngDensityArrayIndex + 1
                      ReDim Preserve lngDensityOIDArray(lngDensityArrayIndex)
                      lngDensityOIDArray(lngDensityArrayIndex) = pFeature2.OID
                      
                      lngIntersectCounter = lngIntersectCounter + 1
                    
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strDensityQueryString = strDensityQueryString & _
                          strPrefix & pDensityFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      
                      booFoundIntersect = True
                    
                    End If
                  
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  
                  If lngDensityArrayIndex > -1 Then
                    strDensityQueryString = Left(strDensityQueryString, Len(strDensityQueryString) - 4)
                    strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                        "  --> Density Query String: " & strDensityQueryString & vbCrLf
                  End If
                End If
                
                If booFoundIntersect And lngIntersectCounter > 1 And _
                        Not MyGeneralOperations.CheckCollectionForKey(pDoneQueryStrings, strQueryLine) Then
                  Debug.Print "  ***** Count = " & Format(lngTotalCount) & "..." & pQueryFilt.WhereClause
                  pDoneQueryStrings.Add True, strQueryLine
                  lngCounter = lngCounter + 1
                  strCheckReport = strCheckReport & Format(lngCounter, "#,##0") & "] " & strQueryLine & vbCrLf
                  
                  ' RUN CLIP OPERATION
                  If IsDimmed(lngDensityOIDArray) Or IsDimmed(lngCoverOIDArray) Then
                    TestFunctions.ClipSetOfPolygons pDensityFClass, lngDensityOIDArray, pCoverFClass, lngCoverOIDArray  ', pMxDoc
                  End If
  
                End If
              End If
              Set pFeature = pFCursor.NextFeature
            Loop
            
            ' NEXT DENSITY
            lngCoverArrayIndex = -1
            lngDensityArrayIndex = -1
            Erase lngCoverOIDArray
            Erase lngDensityOIDArray
            Set pFCursor = pDensityFClass.Search(pQueryFilt, False)
            Set pFeature = pFCursor.NextFeature
            
            Do Until pFeature Is Nothing
            
              Set pPoly1 = pFeature.ShapeCopy
              If Not pPoly1.IsEmpty Then
                Set pBufferPoly = ExpandSmallPolygonByBufferDist(pPoly1)
              Else
                Set pBufferPoly = pPoly1
              End If
              
'              Set pSpFilt.Geometry = pFeature.ShapeCopy
              Set pSpFilt.Geometry = pBufferPoly
              
              lngCoverCount = pCoverFClass.FeatureCount(pSpFilt)
              lngDensityCount = pDensityFClass.FeatureCount(pSpFilt)
              lngTotalCount = lngCoverCount + lngDensityCount
              If lngTotalCount <> 1 Then
                strCoverQueryString = ""
                strDensityQueryString = ""
                strQueryLine = ""
                booFoundIntersect = False
                lngIntersectCounter = 0
                
                strSitePlot = pSiteLookup.Item(strQuadrat)
                strQueryLine = strQuadrat & " (" & strSitePlot & "), " & Format(lngYear, "0") & vbCrLf & _
                    "  --> Species = " & strSpecies & vbCrLf & _
                    "  --> Found " & Format(lngTotalCount, "#,##0") & " overlapping features" & vbCrLf
                    
                If lngCoverCount = 0 And lngDensityCount = 0 Then
                  If lngDensityCount = 0 Then
                    strDensityQueryString = strPrefix & pCoverFClass.OIDFieldName & strSuffix & _
                          " = " & Format(pFeature.OID, "0")
                    strQueryLine = strQueryLine & "  --> EMPTY GEOMETRY!!!  Density Query String: " & strDensityQueryString & vbCrLf
                    booFoundIntersect = True
                  End If
                End If
                
                If lngCoverCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Cover Polygon OID values = "
                  Set pFCursor2 = pCoverFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                      lngCoverArrayIndex = lngCoverArrayIndex + 1
                      ReDim Preserve lngCoverOIDArray(lngCoverArrayIndex)
                      lngCoverOIDArray(lngCoverArrayIndex) = pFeature2.OID
                      
                      lngIntersectCounter = lngIntersectCounter + 1
                    
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strCoverQueryString = strCoverQueryString & _
                          strPrefix & pCoverFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      booFoundIntersect = True
                    End If
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  If lngCoverArrayIndex > -1 Then
                    strCoverQueryString = Left(strCoverQueryString, Len(strCoverQueryString) - 4)
                    strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                        "  --> Cover Query String: " & strCoverQueryString & vbCrLf
                  End If
                End If
                If lngDensityCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Density Polygon OID values = "
                  Set pFCursor2 = pDensityFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                    
                      lngDensityArrayIndex = lngDensityArrayIndex + 1
                      ReDim Preserve lngDensityOIDArray(lngDensityArrayIndex)
                      lngDensityOIDArray(lngDensityArrayIndex) = pFeature2.OID
                      
                      lngIntersectCounter = lngIntersectCounter + 1
                    
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strDensityQueryString = strDensityQueryString & _
                          strPrefix & pDensityFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      booFoundIntersect = True
                    End If
                  
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  If lngDensityArrayIndex > 0 Then
                    strDensityQueryString = Left(strDensityQueryString, Len(strDensityQueryString) - 4)
                    strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                        "  --> Density Query String: " & strDensityQueryString & vbCrLf
                  End If
                End If
                
                If booFoundIntersect And lngIntersectCounter > 1 And _
                      Not MyGeneralOperations.CheckCollectionForKey(pDoneQueryStrings, strQueryLine) Then
                  Debug.Print "  ***** Count = " & Format(lngTotalCount) & "..." & pQueryFilt.WhereClause
                  pDoneQueryStrings.Add True, strQueryLine
                  lngCounter = lngCounter + 1
                  strCheckReport = strCheckReport & Format(lngCounter, "#,##0") & "] " & strQueryLine & vbCrLf
                  
                  ' RUN CLIP OPERATION IF ACTUAL POLYGONS
                  If IsDimmed(lngDensityOIDArray) Or IsDimmed(lngCoverOIDArray) Then
                    TestFunctions.ClipSetOfPolygons pDensityFClass, lngDensityOIDArray, pCoverFClass, lngCoverOIDArray ', pMxDoc
                  End If
  
                End If
              End If
              Set pFeature = pFCursor.NextFeature
            Loop
          End If
        Next lngIndex2
      End If
      
    Next lngIndex
  Next lngYear
  
  ' ERASE ALL EMPTY GEOMETRIES
  pSBar.ShowProgressBar "Pass :  Deleting Empty Geometries...", 0, pCoverFClass.FeatureCount(Nothing) + _
      pDensityFClass.FeatureCount(Nothing), 1, True
  pProg.position = 0
  lngCounter = 0
  
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    If pFeature.Shape.IsEmpty Then
      Debug.Print "...  Deleting Cover Feature #" & Format(pFeature.OID, "0")
      pFeature.DELETE
    End If
    lngCounter = lngCounter + 1
    pProg.Step
    If lngCounter Mod 1000 = 0 Then DoEvents
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    If pFeature.Shape.IsEmpty Then
      Debug.Print "...  Deleting Density Feature #" & Format(pFeature.OID, "0")
      pFeature.DELETE
    End If
    lngCounter = lngCounter + 1
    pProg.Step
    If lngCounter Mod 1000 = 0 Then DoEvents
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Dim pDataObj As New MSForms.DataObject
  pDataObj.SetText strCheckReport
'  pDataObj.PutInClipboard
  
  Dim strPath As String
  strPath = MyGeneralOperations.MakeUniquedBASEName( _
      "D:\arcGIS_stuff\consultation\Margaret_Moore\Intermediate_Analyses\Overlapping_Features.txt")
  MyGeneralOperations.WriteTextFile strPath, strCheckReport, True, False
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSiteLookup = Nothing
  Erase strAllSpecies
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneSpecies = Nothing
  Set pQueryFilt = Nothing
  Set pSpFilt = Nothing
  Set pFCursor2 = Nothing
  Set pFeature2 = Nothing
  Set pDataObj = Nothing



End Sub

Public Function ExpandSmallPolygonByBufferDist(ByVal pPolygon As IPolygon) As IPolygon

  Dim pTransform2D As ITransform2D
  Dim pCentroid As IPoint
  Dim pBuffer As IPolygon
  Dim pClone As IClone
  Dim pModPoly As IPolygon
  
  Set pClone = pPolygon
  Set pModPoly = pClone.Clone
  
  Set pCentroid = MyGeneralOperations.Get_Element_Or_Envelope_Point(pModPoly.Envelope, ENUM_Center_Center)
  
  Dim pBuffCon As IBufferConstruction
  Set pBuffCon = New BufferConstruction
  
  ' SCALE UP
  Set pTransform2D = pModPoly
  With pTransform2D
    .Scale pCentroid, 1000, 1000
  End With
  
  ' BUFFER
  Set pModPoly = pBuffCon.Buffer(pModPoly, 0.75)
  
  ' SCALE BACK DOWN
  Set pTransform2D = pModPoly
  With pTransform2D
    .Scale pCentroid, 0.001, 0.001
  End With
  
  Set ExpandSmallPolygonByBufferDist = pModPoly
  
ClearMemory:
  Set pTransform2D = Nothing
  Set pCentroid = Nothing
  Set pBuffer = Nothing
  Set pClone = Nothing
  Set pModPoly = Nothing
  Set pBuffCon = Nothing
    
End Function

Public Sub CheckForOverlappingPolygons()
  
  ' FIRST GET LISTS OF SPECIES AND PLOTS
  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot)
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityQuadratIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverQuadratIndex As Long
  Dim lngCoverSiteIndex As Long
  Dim lngCoverPlotIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityQuadratIndex = pDensityFClass.FindField("Quadrat")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
  
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverQuadratIndex = pCoverFClass.FindField("Quadrat")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
  
  Dim pSiteLookup As New Collection
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strSpecies As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllSpecies() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneSpecies As New Collection
  
  lngAllSpeciesIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    
    strSite = pFeature.Value(lngDensityQuadratIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
      
      pSiteLookup.Add pFeature.Value(lngDensitySiteIndex) & ", " & pFeature.Value(lngDensityPlotIndex), strSite
      
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    
    strSite = pFeature.Value(lngCoverQuadratIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
      
      pSiteLookup.Add pFeature.Value(lngCoverSiteIndex) & ", " & pFeature.Value(lngCoverPlotIndex), strSite
      
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
'  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  
  ' NOW GO THROUGH EACH UNIQUE COMBINATION OF PLOT, YEAR AND SPECIES.
  ' SELECT ALL CASES WHERE THAT SPECIES EXISTS IN THAT PLOT AND YEAR. BY GOING THROUGH
  Dim pQueryFilt As IQueryFilter
  Set pQueryFilt = New QueryFilter
  Dim lngYear As Long
  Dim strPrefix As String
  Dim strSuffix As String
  Dim strQuadrat As String
  Dim pSpFilt As ISpatialFilter
  Set pSpFilt = New SpatialFilter
  pSpFilt.SpatialRel = esriSpatialRelIntersects
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pCoverFClass, strPrefix, strSuffix)
  Dim lngTotalCount As Long
  Dim strSitePlot As String
  
  Dim strCheckReport As String
      
  pSBar.ShowProgressBar "Working through sites and years...", 0, 18 * (UBound(strAllSites) + 1), 1, True
  pProg.position = 0
  
  lngCounter = 0
  Dim pFCursor2 As IFeatureCursor
  Dim pFeature2 As IFeature
  Dim lngCoverCount As Long
  Dim lngDensityCount As Long
  Dim strCoverQueryString As String
  Dim strDensityQueryString As String
  Dim pDoneQueryStrings As New Collection
  Dim strQueryLine As String
  Dim pPoly1 As IPolygon
  Dim pPoly2 As IPolygon
  Dim dblDist As Double
  Dim booFoundIntersect As Boolean
  Dim lngIntersectCounter As Long
  
  For lngYear = 2002 To 2020
    DoEvents
    For lngIndex = 0 To UBound(strAllSites)
      If lngIndex Mod 10 = 0 Then
        DoEvents
      End If
      pProg.Step
      
      strQuadrat = strAllSites(lngIndex)
      pQueryFilt.WhereClause = strPrefix & "Quadrat" & strSuffix & " = '" & strQuadrat & "' AND " & _
          strPrefix & "Year" & strSuffix & " = '" & Format(lngYear, "0") & "'"
      
      ' MAKE LIST OF SPECIES OBSERVED ON THIS QUADRAT THIS YEAR
      
'      Debug.Print "Quadrat = " & strQuadrat & "; Year = " & Format(lngYear, "0") & vbCrLf & _
'          "  --> Cover Count = " & Format(pCoverFClass.FeatureCount(pQueryFilt), "#,##0") & vbCrLf & _
'          "  --> Density Count = " & Format(pDensityFClass.FeatureCount(pQueryFilt), "#,##0")
                  
      lngAllSpeciesIndex = -1
      Erase strAllSpecies
      Set pDoneSpecies = New Collection
        
      Set pFCursor = pDensityFClass.Search(pQueryFilt, False)
      Set pFeature = pFCursor.NextFeature
      Do Until pFeature Is Nothing
        strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
        If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
        If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
          pDoneSpecies.Add True, strSpecies
          lngAllSpeciesIndex = lngAllSpeciesIndex + 1
          ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
          strAllSpecies(lngAllSpeciesIndex) = strSpecies
        End If
        
        Set pFeature = pFCursor.NextFeature
      Loop
        
      Set pFCursor = pCoverFClass.Search(pQueryFilt, False)
      Set pFeature = pFCursor.NextFeature
      Do Until pFeature Is Nothing
        strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
        If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
        If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
          pDoneSpecies.Add True, strSpecies
          lngAllSpeciesIndex = lngAllSpeciesIndex + 1
          ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
          strAllSpecies(lngAllSpeciesIndex) = strSpecies
        End If
        
        Set pFeature = pFCursor.NextFeature
      Loop
      
      If lngAllSpeciesIndex > -1 Then
        DoEvents
      End If
      
'      Debug.Print "Quadrat = " & strQuadrat & "; Year = " & Format(lngYear, "0") & vbCrLf & _
'          "  --> Cover Count = " & Format(pCoverFClass.FeatureCount(pQueryFilt), "#,##0") & vbCrLf & _
'          "  --> Density Count = " & Format(pDensityFClass.FeatureCount(pQueryFilt), "#,##0") & vbCrLf & _
'          "  --> Found " & Format(lngAllSpeciesIndex + 1, "0") & " unique species"
      
      If lngAllSpeciesIndex > -1 Then
        For lngIndex2 = 0 To lngAllSpeciesIndex
          strSpecies = strAllSpecies(lngIndex2)
          If strSpecies <> "No Polygon Species" Then
            pQueryFilt.WhereClause = strPrefix & "Quadrat" & strSuffix & " = '" & strQuadrat & "' AND " & _
                strPrefix & "Year" & strSuffix & " = '" & Format(lngYear, "0") & "' AND " & _
                strPrefix & "Species" & strSuffix & " = '" & strSpecies & "'"
            pSpFilt.WhereClause = pQueryFilt.WhereClause
            
            ' WORK THOUGH EACH INSTANCE OF A PARTICULAR SPECIES IN THIS QUADRAT IN THIS YEAR.
            ' CHECK TO SEE HOW MANY COVER AND DENSITY FEATURES WITH THIS SPECIES INTERSECT THIS INSTANCE.  SHOULD BE EXACTLY ONE (ITSELF).
            ' FIRST COVER
            Set pFCursor = pCoverFClass.Search(pQueryFilt, False)
            Set pFeature = pFCursor.NextFeature
            Do Until pFeature Is Nothing
            
              Set pPoly1 = pFeature.ShapeCopy
            
              Set pSpFilt.Geometry = pFeature.ShapeCopy
              lngCoverCount = pCoverFClass.FeatureCount(pSpFilt)
              lngDensityCount = pDensityFClass.FeatureCount(pSpFilt)
              lngTotalCount = lngCoverCount + lngDensityCount
              
              If lngTotalCount <> 1 Then
                strCoverQueryString = ""
                strDensityQueryString = ""
                strQueryLine = ""
                booFoundIntersect = False
                lngIntersectCounter = 0
                
                strSitePlot = pSiteLookup.Item(strQuadrat)
                strQueryLine = strQuadrat & " (" & strSitePlot & "), " & Format(lngYear, "0") & vbCrLf & _
                    "  --> Species = " & strSpecies & vbCrLf & _
                    "  --> Found " & Format(lngTotalCount, "#,##0") & " overlapping features" & vbCrLf
                
                If lngCoverCount = 0 And lngDensityCount = 0 Then
                  If lngCoverCount = 0 Then
                    strCoverQueryString = strPrefix & pCoverFClass.OIDFieldName & strSuffix & " = " & Format(pFeature.OID, "0")
                    strQueryLine = strQueryLine & "  --> Cover Query String: " & strCoverQueryString & vbCrLf
                    booFoundIntersect = True  ' JUST TO FORCE EVENT TO BE WRITTEN TO REPORT
                    lngIntersectCounter = 2 ' JUST TO FORCE EVENT TO BE WRITTEN TO REPORT
                  End If
                End If
                
                If lngCoverCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Cover Polygon OID values = "
                  Set pFCursor2 = pCoverFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                      lngIntersectCounter = lngIntersectCounter + 1
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strCoverQueryString = strCoverQueryString & _
                          strPrefix & pCoverFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      
                      booFoundIntersect = True
                    
                    End If
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  
                  strCoverQueryString = Left(strCoverQueryString, Len(strCoverQueryString) - 4)
                  strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                      "  --> Cover Query String: " & strCoverQueryString & vbCrLf
                End If
                
                If lngDensityCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Density Polygon OID values = "
                  Set pFCursor2 = pDensityFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                      lngIntersectCounter = lngIntersectCounter + 1
                    
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strDensityQueryString = strDensityQueryString & _
                          strPrefix & pDensityFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      
                      booFoundIntersect = True
                    
                    End If
                  
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  
                  If lngIntersectCounter > 0 Then
                    strDensityQueryString = Left(strDensityQueryString, Len(strDensityQueryString) - 4)
                    strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                        "  --> Density Query String: " & strDensityQueryString & vbCrLf
                  End If
                End If
                
                If booFoundIntersect And lngIntersectCounter > 1 And _
                        Not MyGeneralOperations.CheckCollectionForKey(pDoneQueryStrings, strQueryLine) Then
                  Debug.Print "  ***** Count = " & Format(lngTotalCount) & "..." & pQueryFilt.WhereClause
                  pDoneQueryStrings.Add True, strQueryLine
                  lngCounter = lngCounter + 1
                  strCheckReport = strCheckReport & Format(lngCounter, "#,##0") & "] " & strQueryLine & vbCrLf
                End If
              End If
              Set pFeature = pFCursor.NextFeature
            Loop
            
            ' NEXT DENSITY
            Set pFCursor = pDensityFClass.Search(pQueryFilt, False)
            Set pFeature = pFCursor.NextFeature
            Do Until pFeature Is Nothing
              Set pSpFilt.Geometry = pFeature.ShapeCopy
              lngCoverCount = pCoverFClass.FeatureCount(pSpFilt)
              lngDensityCount = pDensityFClass.FeatureCount(pSpFilt)
              lngTotalCount = lngCoverCount + lngDensityCount
              If lngTotalCount <> 1 Then
                strCoverQueryString = ""
                strDensityQueryString = ""
                strQueryLine = ""
                booFoundIntersect = False
                lngIntersectCounter = 0
                
                strSitePlot = pSiteLookup.Item(strQuadrat)
                strQueryLine = strQuadrat & " (" & strSitePlot & "), " & Format(lngYear, "0") & vbCrLf & _
                    "  --> Species = " & strSpecies & vbCrLf & _
                    "  --> Found " & Format(lngTotalCount, "#,##0") & " overlapping features" & vbCrLf
                If lngCoverCount = 0 And lngDensityCount = 0 Then
                  If lngDensityCount = 0 Then
                    strDensityQueryString = strPrefix & pCoverFClass.OIDFieldName & strSuffix & " = " & Format(pFeature.OID, "0")
                    strQueryLine = strQueryLine & "  --> Density Query String: " & strDensityQueryString & vbCrLf
                    booFoundIntersect = True
                  End If
                End If
                If lngCoverCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Cover Polygon OID values = "
                  Set pFCursor2 = pCoverFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                      lngIntersectCounter = lngIntersectCounter + 1
                    
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strCoverQueryString = strCoverQueryString & _
                          strPrefix & pCoverFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      booFoundIntersect = True
                    End If
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  strCoverQueryString = Left(strCoverQueryString, Len(strCoverQueryString) - 4)
                  strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                      "  --> Cover Query String: " & strCoverQueryString & vbCrLf
                End If
                If lngDensityCount > 0 Then
                  strQueryLine = strQueryLine & "  --> Density Polygon OID values = "
                  Set pFCursor2 = pDensityFClass.Search(pSpFilt, True)
                  Set pFeature2 = pFCursor2.NextFeature
                  Do Until pFeature2 Is Nothing
                    Set pPoly2 = pFeature2.ShapeCopy
                    dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2))
                    If dblDist = 0 Then
                      lngIntersectCounter = lngIntersectCounter + 1
                    
                      strQueryLine = strQueryLine & Format(pFeature2.OID, "0") & ", "
                      strDensityQueryString = strDensityQueryString & _
                          strPrefix & pDensityFClass.OIDFieldName & strSuffix & " = " & Format(pFeature2.OID, "0") & " OR "
                      booFoundIntersect = True
                    End If
                  
                    Set pFeature2 = pFCursor2.NextFeature
                  Loop
                  If lngIntersectCounter > 0 Then
                    strDensityQueryString = Left(strDensityQueryString, Len(strDensityQueryString) - 4)
                    strQueryLine = Left(strQueryLine, Len(strQueryLine) - 2) & vbCrLf & _
                        "  --> Density Query String: " & strDensityQueryString & vbCrLf
                  End If
                End If
                
                If booFoundIntersect And lngIntersectCounter > 1 And _
                      Not MyGeneralOperations.CheckCollectionForKey(pDoneQueryStrings, strQueryLine) Then
                  Debug.Print "  ***** Count = " & Format(lngTotalCount) & "..." & pQueryFilt.WhereClause
                  pDoneQueryStrings.Add True, strQueryLine
                  lngCounter = lngCounter + 1
                  strCheckReport = strCheckReport & Format(lngCounter, "#,##0") & "] " & strQueryLine & vbCrLf
                End If
              End If
              Set pFeature = pFCursor.NextFeature
            Loop
          End If
        Next lngIndex2
      End If
      
    Next lngIndex
  Next lngYear
    
  Dim pDataObj As New MSForms.DataObject
  pDataObj.SetText strCheckReport
  pDataObj.PutInClipboard
  
  Dim strPath As String
  strPath = MyGeneralOperations.MakeUniquedBASEName( _
      "D:\arcGIS_stuff\consultation\Margaret_Moore\Intermediate_Analyses\Overlapping_Features.txt")
  MyGeneralOperations.WriteTextFile strPath, strCheckReport, True, False
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSiteLookup = Nothing
  Erase strAllSpecies
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneSpecies = Nothing
  Set pQueryFilt = Nothing
  Set pSpFilt = Nothing
  Set pFCursor2 = Nothing
  Set pFeature2 = Nothing
  Set pDataObj = Nothing



End Sub
Public Sub SummarizeSpeciesByCorrectQuadrat()

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, , , , strModifiedRoot, strContainerFolder)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_by_Quadrat.csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverPlotIndex As Long
  Dim lngCoverSiteIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strSpecies As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllSpecies() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneSpecies As New Collection
  
  lngAllSpeciesIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
'    strSite = pFeature.Value(lngDensityPlotIndex)
'    strSite = pFeature.Value(lngDensitySiteIndex) & ": Plot " & pFeature.Value(lngDensityPlotIndex)
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Quadrat " & pFeature.Value(lngDensityPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
'    strSite = pFeature.Value(lngCoverPlotIndex)
'    strSite = pFeature.Value(lngCoverSiteIndex) & ": Plot " & pFeature.Value(lngCoverPlotIndex)
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Quadrat " & pFeature.Value(lngCoverPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  
  Dim pSpeciesIndexes As New Collection
  Dim pSiteIndexes As New Collection
    
  For lngIndex = 0 To UBound(strAllSpecies)
    pSpeciesIndexes.Add lngIndex, strAllSpecies(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(strAllSites)
    pSiteIndexes.Add lngIndex, strAllSites(lngIndex)
  Next lngIndex
  
  Dim lngCounts() As Long
  ReDim lngCounts(UBound(strAllSpecies), UBound(strAllSites))
  
  Dim lngSpeciesIndex As Long
  Dim lngSiteIndex As Long
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
'    strSite = pFeature.Value(lngDensityPlotIndex)
'    strSite = pFeature.Value(lngDensitySiteIndex) & ": Plot " & pFeature.Value(lngDensityPlotIndex)
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Quadrat " & pFeature.Value(lngDensityPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
    lngCounts(lngSpeciesIndex, lngSiteIndex) = lngCounts(lngSpeciesIndex, lngSiteIndex) + 1
        
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
'    strSite = pFeature.Value(lngCoverSiteIndex) & ": Plot " & pFeature.Value(lngCoverPlotIndex)
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Quadrat " & pFeature.Value(lngCoverPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
    lngCounts(lngSpeciesIndex, lngSiteIndex) = lngCounts(lngSpeciesIndex, lngSiteIndex) + 1
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Dim strLine As String
  Dim strReport As String
  
  strLine = """Species Name"","
  For lngIndex = 0 To UBound(strAllSites)
    strLine = strLine & """" & strAllSites(lngIndex) & IIf(lngIndex = UBound(strAllSites), """", """,")
  Next lngIndex
  strReport = strLine & vbCrLf
  
  For lngIndex = 0 To UBound(strAllSpecies)
    strSpecies = Trim(strAllSpecies(lngIndex))
    strLine = """" & IIf(strSpecies = "", "<Null>", strSpecies) & ""","
    For lngIndex2 = 0 To UBound(strAllSites)
      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
    Next lngIndex2
    strReport = strReport & strLine & vbCrLf
  Next lngIndex
  
  
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strReport, ",", vbTab)
'  pDataObj.PutInClipboard
  
  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllSpecies
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneSpecies = Nothing
  Set pSpeciesIndexes = Nothing
  Set pSiteIndexes = Nothing
  Erase lngCounts



End Sub




Public Sub SummarizeSpeciesByPlot()

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, , , , strModifiedRoot, strContainerFolder)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_by_Plot.csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverPlotIndex As Long
  Dim lngCoverSiteIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strSpecies As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllSpecies() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneSpecies As New Collection
  
  lngAllSpeciesIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
'    strSite = pFeature.Value(lngDensityPlotIndex)
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Plot " & pFeature.Value(lngDensityPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
'    strSite = pFeature.Value(lngCoverPlotIndex)
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Plot " & pFeature.Value(lngCoverPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  
  Dim pSpeciesIndexes As New Collection
  Dim pSiteIndexes As New Collection
    
  For lngIndex = 0 To UBound(strAllSpecies)
    pSpeciesIndexes.Add lngIndex, strAllSpecies(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(strAllSites)
    pSiteIndexes.Add lngIndex, strAllSites(lngIndex)
  Next lngIndex
  
  Dim lngCounts() As Long
  ReDim lngCounts(UBound(strAllSpecies), UBound(strAllSites))
  
  Dim lngSpeciesIndex As Long
  Dim lngSiteIndex As Long
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
'    strSite = pFeature.Value(lngDensityPlotIndex)
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Plot " & pFeature.Value(lngDensityPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
    lngCounts(lngSpeciesIndex, lngSiteIndex) = lngCounts(lngSpeciesIndex, lngSiteIndex) + 1
        
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Plot " & pFeature.Value(lngCoverPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
    lngCounts(lngSpeciesIndex, lngSiteIndex) = lngCounts(lngSpeciesIndex, lngSiteIndex) + 1
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Dim strLine As String
  Dim strReport As String
  
  strLine = """Species Name"","
  For lngIndex = 0 To UBound(strAllSites)
    strLine = strLine & """" & strAllSites(lngIndex) & IIf(lngIndex = UBound(strAllSites), """", """,")
  Next lngIndex
  strReport = strLine & vbCrLf
  
  For lngIndex = 0 To UBound(strAllSpecies)
    strSpecies = Trim(strAllSpecies(lngIndex))
    strLine = """" & IIf(strSpecies = "", "<Null>", strSpecies) & ""","
    For lngIndex2 = 0 To UBound(strAllSites)
      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
    Next lngIndex2
    strReport = strReport & strLine & vbCrLf
  Next lngIndex
  
  
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strReport, ",", vbTab)
'  pDataObj.PutInClipboard
  
  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllSpecies
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneSpecies = Nothing
  Set pSpeciesIndexes = Nothing
  Set pSiteIndexes = Nothing
  Erase lngCounts



End Sub


Public Sub SummarizeSpeciesBySite()

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, , , , strModifiedRoot, strContainerFolder)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_by_Site.csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverPlotIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityPlotIndex = pDensityFClass.FindField("Site")
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverPlotIndex = pCoverFClass.FindField("Site")
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strSpecies As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllSpecies() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneSpecies As New Collection
  
  lngAllSpeciesIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
    strSite = pFeature.Value(lngDensityPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
    strSite = pFeature.Value(lngCoverPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  
  Dim pSpeciesIndexes As New Collection
  Dim pSiteIndexes As New Collection
    
  For lngIndex = 0 To UBound(strAllSpecies)
    pSpeciesIndexes.Add lngIndex, strAllSpecies(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(strAllSites)
    pSiteIndexes.Add lngIndex, strAllSites(lngIndex)
  Next lngIndex
  
  Dim lngCounts() As Long
  ReDim lngCounts(UBound(strAllSpecies), UBound(strAllSites))
  
  Dim lngSpeciesIndex As Long
  Dim lngSiteIndex As Long
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
    strSite = pFeature.Value(lngDensityPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
    lngCounts(lngSpeciesIndex, lngSiteIndex) = lngCounts(lngSpeciesIndex, lngSiteIndex) + 1
        
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
    strSite = pFeature.Value(lngCoverPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
    lngCounts(lngSpeciesIndex, lngSiteIndex) = lngCounts(lngSpeciesIndex, lngSiteIndex) + 1
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Dim strLine As String
  Dim strReport As String
  
  strLine = """Species Name"","
  For lngIndex = 0 To UBound(strAllSites)
    strLine = strLine & """" & strAllSites(lngIndex) & IIf(lngIndex = UBound(strAllSites), """", """,")
  Next lngIndex
  strReport = strLine & vbCrLf
  
  For lngIndex = 0 To UBound(strAllSpecies)
    strSpecies = Trim(strAllSpecies(lngIndex))
    strLine = """" & IIf(strSpecies = "", "<Null>", strSpecies) & ""","
    For lngIndex2 = 0 To UBound(strAllSites)
      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
    Next lngIndex2
    strReport = strReport & strLine & vbCrLf
  Next lngIndex
  
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strReport, ",", vbTab)
'  pDataObj.PutInClipboard
  
  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllSpecies
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneSpecies = Nothing
  Set pSpeciesIndexes = Nothing
  Set pSiteIndexes = Nothing
  Erase lngCounts



End Sub

Public Sub SummarizeSpeciesByQuadrat()

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, , , , strModifiedRoot, strContainerFolder)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_by_Quadrat.csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverPlotIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityPlotIndex = pDensityFClass.FindField("Quadrat")
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverPlotIndex = pCoverFClass.FindField("Quadrat")
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strSpecies As String
  Dim strPlot As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllSpecies() As String
  Dim strAllPlots() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneSpecies As New Collection
  
  lngAllSpeciesIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
    strPlot = pFeature.Value(lngDensityPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strPlot) Then
      pDonePlots.Add True, strPlot
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllPlots(lngAllPlotsIndex)
      strAllPlots(lngAllPlotsIndex) = strPlot
      lngQuadrat = CLng(Replace(strPlot, "Q", ""))
      ReDim Preserve lngAllQuadrats(lngAllPlotsIndex)
      lngAllQuadrats(lngAllPlotsIndex) = lngQuadrat
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
      pDoneSpecies.Add True, strSpecies
      lngAllSpeciesIndex = lngAllSpeciesIndex + 1
      ReDim Preserve strAllSpecies(lngAllSpeciesIndex)
      strAllSpecies(lngAllSpeciesIndex) = strSpecies
    End If
    
    strPlot = pFeature.Value(lngCoverPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strPlot) Then
      pDonePlots.Add True, strPlot
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllPlots(lngAllPlotsIndex)
      strAllPlots(lngAllPlotsIndex) = strPlot
      lngQuadrat = CLng(Replace(strPlot, "Q", ""))
      ReDim Preserve lngAllQuadrats(lngAllPlotsIndex)
      lngAllQuadrats(lngAllPlotsIndex) = lngQuadrat
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.LongAscending lngAllQuadrats, 0, UBound(lngAllQuadrats)
  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  
  Dim pSpeciesIndexes As New Collection
  Dim pQuadratIndexes As New Collection
  Dim strQuadrat As String
  Dim strAllQuadrats() As String
  
  
  ReDim strAllQuadrats(UBound(lngAllQuadrats))
  
  For lngIndex = 0 To UBound(strAllSpecies)
    pSpeciesIndexes.Add lngIndex, strAllSpecies(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(lngAllQuadrats)
    strQuadrat = "Q" & Format(lngAllQuadrats(lngIndex), "0")
    pQuadratIndexes.Add lngIndex, strQuadrat
    strAllQuadrats(lngIndex) = strQuadrat
  Next lngIndex
  
  Dim lngCounts() As Long
  ReDim lngCounts(UBound(strAllSpecies), UBound(strAllQuadrats))
  
  Dim lngSpeciesIndex As Long
  Dim lngQuadratIndex As Long
  
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
    strQuadrat = pFeature.Value(lngDensityPlotIndex)
    lngQuadratIndex = pQuadratIndexes.Item(strQuadrat)
    
    lngCounts(lngSpeciesIndex, lngQuadratIndex) = lngCounts(lngSpeciesIndex, lngQuadratIndex) + 1
        
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngCoverSpeciesIndex))
    If strSpecies = "" Then strSpecies = "<-- Species Name Missing -->"
    lngSpeciesIndex = pSpeciesIndexes.Item(strSpecies)
    
    strQuadrat = pFeature.Value(lngCoverPlotIndex)
    lngQuadratIndex = pQuadratIndexes.Item(strQuadrat)
    
    lngCounts(lngSpeciesIndex, lngQuadratIndex) = lngCounts(lngSpeciesIndex, lngQuadratIndex) + 1
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
'  Dim strLine As String
'
'  strLine = ","
'  For lngIndex = 0 To UBound(lngAllQuadrats)
'    strLine = strLine & """" & strAllQuadrats(lngIndex) & IIf(lngIndex = UBound(strAllQuadrats), """", """,")
'  Next lngIndex
'
'  For lngIndex = 0 To UBound(strAllSpecies)
'    strSpecies = Trim(strAllSpecies(lngIndex))
'    strLine = """" & IIf(strSpecies = "", "<Null>", strSpecies) & ""","
'    For lngIndex2 = 0 To UBound(strAllQuadrats)
'      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllQuadrats), "", ",")
'    Next lngIndex2
'    strLine = strLine & vbCrLf
'  Next lngIndex
  
  
  
  Dim strLine As String
  Dim strReport As String
  
  strLine = """Species Name"","
  For lngIndex = 0 To UBound(lngAllQuadrats)
    strLine = strLine & """" & strAllQuadrats(lngIndex) & IIf(lngIndex = UBound(strAllQuadrats), """", """,")
  Next lngIndex
  strReport = strLine & vbCrLf
  
  For lngIndex = 0 To UBound(strAllSpecies)
    strSpecies = Trim(strAllSpecies(lngIndex))
    strLine = """" & IIf(strSpecies = "", "<Null>", strSpecies) & ""","
    For lngIndex2 = 0 To UBound(strAllQuadrats)
      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllQuadrats), "", ",")
    Next lngIndex2
    strReport = strReport & strLine & vbCrLf
  Next lngIndex
  
  
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strReport, ",", vbTab)
'  pDataObj.PutInClipboard
  
  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
'
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strLine, ",", vbTab)
'  pDataObj.PutInClipboard
'
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllSpecies
  Erase strAllPlots
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneSpecies = Nothing
  Set pSpeciesIndexes = Nothing
  Set pQuadratIndexes = Nothing
  Erase strAllQuadrats
  Erase lngCounts

End Sub
Public Sub SummarizeYearByPlotByYear()

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot, , , , strContainerFolder)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_Plots_by_Year.csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensityYearIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverPlotIndex As Long
  Dim lngCoverSiteIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensityYearIndex = pDensityFClass.FindField("Year")
  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverYearIndex = pCoverFClass.FindField("Year")
  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strYear As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllYear() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllYearIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneYear As New Collection
  
  lngAllYearIndex = -1
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngDensityYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
      pDoneYear.Add True, strYear
      lngAllYearIndex = lngAllYearIndex + 1
      ReDim Preserve strAllYear(lngAllYearIndex)
      strAllYear(lngAllYearIndex) = strYear
    End If
    
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Plot " & pFeature.Value(lngDensityPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngCoverYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
      pDoneYear.Add True, strYear
      lngAllYearIndex = lngAllYearIndex + 1
      ReDim Preserve strAllYear(lngAllYearIndex)
      strAllYear(lngAllYearIndex) = strYear
    End If
    
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Plot " & pFeature.Value(lngCoverPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
  QuickSort.StringsAscending strAllYear, 0, UBound(strAllYear)
  
  Dim pYearIndexes As New Collection
  Dim pSiteIndexes As New Collection
    
  For lngIndex = 0 To UBound(strAllYear)
    pYearIndexes.Add lngIndex, strAllYear(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(strAllSites)
    pSiteIndexes.Add lngIndex, strAllSites(lngIndex)
  Next lngIndex
  
  Dim lngCounts() As Long
'  ReDim lngCounts(UBound(strAllYear), UBound(strAllSites))
  ReDim lngCounts(UBound(strAllSites), UBound(strAllYear))
  
  Dim lngYearIndex As Long
  Dim lngSiteIndex As Long
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngDensityYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    lngYearIndex = pYearIndexes.Item(strYear)
    
'    strSite = pFeature.Value(lngDensityPlotIndex)
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Plot " & pFeature.Value(lngDensityPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
'    lngCounts(lngYearIndex, lngSiteIndex) = lngCounts(lngYearIndex, lngSiteIndex) + 1
    lngCounts(lngSiteIndex, lngYearIndex) = lngCounts(lngSiteIndex, lngYearIndex) + 1
        
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngCoverYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    lngYearIndex = pYearIndexes.Item(strYear)
    
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Plot " & pFeature.Value(lngCoverPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
'    lngCounts(lngYearIndex, lngSiteIndex) = lngCounts(lngYearIndex, lngSiteIndex) + 1
    lngCounts(lngSiteIndex, lngYearIndex) = lngCounts(lngSiteIndex, lngYearIndex) + 1
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Dim strLine As String
  Dim strReport As String
  
  strLine = """Plot Name"","
  For lngIndex = 0 To UBound(strAllYear)
    strLine = strLine & """" & strAllYear(lngIndex) & IIf(lngIndex = UBound(strAllYear), """", """,")
  Next lngIndex
  strReport = strLine & vbCrLf
  
  For lngIndex = 0 To UBound(strAllSites)
    strSite = Trim(strAllSites(lngIndex))
    strLine = """" & IIf(strSite = "", "<Null>", strSite) & ""","
    For lngIndex2 = 0 To UBound(strAllYear)
'      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
      strLine = strLine & IIf(lngCounts(lngIndex, lngIndex2) = 0, "", "X") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
    Next lngIndex2
    strReport = strReport & strLine & vbCrLf
  Next lngIndex
  
  
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strReport, ",", vbTab)
'  pDataObj.PutInClipboard
  
  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllYear
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneYear = Nothing
  Set pYearIndexes = Nothing
  Set pSiteIndexes = Nothing
  Erase lngCounts



End Sub

Public Sub SummarizeYearByCorrectQuadratByYear()

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim pCommentColl As Collection
  Set pCommentColl = MakeCollectionOfComments
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot, , , , strContainerFolder)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_Quadrats_by_Year.csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
  
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensityYearIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverPlotIndex As Long
  Dim lngCoverSiteIndex As Long
  
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensityYearIndex = pDensityFClass.FindField("Year")
  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverYearIndex = pCoverFClass.FindField("Year")
  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
  
  Dim strYear As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim strAllYear() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllYearIndex As Long
  Dim lngAllPlotsIndex As Long
  
  Dim pDonePlots As New Collection
  Dim pDoneYear As New Collection
  
  lngAllYearIndex = 0
  ' MANUALLY ADD 2008
  ReDim strAllYear(lngAllYearIndex)
  strAllYear(lngAllYearIndex) = "2008"
  
  lngAllPlotsIndex = -1
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngCounter = 0
    
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngDensityYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
      pDoneYear.Add True, strYear
      lngAllYearIndex = lngAllYearIndex + 1
      ReDim Preserve strAllYear(lngAllYearIndex)
      strAllYear(lngAllYearIndex) = strYear
    End If
    
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Quadrat " & pFeature.Value(lngDensityPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngCoverYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
      pDoneYear.Add True, strYear
      lngAllYearIndex = lngAllYearIndex + 1
      ReDim Preserve strAllYear(lngAllYearIndex)
      strAllYear(lngAllYearIndex) = strYear
    End If
    
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Quadrat " & pFeature.Value(lngCoverPlotIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strSite) Then
      pDonePlots.Add True, strSite
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strSite
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
  QuickSort.StringsAscending strAllYear, 0, UBound(strAllYear)
  
  Dim pYearIndexes As New Collection
  Dim pSiteIndexes As New Collection
    
  For lngIndex = 0 To UBound(strAllYear)
    pYearIndexes.Add lngIndex, strAllYear(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(strAllSites)
    pSiteIndexes.Add lngIndex, strAllSites(lngIndex)
  Next lngIndex
  
  Dim lngCounts() As Long
'  ReDim lngCounts(UBound(strAllYear), UBound(strAllSites))
  ReDim lngCounts(UBound(strAllSites), UBound(strAllYear))
  
  Dim lngYearIndex As Long
  Dim lngSiteIndex As Long
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngDensityYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    lngYearIndex = pYearIndexes.Item(strYear)
    
'    strSite = pFeature.Value(lngDensityPlotIndex)
    strSite = pFeature.Value(lngDensitySiteIndex) & ": Quadrat " & pFeature.Value(lngDensityPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
'    lngCounts(lngYearIndex, lngSiteIndex) = lngCounts(lngYearIndex, lngSiteIndex) + 1
    lngCounts(lngSiteIndex, lngYearIndex) = lngCounts(lngSiteIndex, lngYearIndex) + 1
        
    Set pFeature = pFCursor.NextFeature
  Loop
    
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      DoEvents
    End If
    strYear = Trim(pFeature.Value(lngCoverYearIndex))
    If strYear = "" Then strYear = "<-- Year Name Missing -->"
    lngYearIndex = pYearIndexes.Item(strYear)
    
    strSite = pFeature.Value(lngCoverSiteIndex) & ": Quadrat " & pFeature.Value(lngCoverPlotIndex)
    lngSiteIndex = pSiteIndexes.Item(strSite)
    
'    lngCounts(lngYearIndex, lngSiteIndex) = lngCounts(lngYearIndex, lngSiteIndex) + 1
    lngCounts(lngSiteIndex, lngYearIndex) = lngCounts(lngSiteIndex, lngYearIndex) + 1
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  Dim strLine As String
  Dim strReport As String
  
  strLine = """Quadrat"","
  For lngIndex = 0 To UBound(strAllYear)
    strLine = strLine & """" & strAllYear(lngIndex) & ""","
  Next lngIndex
  strReport = strLine & """Years_Surveyed"",""Proportion"",""Comments""" & vbCrLf
  
  Dim lngYearCounter As Long
  Dim lngCountsByYear() As Long
  Dim strComment As String
  ReDim lngCountsByYear(UBound(strAllYear))
  For lngIndex = 0 To UBound(strAllSites)
    strSite = Trim(strAllSites(lngIndex))
    strComment = pCommentColl.Item(strSite)
    strLine = """" & IIf(strSite = "", "<Null>", strSite) & ""","
    lngYearCounter = 0
    For lngIndex2 = 0 To UBound(strAllYear)
'      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
      strLine = strLine & IIf(lngCounts(lngIndex, lngIndex2) = 0, "", "X") & ","
      If lngCounts(lngIndex, lngIndex2) <> 0 Then
        lngYearCounter = lngYearCounter + 1
        lngCountsByYear(lngIndex2) = lngCountsByYear(lngIndex2) + 1
      End If
    Next lngIndex2
    strReport = strReport & strLine & Format(lngYearCounter, "0") & "," & _
        Format(CDbl(lngYearCounter) / (UBound(strAllYear) + 1) * 100, "0.00") & "%" & ",""" & strComment & """" & vbCrLf
  Next lngIndex
  ' summarize columns
  strReport = strReport & "Sites_Surveyed,"
  For lngIndex = 0 To UBound(lngCountsByYear)
    strReport = strReport & Format(lngCountsByYear(lngIndex), "0") & ","
  Next lngIndex
  strReport = strReport & ",,," & vbCrLf
  strReport = strReport & "Proportion,"
  For lngIndex = 0 To UBound(lngCountsByYear)
    strReport = strReport & Format(CDbl(lngCountsByYear(lngIndex)) / (UBound(strAllSites) + 1) * 100, "0.00") & "%" & ","
  Next lngIndex
  strReport = strReport & ",,," & vbCrLf & """" & pCommentColl.Item("General") & """"
  For lngIndex = 0 To UBound(lngCountsByYear)
    strReport = strReport & ","
  Next lngIndex
  strReport = strReport & ",,,"
  
  
  
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText Replace(strReport, ",", vbTab)
'  pDataObj.PutInClipboard
  
  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllYear
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDonePlots = Nothing
  Set pDoneYear = Nothing
  Set pYearIndexes = Nothing
  Set pSiteIndexes = Nothing
  Erase lngCounts



End Sub

Public Function MakeCollectionOfComments() As Collection

  Dim pCommentColl As New Collection
  pCommentColl.Add "Comments", "Quadrat"
  pCommentColl.Add "", "Big Fill: Quadrat 11999"
  pCommentColl.Add "", "Big Fill: Quadrat 12000"
  pCommentColl.Add "Relocated in 2004.  Located original 1912 quadrat corners as precisely as possible and all quadrat corners are rebar.  .", "Big Fill: Quadrat 16 / 30716"
  pCommentColl.Add "Relocated in 2004.  Located original 1912 quadrat corners as precisely as possible and all quadrat corners are rebar.  Not measured after 2016 due to large ponderosa pine tree fall next to it.", "Big Fill: Quadrat 18 / 30718"
  pCommentColl.Add "Powerline corridor.  Ponderosa pine trees cut in corridor.  Not measured after 2016 due to disturbance.", "Big Fill: Quadrat 30711"
  pCommentColl.Add "Powerline corridor.  Ponderosa pine trees cut in corridor.  Not measured after 2016 due to disturbance.", "Big Fill: Quadrat 30712"
  pCommentColl.Add "", "Big Fill: Quadrat 30713"
  pCommentColl.Add "", "Big Fill: Quadrat 30714"
  pCommentColl.Add "", "Big Fill: Quadrat 30715"
  pCommentColl.Add "", "Big Fill: Quadrat 30717"
  pCommentColl.Add "", "Big Fill: Quadrat 30719"
  pCommentColl.Add "", "Big Fill: Quadrat 30720"
  pCommentColl.Add "", "Black Springs: Quadrat 26345"
  pCommentColl.Add "", "Black Springs: Quadrat 26346"
  pCommentColl.Add "", "Black Springs: Quadrat 26347"
  pCommentColl.Add "", "Black Springs: Quadrat 26348"
  pCommentColl.Add "", "Black Springs: Quadrat 30741"
  pCommentColl.Add "", "Black Springs: Quadrat 30742"
  pCommentColl.Add "", "Black Springs: Quadrat 30743"
  pCommentColl.Add "", "Black Springs: Quadrat 30744"
  pCommentColl.Add "Relocated in 2004.  Truck ruts on or near quadrat.  Not measured after 2016 due to disturbance. ", "Black Springs: Quadrat 30745"
  pCommentColl.Add "Relocated in 2004.  Social trail goes through the middle of it.   Not measured after 2016 due to disturbance. ", "Black Springs: Quadrat 30746"
  pCommentColl.Add "Relocated in 2004.  Has all original 1912 angle iron corners. ", "Black Springs: Quadrat 30747"
  pCommentColl.Add "", "Black Springs: Quadrat 30748"
  pCommentColl.Add "", "Black Springs: Quadrat 30749"
  pCommentColl.Add "Located on edge of ditch directly next to I-17S.  Much BROTEC near it.  Not measured after 2016 due to location.", "Black Springs: Quadrat 30750"
  pCommentColl.Add "Relocated in 2004.  Located original 1912 quadrat corners as precisely as possible and all quadrat corners are rebar. ", "Black Springs: Quadrat 46"
  pCommentColl.Add "Relocated in 2016.  Four quadrats with angle iron and tags within old exclosure, same age as nearby Kendrick Panels. ", "FS 9009H Road - Old Exclosure: Quadrat 494"
  pCommentColl.Add "Relocated in 2016.  Four quadrats with angle iron and tags within old exclosure, same age as nearby Kendrick Panels. ", "FS 9009H Road - Old Exclosure: Quadrat 498"
  pCommentColl.Add "", "Fry Park: Quadrat 30731"
  pCommentColl.Add "", "Fry Park: Quadrat 30732"
  pCommentColl.Add "", "Fry Park: Quadrat 30733"
  pCommentColl.Add "", "Fry Park: Quadrat 30734"
  pCommentColl.Add "", "Fry Park: Quadrat 30735"
  pCommentColl.Add "Within relatively large clay flat.  Original 1912 quadrat angle iron exits.  Few perennial plants, mostly native annual plants.  Not measured after 2016 due to time.  ", "Fry Park: Quadrat 30736"
  pCommentColl.Add "Relocated in 2004.  Two angle irons from original 1912 quadrat were located.  Livestock trail goes through edge of it.   Not measured after 2016 due to disturbance. ", "Fry Park: Quadrat 30737"
  pCommentColl.Add "Original 1912 quadrat located under large ponderosa pine tree where livestock gather beneath it.  Much POAPRA on it.  Not measured after 2017 due to time.  ", "Fry Park: Quadrat 30738"
  pCommentColl.Add "", "Fry Park: Quadrat 30739"
  pCommentColl.Add "", "Fry Park: Quadrat 30740"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S1A: Quadrat 21114"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S1A: Quadrat 21174"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S1B: Quadrat 21262"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S1B: Quadrat 21269"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S2A: Quadrat 22126"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S2A: Quadrat 22156"
  pCommentColl.Add "Located in 2006.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S2B: Quadrat 22241"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S2B: Quadrat 22244"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S3A: Quadrat 23155"
  pCommentColl.Add "Located in 2007.  Established a 1-m2 quadrats within the original 1914 not spaded 1.5 m x 3.0 m understory quadrat. These 1914 quadrats are nested within the larger silviculture plots (S1A, etc.) on Ft Valley Exp. Forest.", "Ft Valley - COC-S3A: Quadrat 23159"
  pCommentColl.Add "Relocated in 2004.  Has all original 1912 angle iron corners.  Reese Tank also listed as Rees Tank in historical literature. ", "Reese Tank: Quadrat 10 / 30710"
  pCommentColl.Add "", "Reese Tank: Quadrat 30701"
  pCommentColl.Add "", "Reese Tank: Quadrat 30702"
  pCommentColl.Add "", "Reese Tank: Quadrat 30703"
  pCommentColl.Add "Relocated in 2004.  All angle iron from original 1912 quadrat were located.  Quadrat close to ditch of FS Road 418.  Not measured after 2016 due to some disturbance and time. ", "Reese Tank: Quadrat 30704"
  pCommentColl.Add "", "Reese Tank: Quadrat 30705"
  pCommentColl.Add "", "Reese Tank: Quadrat 30706"
  pCommentColl.Add "", "Reese Tank: Quadrat 30707"
  pCommentColl.Add "", "Reese Tank: Quadrat 30709"
  pCommentColl.Add "Relocated in 2006.  All angle iron from original 1912 quadrat were located.  Quadrat is under old slash pile from 1980s.  Not measured after 2016 due to disturbance.   ", "Reese Tank: Quadrat 8 / 30708"
  pCommentColl.Add "", "Rogers Lake: Quadrat 26339"
  pCommentColl.Add "Relocated in 2004.  All angle iron from original 1912 quadrat were located.  Quadrat is island of 3-4 ponderosa pine trees at logging roads intersection.  Not measured after 2016 due to disturbance.   ", "Rogers Lake: Quadrat 26340"
  pCommentColl.Add "", "Rogers Lake: Quadrat 26369"
  pCommentColl.Add "", "Rogers Lake: Quadrat 26370"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30721"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30722"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30723"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30724"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30725"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30726"
  pCommentColl.Add "Relocated in 2004.  At least one angle iron located from original 1912.  Quadrat is west of original exclosure in heavily used camp site.  Not measured after 2016 due to disturbance.   ", "Rogers Lake: Quadrat 30727"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30728"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30729"
  pCommentColl.Add "", "Rogers Lake: Quadrat 30730"
  pCommentColl.Add "Relocated in 2006.  All angle iron in place from 1927 on these 'Wild Bill' or 'Cooperrider-Cassidy Study'", "Wild Bill - Dispersed Quadrats: Quadrat 122"
  pCommentColl.Add "Relocated in 2006.  All angle iron in place from 1927 on these 'Wild Bill' or 'Cooperrider-Cassidy Study'", "Wild Bill - Dispersed Quadrats: Quadrat 123"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 124"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 125"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 29003"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 29004"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 29016"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 29017"
  pCommentColl.Add "", "Wild Bill - Dispersed Quadrats: Quadrat 29025"
  pCommentColl.Add "", "Wild Bill - Government Knolls: Quadrat 101"
  pCommentColl.Add "", "Wild Bill - Government Knolls: Quadrat 102"
  pCommentColl.Add "", "Wild Bill - Government Knolls: Quadrat 103"
  pCommentColl.Add "Relocated in 2007.  ", "Wild Bill - Government Knolls: Quadrat 104"
  pCommentColl.Add "", "Wild Bill - Government Knolls: Quadrat 105"
  pCommentColl.Add "Relocated in 2016.  On edge of very old slash pile, but not under it.", "Wild Bill - Government Knolls: Quadrat 106"
  pCommentColl.Add "", "Wild Bill - Kendrick Panel: Quadrat 107"
  pCommentColl.Add "", "Wild Bill - Kendrick Panel: Quadrat 108"
  pCommentColl.Add "", "Wild Bill - Kendrick Panel: Quadrat 109"
  pCommentColl.Add "", "Wild Bill - Kendrick Panel: Quadrat 110"
  pCommentColl.Add "", "Wild Bill - SI Panel: Quadrat 115"
  pCommentColl.Add "", "Wild Bill - SI Panel: Quadrat 120"
  pCommentColl.Add "", "Wild Bill - SI Panel: Quadrat 121"
  pCommentColl.Add "Relocated in 2006.  All angle iron from original 1912 quadrat were located.  ", "Wild Bill - Wild Bill Panel: Quadrat 114"
  pCommentColl.Add "", "Wild Bill - Wild Bill Panel: Quadrat 119"
  pCommentColl.Add "Relocated in 2006. Did not burn in Pumpkin Fire of May 2000, but burned in Boundary Fire (Kendrick Mountain) in June 2017, which resulted in complete overstory tree mortality and forest floor consumption.  ", "Wild Bill - Willaha: Quadrat 1"
  pCommentColl.Add "Relocated in 2006. Did not burn in Pumpkin Fire of May 2000, but burned in Boundary Fire (Kendrick Mountain) in June 2017, which resulted in complete overstory tree mortality and forest floor consumption.  ", "Wild Bill - Willaha: Quadrat 2"
  pCommentColl.Add "Relocated in 2006. Partial burn of overstory trees in Pumpkin Fire of May 2000, but did not burn in Boundary Fire (Kendrick Mountain) in June 2017.", "Wild Bill - Willaha: Quadrat 3"
  pCommentColl.Add "Relocated in 2006. Tree overstory burn in Pumpkin Fire of May 2000, which resulted in tree downfall (jackpot) over the 1-m2 quadrat. Quadrat did not burn in Boundary Fire of 2017.  Did not measure in 2016-2018 due to time. Prescribed fire in fall 2018 resulted in complete consumption of downfall.  Started remeasuring again in 2019 and 2020...  ", "Wild Bill - Willaha: Quadrat 4"
  pCommentColl.Add "NOTE:  All angle iron (or galvanized pipe) relocated from original studies 1912 (Hill Range Exclosures), 1914 (Fort Valley Exp Forest Sample Plots), or 1927 (Cooperrider and Cassidy Range Study). ", "General"

  Set MakeCollectionOfComments = pCommentColl

'  Dim pWS As IFeatureWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Set pWSFact = New TextFileWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Data_to_include_in_publication\", 0)
'
'  Dim pTable As ITable
'  Set pTable = pWS.OpenTable("Source_for_Quadrats_by_Year_Comments.csv")
'
'  Debug.Print pTable.RowCount(Nothing)
'  Dim pFields As IFields
'  Dim lngIndex As Long
'  Dim pField As iField
'  Set pFields = pTable.Fields
''  For lngIndex = 0 To pFields.FieldCount - 1
''    Set pField = pFields.Field(lngIndex)
''    Debug.Print CStr(lngIndex) & "] " & pField.Name
''  Next lngIndex
'
'  Dim pCursor As ICursor
'  Dim pRow As IRow
'  Set pCursor = pTable.Search(Nothing, False)
'  Set pRow = pCursor.NextRow
'  Dim strQuadrat As String
'  Dim strComment As String
'
'  Dim strReport As String
'  strReport = "  dim pCommentColl as new collection" & vbCrLf
'  Do Until pRow Is Nothing
'    If Not IsNull(pRow.Value(0)) Then
'      strQuadrat = pRow.Value(0)
'      If strQuadrat <> "Sites_Surveyed" And strQuadrat <> "Proportion" Then
'        If IsNull(pRow.Value(22)) Then
'          strComment = ""
'        Else
'          strComment = pRow.Value(22)
'        End If
'        If Left(strQuadrat, 5) = "NOTE:" Then
'          strComment = strQuadrat
'          strQuadrat = "General"
'        End If
'        strReport = strReport & "  pCommentColl.Add " & """" & strComment & """, """ & strQuadrat & """" & vbCrLf
'
'        Debug.Print pRow.Value(0) & "..." & pRow.Value(22)
'      End If
'    End If
'
'    Set pRow = pCursor.NextRow
'  Loop
'
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.SetText strReport
'  pDataObj.PutInClipboard
'  Set pDataObj = Nothing
'
'  Debug.Print strReport
End Function


Public Sub TestReturnColYearsSurveyedByQuadrat()
  Dim lngYear1 As Long
  Dim lngYear2 As Long
  Dim pReturn As Collection
  
  lngYear1 = 2002
  lngYear2 = 2020
  Set pReturn = ReturnCollectionOfYearsSurveyedByQuadrat(lngYear1, lngYear2)
End Sub

Public Function ReturnCollectionOfYearsSurveyedByQuadrat(lngYear1 As Long, lngYear2 As Long) As Collection

  Dim lngStart As Long
  lngStart = GetTickCount
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot, , , , strContainerFolder)
  
  Dim pFiles As esriSystem.IStringArray
  Set pFiles = MyGeneralOperations.ReturnFilesFromNestedFolders2(strCombinePath, ".shp")
  
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\Summarize_Plots_by_Year.csv")
  Dim lngIndex As Long
  Dim strPath As String
  Dim strFile As String
  Dim strYear As String
  Dim strQuadrat As String
  Dim strSplit() As String
  
  Dim strAllYear() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllYearIndex As Long
  Dim lngAllPlotsIndex As Long
  Dim strPairArray() As String
  Dim lngPairArrayIndex As Long
  
  Dim pDonePairs As New Collection
  Dim pDonePlots As New Collection
  Dim pDoneYear As New Collection
  
  lngAllYearIndex = -1
  lngAllPlotsIndex = -1
  lngPairArrayIndex = -1
  
  lngCount = pFiles.Count
  lngCounter = 0
  
  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0
  
  For lngIndex = 0 To pFiles.Count - 1
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    
    strPath = aml_func_mod.ReturnFilename2(pFiles.Element(lngIndex))
    strSplit = Split(strPath, "_")
    strQuadrat = strSplit(0)
    strYear = strSplit(1)
    
    If strYear = "" Then
      MsgBox "No Year!"
      strYear = "<-- Year Name Missing -->"
    End If
    
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePairs, strQuadrat & "_" & strYear) Then
      pDonePairs.Add True, strQuadrat & "_" & strYear
      lngPairArrayIndex = lngPairArrayIndex + 1
      ReDim Preserve strPairArray(1, lngPairArrayIndex)
      strPairArray(0, lngPairArrayIndex) = strQuadrat
      strPairArray(1, lngPairArrayIndex) = strYear
    End If
    
    
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
      pDoneYear.Add True, strYear
      lngAllYearIndex = lngAllYearIndex + 1
      ReDim Preserve strAllYear(lngAllYearIndex)
      strAllYear(lngAllYearIndex) = strYear
    End If
    
    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strQuadrat) Then
      pDonePlots.Add True, strQuadrat
      lngAllPlotsIndex = lngAllPlotsIndex + 1
      ReDim Preserve strAllSites(lngAllPlotsIndex)
      strAllSites(lngAllPlotsIndex) = strQuadrat
    End If
  Next lngIndex
    
  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
  QuickSort.StringsAscending strAllYear, 0, UBound(strAllYear)
  
  Dim pYearIndexes As New Collection
  Dim pQuadratIndexes As New Collection
    
  For lngIndex = 0 To UBound(strAllYear)
    pYearIndexes.Add lngIndex, strAllYear(lngIndex)
  Next lngIndex
  For lngIndex = 0 To UBound(strAllSites)
    pQuadratIndexes.Add lngIndex, strAllSites(lngIndex)
  Next lngIndex
  
  Dim lngYearIndex As Long
  Dim lngQuadratIndex As Long
  Dim lngIndex2 As Long
  
  pSBar.ShowProgressBar "Second Pass...", 0, lngPairArrayIndex, 1, True
  pProg.position = 0
  
  Dim pReturnColl As Collection
  Dim pSubColl As Collection
  
  Set pReturnColl = New Collection
  
  For lngIndex = 0 To lngPairArrayIndex
    strQuadrat = strPairArray(0, lngIndex)
    strYear = strPairArray(1, lngIndex)
    
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    
    lngYearIndex = pYearIndexes.Item(strYear)
    lngQuadratIndex = pQuadratIndexes.Item(strQuadrat)
    
    If MyGeneralOperations.CheckCollectionForKey(pReturnColl, strQuadrat) Then
      Set pSubColl = pReturnColl.Item(strQuadrat)
      pReturnColl.Remove strQuadrat
    Else
      Set pSubColl = New Collection
      For lngIndex2 = lngYear1 To lngYear2
        pSubColl.Add False, Format(lngIndex2, "0")
      Next lngIndex2
    End If
    
    If MyGeneralOperations.CheckCollectionForKey(pSubColl, strYear) Then
      If pSubColl.Item(strYear) = False Then
        pSubColl.Remove strYear
        pSubColl.Add True, strYear
      End If
    End If
    pReturnColl.Add pSubColl, strQuadrat
    
  Next lngIndex
  
  
  
  
'
'
'
'
'
'
'
'
'
'
'
'
'  Dim pWS As IFeatureWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Set pWSFact = New FileGDBWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)
'
'  Dim pDensityFClass As IFeatureClass
'  Dim pCoverFClass As IFeatureClass
'  Dim pFCursor As IFeatureCursor
'  Dim pFeature As IFeature
'  Dim lngDensityYearIndex As Long
'  Dim lngDensityPlotIndex As Long
'  Dim lngDensityQuadratIndex As Long
'  Dim lngCoverYearIndex As Long
'  Dim lngCoverPlotIndex As Long
'  Dim lngCoverQuadratIndex As Long
'
'  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
'  lngDensityYearIndex = pDensityFClass.FindField("Year")
'  lngDensityPlotIndex = pDensityFClass.FindField("Plot")
'  lngDensityQuadratIndex = pDensityFClass.FindField("Quadrat")
'  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
'  lngCoverYearIndex = pCoverFClass.FindField("Year")
'  lngCoverPlotIndex = pCoverFClass.FindField("Plot")
'  lngCoverQuadratIndex = pCoverFClass.FindField("Quadrat")
'
'  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
'
'  Dim strYear As String
'  Dim strQuadrat As String
'  Dim lngQuadrat As Long
'  Dim lngIndex As Long
'  Dim lngIndex2 As Long
'
'  Dim strAllYear() As String
'  Dim strAllSites() As String
'  Dim lngAllQuadrats() As Long
'  Dim lngAllYearIndex As Long
'  Dim lngAllPlotsIndex As Long
'
'  Dim pDonePlots As New Collection
'  Dim pDoneYear As New Collection
'
'  lngAllYearIndex = -1
'  lngAllPlotsIndex = -1
'
'  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
'  pProg.position = 0
'
'  lngCounter = 0
'
'  Set pFCursor = pDensityFClass.Search(Nothing, False)
'  Set pFeature = pFCursor.NextFeature
'  Do Until pFeature Is Nothing
'    pProg.Step
'    lngCounter = lngCounter + 1
'    If lngCounter Mod 100 = 0 Then
'      DoEvents
'    End If
'    strYear = Trim(pFeature.Value(lngDensityYearIndex))
'    If strYear = "" Then strYear = "<-- Year Name Missing -->"
'    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
'      pDoneYear.Add True, strYear
'      lngAllYearIndex = lngAllYearIndex + 1
'      ReDim Preserve strAllYear(lngAllYearIndex)
'      strAllYear(lngAllYearIndex) = strYear
'    End If
'
'    strQuadrat = pFeature.Value(lngDensityQuadratIndex) ' & ": Plot " & pFeature.Value(lngDensityPlotIndex)
'    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strQuadrat) Then
'      pDonePlots.Add True, strQuadrat
'      lngAllPlotsIndex = lngAllPlotsIndex + 1
'      ReDim Preserve strAllSites(lngAllPlotsIndex)
'      strAllSites(lngAllPlotsIndex) = strQuadrat
'    End If
'
'    Set pFeature = pFCursor.NextFeature
'  Loop
'
'  Set pFCursor = pCoverFClass.Search(Nothing, False)
'  Set pFeature = pFCursor.NextFeature
'  Do Until pFeature Is Nothing
'    pProg.Step
'    lngCounter = lngCounter + 1
'    If lngCounter Mod 100 = 0 Then
'      DoEvents
'    End If
'    strYear = Trim(pFeature.Value(lngCoverYearIndex))
'    If strYear = "" Then strYear = "<-- Year Name Missing -->"
'    If Not MyGeneralOperations.CheckCollectionForKey(pDoneYear, strYear) Then
'      pDoneYear.Add True, strYear
'      lngAllYearIndex = lngAllYearIndex + 1
'      ReDim Preserve strAllYear(lngAllYearIndex)
'      strAllYear(lngAllYearIndex) = strYear
'    End If
'
'    strQuadrat = pFeature.Value(lngCoverQuadratIndex) ' & ": Plot " & pFeature.Value(lngCoverPlotIndex)
'    If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strQuadrat) Then
'      pDonePlots.Add True, strQuadrat
'      lngAllPlotsIndex = lngAllPlotsIndex + 1
'      ReDim Preserve strAllSites(lngAllPlotsIndex)
'      strAllSites(lngAllPlotsIndex) = strQuadrat
'    End If
'
'    Set pFeature = pFCursor.NextFeature
'  Loop
'
'  QuickSort.StringsAscending strAllSites, 0, UBound(strAllSites)
'  QuickSort.StringsAscending strAllYear, 0, UBound(strAllYear)
'
'  Dim pYearIndexes As New Collection
'  Dim pQuadratIndexes As New Collection
'
'  For lngIndex = 0 To UBound(strAllYear)
'    pYearIndexes.Add lngIndex, strAllYear(lngIndex)
'  Next lngIndex
'  For lngIndex = 0 To UBound(strAllSites)
'    pQuadratIndexes.Add lngIndex, strAllSites(lngIndex)
'  Next lngIndex
'
'  Dim lngCounts() As Long
''  ReDim lngCounts(UBound(strAllYear), UBound(strAllSites))
'  ReDim lngCounts(UBound(strAllSites), UBound(strAllYear))
'
'  Dim lngYearIndex As Long
'  Dim lngQuadratIndex As Long
'
'  pSBar.ShowProgressBar "Second Pass...", 0, lngCount, 1, True
'  pProg.position = 0
'
'  Dim pReturnColl As Collection
'  Dim pSubColl As Collection
'
'  Set pReturnColl = New Collection
'
'  Set pFCursor = pDensityFClass.Search(Nothing, False)
'  Set pFeature = pFCursor.NextFeature
'  Do Until pFeature Is Nothing
'    pProg.Step
'    lngCounter = lngCounter + 1
'    If lngCounter Mod 100 = 0 Then
'      DoEvents
'    End If
'    strYear = Trim(pFeature.Value(lngDensityYearIndex))
'    If strYear = "" Then strYear = "<-- Year Name Missing -->"
'    lngYearIndex = pYearIndexes.Item(strYear)
'
''    strQuadrat = pFeature.Value(lngDensityPlotIndex)
'    strQuadrat = pFeature.Value(lngDensityQuadratIndex) ' & ": Plot " & pFeature.Value(lngDensityPlotIndex)
'    lngQuadratIndex = pQuadratIndexes.Item(strQuadrat)
'
'    If MyGeneralOperations.CheckCollectionForKey(pReturnColl, strQuadrat) Then
'      Set pSubColl = pReturnColl.Item(strQuadrat)
'      pReturnColl.Remove strQuadrat
'    Else
'      Set pSubColl = New Collection
'      For lngIndex = lngYear1 To lngYear2
'        pSubColl.Add False, Format(lngIndex, "0")
'      Next lngIndex
'    End If
'
'    If MyGeneralOperations.CheckCollectionForKey(pSubColl, strYear) Then
'      If pSubColl.Item(strYear) = False Then
'        pSubColl.Remove strYear
'        pSubColl.Add True, strYear
'      End If
'    End If
'    pReturnColl.Add pSubColl, strQuadrat
'
''    lngCounts(lngYearIndex, lngQuadratIndex) = lngCounts(lngYearIndex, lngQuadratIndex) + 1
''    lngCounts(lngQuadratIndex, lngYearIndex) = lngCounts(lngQuadratIndex, lngYearIndex) + 1
'
'    Set pFeature = pFCursor.NextFeature
'  Loop
'
'  Set pFCursor = pCoverFClass.Search(Nothing, False)
'  Set pFeature = pFCursor.NextFeature
'  Do Until pFeature Is Nothing
'    pProg.Step
'    lngCounter = lngCounter + 1
'    If lngCounter Mod 1000 = 0 Then
'      DoEvents
'    End If
'    strYear = Trim(pFeature.Value(lngCoverYearIndex))
'    If strYear = "" Then strYear = "<-- Year Name Missing -->"
'    lngYearIndex = pYearIndexes.Item(strYear)
'
'    strQuadrat = pFeature.Value(lngCoverQuadratIndex) ' & ": Plot " & pFeature.Value(lngCoverPlotIndex)
'    lngQuadratIndex = pQuadratIndexes.Item(strQuadrat)
'
'
'    If MyGeneralOperations.CheckCollectionForKey(pReturnColl, strQuadrat) Then
'      Set pSubColl = pReturnColl.Item(strQuadrat)
'      pReturnColl.Remove strQuadrat
'    Else
'      Set pSubColl = New Collection
'      For lngIndex = lngYear1 To lngYear2
'        pSubColl.Add False, Format(lngIndex, "0")
'      Next lngIndex
'    End If
'
'    If MyGeneralOperations.CheckCollectionForKey(pSubColl, strYear) Then
'      If pSubColl.Item(strYear) = False Then
'        pSubColl.Remove strYear
'        pSubColl.Add True, strYear
'      End If
'    End If
'    pReturnColl.Add pSubColl, strQuadrat
'
''    lngCounts(lngYearIndex, lngQuadratIndex) = lngCounts(lngYearIndex, lngQuadratIndex) + 1
''    lngCounts(lngQuadratIndex, lngYearIndex) = lngCounts(lngQuadratIndex, lngYearIndex) + 1
'
'    Set pFeature = pFCursor.NextFeature
'  Loop
  
'  Dim strLine As String
'  Dim strReport As String
'
'  strLine = """Plot Name"","
'  For lngIndex = 0 To UBound(strAllYear)
'    strLine = strLine & """" & strAllYear(lngIndex) & IIf(lngIndex = UBound(strAllYear), """", """,")
'  Next lngIndex
'  strReport = strLine & vbCrLf
'
'  For lngIndex = 0 To UBound(strAllSites)
'    strQuadrat = Trim(strAllSites(lngIndex))
'    strLine = """" & IIf(strQuadrat = "", "<Null>", strQuadrat) & ""","
'    For lngIndex2 = 0 To UBound(strAllYear)
''      strLine = strLine & Format(lngCounts(lngIndex, lngIndex2), "0") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
'      strLine = strLine & IIf(lngCounts(lngIndex, lngIndex2) = 0, "", "X") & IIf(lngIndex2 = UBound(strAllSites), "", ",")
'    Next lngIndex2
'    strReport = strReport & strLine & vbCrLf
'  Next lngIndex
'
'
''  Dim pDataObj As New MSForms.DataObject
''  pDataObj.SetText Replace(strReport, ",", vbTab)
''  pDataObj.PutInClipboard
'
'  MyGeneralOperations.WriteTextFile strExportPath, strReport, True, False
  
  Set ReturnCollectionOfYearsSurveyedByQuadrat = pReturnColl
  
  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pFiles = Nothing
  Erase strSplit
  Erase strAllYear
  Erase strAllSites
  Erase lngAllQuadrats
  Erase strPairArray
  Set pDonePairs = Nothing
  Set pDonePlots = Nothing
  Set pDoneYear = Nothing
  Set pYearIndexes = Nothing
  Set pQuadratIndexes = Nothing
  Set pReturnColl = Nothing
  Set pSubColl = Nothing






End Function





