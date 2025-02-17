Attribute VB_Name = "H_WB_Analysis"
Option Explicit

Public Sub RunAsBatch()

  Dim lngTimeStart As Long
  lngTimeStart = GetTickCount

  OrganizeData
  ReviseShapefiles
  ReviseShapefiles
  ConvertPointShapefiles
  More_Margaret_Functions.RepairOverlappingPolygons
  AddEmptyFeaturesAndFeatureClasses (False)
  More_Margaret_Functions.RecreateSubsetsOfConvertedDatasets
  AddEmptyFeaturesAndFeatureClassesToCleaned
  ShiftFinishedShapefilesToCoordinateSystem
  ExportFinalDataset
  More_Margaret_Functions.SummarizeSpeciesBySite
  More_Margaret_Functions.SummarizeSpeciesByCorrectQuadrat
  More_Margaret_Functions.SummarizeYearByCorrectQuadratByYear
  CreateFinalTables
  Margaret_Functions_3.ExportSubsetsOfSpeciesShapefiles True, False, False

  Debug.Print "============================"
  Debug.Print "Batch Process Complete:"
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngTimeStart)
End Sub

Public Sub CreateFinalTables()

  Debug.Print "-----------------------------------"

  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFinalTable As ITable
  Dim pAddTable As ITable

  Dim strNewSource As String
  strNewSource = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Source_Files_March_2018\HillPlotQC_Laughlin.xlsx"

  Dim strOrigRoot As String
  Dim strModifiedRoot As String
  Dim strShiftRoot As String
  Dim strFinalFolder As String
  Dim strExportBase As String
  Dim strSetFolder As String
  Call DeclareWorkspaces(strOrigRoot, , strShiftRoot, strExportBase, strModifiedRoot, strSetFolder, , strFinalFolder)

  Dim strFolder As String
  Dim lngIndex As Long

  Dim strPlotLocNames() As String
  Dim pPlotLocColl As Collection

  Dim strPlotDataNames() As String
  Dim pPlotDataColl As Collection

  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim varSites() As Variant
  Dim varSiteSpecifics() As Variant
  Dim varArray() As Variant

  Call ReturnQuadratVegSoilData(pPlotDataColl, strPlotDataNames)

  Call ReturnQuadratCoordsAndNames(pPlotLocColl, strPlotLocNames)
  For lngIndex = 0 To UBound(strPlotLocNames)
  Next lngIndex

  Dim pPlotToQuadratColl As Collection
  Dim pQuadratToPlotColl As Collection
  Set pQuadratColl = FillQuadratNameColl_Rev(strQuadratNames, pPlotToQuadratColl, pQuadratToPlotColl, varSites, varSiteSpecifics)

  Dim pVegDataAndElevations As Collection
  Dim strVegDataElevNames() As String
  Call ReturnVegDataElevAndNames(pVegDataAndElevations, strVegDataElevNames, pPlotLocColl)

  Dim pFullQuadratData As Collection
  Set pFullQuadratData = ReturnQuadratData(pPlotLocColl)

  For lngIndex = 0 To UBound(strVegDataElevNames)
    Debug.Print CStr(lngIndex) & "] " & strVegDataElevNames(lngIndex)
    varArray = pVegDataAndElevations.Item(strVegDataElevNames(lngIndex))
  Next lngIndex

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Dim pNewWSFact As IWorkspaceFactory
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pSrcWS As IFeatureWorkspace
  Dim pNewWS As IFeatureWorkspace
  Dim pSrcCoverFClass As IFeatureClass
  Dim pSrcDensFClass As IFeatureClass
  Dim pTopoOp As ITopologicalOperator4
  Dim lngQuadIndex As Long

  Dim strQuadrat As String
  Dim strDestFolder As String
  Dim strItem() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strFileHeader As String
  Dim dblCentroidX As Double
  Dim dblCentroidY As Double

  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace

  Dim strFClassName As String
  Dim strNameSplit() As String
  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Set pNewWSFact = New FileGDBWorkspaceFactory
  Dim pWStoUpdate As IWorkspace
  Set pWStoUpdate = pNewWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
  Dim pEnumDataset As IEnumDataset
  Dim pUpdateDataset As IDataset
  Dim pFClass As IFeatureClass

  Dim strNewAncillaryFolder As String
  strNewAncillaryFolder = strFinalFolder & "\Ancillary_Data_CSVs"

  MyGeneralOperations.CreateNestedFoldersByPath strNewAncillaryFolder

  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strFinalFolder & "\Ancillary_Data_GDB")
  Dim pWS2 As IWorkspace2
  Dim pFeatWS As IFeatureWorkspace
  Set pWS2 = pNewWS
  Set pFeatWS = pWS2

  Dim pFCursor As IFeatureCursor
  Dim pFBuffer As IFeatureBuffer

  Dim pNewFClass As IFeatureClass
  Dim pFields As esriSystem.IVariantArray
  Dim lngSiteIndex As Long
  Dim lngAspectIndex As Long
  Dim lngSlopeIndex As Long
  Dim lngCanopyCoverIndex As Long
  Dim lngBasalAreaIndex As Long
  Dim lngAltBasalAreaIndex As Long
  Dim lngSoilIndex As Long
  Dim lngElevIndex As Long
  Dim lngNorthingIndex As Long
  Dim lngEastingIndex As Long
  Dim lngSpeciesIndex As Long
  Dim lngAbbrevIndex As Long
  Dim lngTypeIndex As Long
  Dim lngLatitudeIndex As Long
  Dim lngLongitudeIndex As Long

  Dim strSpeciesData() As String
  Dim lngSpeciesArrayIndex As Long
  Dim pDoneSpecies As New Collection
  Dim pNewTable As ITable
  Dim pTestWS As IFeatureWorkspace
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngDensityYearIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngDensitySpeciesIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverPlotIndex As Long
  Dim lngCoverSiteIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngYearIndex As Long
  Dim lngCommentIndex As Long
  Dim strSpecies As String
  Dim strAbbrev As String
  Dim strType As String
  Dim strSplit() As String
  Dim pRowBuffer As IRowBuffer
  Dim pFeature As IFeature

  Dim pVegComment As Collection
  Set pVegComment = New Collection
  pVegComment.Add "Previously known as Arenaria fendleri; Mat forming perennial forb", "Eremogone fendleri"
  pVegComment.Add "Previously known as Blepharoneuron tricholepis", "Muhlenbergia tricholepis"
  pVegComment.Add "Previously known as Lotus wrightii", "Acmispon wrightii"
  pVegComment.Add "Previously known as Bahia dissecta", "Amauriopsis dissecta"
  pVegComment.Add "Previously known as Chamaesyce fendleri", "Euphorbia fendleri"
  pVegComment.Add "Previously known as Chamaesyce revulata", "Euphorbia revoluta"
  pVegComment.Add "Previously known as Chamaesyce serpyllifolia", "Euphorbia serpyllifolia"
  pVegComment.Add "Previously known as Chamaesyce; Could not identify to species level", "Euphorbia sp."
  pVegComment.Add "Previously known as Chenopodium graveolens", "Dysphania graveolens"
  pVegComment.Add "Previously known as Machaeranthera canescens", "Dieteria canescens"
  pVegComment.Add "Previously known as Machaeranthera gracilis", "Xanthisma gracile"
  pVegComment.Add "Previously known as Noccaea montana", "Noccaea fendleri"
  pVegComment.Add "Mat forming perennial forb", "Antennaria parvifolia"
  pVegComment.Add "Mat forming perennial forb", "Antennaria rosulata"
  pVegComment.Add "Mat forming perennial forb", "Arenaria lanuginosa"
  pVegComment.Add "Could not identify to species level", "Allium sp."
  pVegComment.Add "Could not identify to species level", "Asclepias sp."
  pVegComment.Add "Could not identify to species level", "Astragalus sp."
  pVegComment.Add "Could not identify to species level", "Castilleja sp."
  pVegComment.Add "Could not identify to species level", "Cirsium sp."
  pVegComment.Add "Could not identify to species level", "Erigeron sp."
  pVegComment.Add "Could not identify to species level", "Geranium sp."
  pVegComment.Add "Could not identify to species level", "Linum sp."
  pVegComment.Add "Could not identify to species level", "Lupinus sp."
  pVegComment.Add "Could not identify to species level", "Oxalis sp."
  pVegComment.Add "Could not identify to species level", "Phlox sp."
  pVegComment.Add "Could not identify to species level", "Physaria sp."
  pVegComment.Add "Could not identify to species level", "Potentilla sp."
  pVegComment.Add "Could not identify to species level", "Senecio sp."
  pVegComment.Add "Could not identify to species level", "Solidago sp."
  pVegComment.Add "Could not identify to species level", "Vicia sp."
  pVegComment.Add "Unknown perennial graminoid", "Unknown graminoid"
  Dim strComment As String

  Dim strFinalQuadratList() As String
  Dim pDoneQuadratColl As New Collection
  Dim lngQuadratArrayIndex As Long
  lngQuadratArrayIndex = -1

  lngSpeciesArrayIndex = -1
  If Not pWS2.NameExists(esriDTTable, "Vegetation_Species") Then
    Set pFields = New esriSystem.varArray
    pFields.Add MyGeneralOperations.CreateNewField("Species", esriFieldTypeString, , 255)
    pFields.Add MyGeneralOperations.CreateNewField("Abbreviation", esriFieldTypeString, , 18)
    pFields.Add MyGeneralOperations.CreateNewField("Type", esriFieldTypeString, , 7)
    pFields.Add MyGeneralOperations.CreateNewField("Notes", esriFieldTypeString, , 75)

    Set pNewTable = MyGeneralOperations.CreateGDBTable(pNewWS, "Plant_Species_List", pFields)
    lngSpeciesIndex = pNewTable.FindField("Species")
    lngAbbrevIndex = pNewTable.FindField("Abbreviation")
    lngTypeIndex = pNewTable.FindField("Type")
    lngCommentIndex = pNewTable.FindField("Notes")
    Set pCursor = pNewTable.Insert(True)
    Set pRowBuffer = pNewTable.CreateRowBuffer

    strPurpose = "List of all species observed in all quadrats over all years."

    Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewTable, strAbstract, strPurpose)

    Set pTestWS = pNewWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
    Set pDensityFClass = pTestWS.OpenFeatureClass("Density_All")
    lngDensityYearIndex = pDensityFClass.FindField("Year")
    lngDensityPlotIndex = pDensityFClass.FindField("Quadrat")
    lngDensitySiteIndex = pDensityFClass.FindField("Site")
    lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
    Set pCoverFClass = pTestWS.OpenFeatureClass("Cover_All")
    lngCoverYearIndex = pCoverFClass.FindField("Year")
    lngCoverPlotIndex = pCoverFClass.FindField("Quadrat")
    lngCoverSiteIndex = pCoverFClass.FindField("Site")
    lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
    Dim lngCount As Long
    Dim lngCounter As Long

    lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
    pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
    pProg.position = 0

    lngCounter = 0

    strType = "Density"
    Set pFCursor = pDensityFClass.Search(Nothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      pProg.Step
      lngCounter = lngCounter + 1
      If lngCounter Mod 100 = 0 Then
        DoEvents
      End If
      strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
      If MyGeneralOperations.CheckCollectionForKey(pVegComment, strSpecies) Then
        strComment = pVegComment.Item(strSpecies)
      Else
        strComment = ""
      End If
      strPlot = Trim(pFeature.Value(lngDensityPlotIndex))
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneQuadratColl, strPlot) Then
        lngQuadratArrayIndex = lngQuadratArrayIndex + 1
        ReDim Preserve strFinalQuadratList(lngQuadratArrayIndex)
        strFinalQuadratList(lngQuadratArrayIndex) = strPlot
        pDoneQuadratColl.Add True, strPlot
      End If
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then ' And _
            InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
        pDoneSpecies.Add True, strSpecies
        lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
        ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
        strSplit = Split(strSpecies, " ")
        If InStr(1, strSpecies, "No Cover", vbTextCompare) > 0 Then
          strAbbrev = "No Cover Species"
        ElseIf InStr(1, strSpecies, "No Density", vbTextCompare) > 0 Then
          strAbbrev = "No Density Species"
        ElseIf StrComp(Trim(strSplit(1)), "Sp.", vbTextCompare) = 0 Then
          strAbbrev = UCase(Left(strSplit(0), 3)) & "SPP"
        Else
          strAbbrev = UCase(Left(strSplit(0), 3) & Left(strSplit(1), 3))
        End If
        strSpeciesData(0, lngSpeciesArrayIndex) = strSpecies
        strSpeciesData(1, lngSpeciesArrayIndex) = strAbbrev
        strSpeciesData(2, lngSpeciesArrayIndex) = strType
        strSpeciesData(3, lngSpeciesArrayIndex) = strComment
      End If

      Set pFeature = pFCursor.NextFeature
    Loop
    strType = "Cover"
    Set pFCursor = pCoverFClass.Search(Nothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      pProg.Step
      lngCounter = lngCounter + 1
      If lngCounter Mod 100 = 0 Then
        DoEvents
      End If
      strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
      If MyGeneralOperations.CheckCollectionForKey(pVegComment, strSpecies) Then
        strComment = pVegComment.Item(strSpecies)
      Else
        strComment = ""
      End If
      strPlot = Trim(pFeature.Value(lngCoverPlotIndex))
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneQuadratColl, strPlot) Then
        lngQuadratArrayIndex = lngQuadratArrayIndex + 1
        ReDim Preserve strFinalQuadratList(lngQuadratArrayIndex)
        strFinalQuadratList(lngQuadratArrayIndex) = strPlot
        pDoneQuadratColl.Add True, strPlot
      End If
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then ' And _
            InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
        pDoneSpecies.Add True, strSpecies
        lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
        ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
        strSplit = Split(strSpecies, " ")
        If InStr(1, strSpecies, "No Cover", vbTextCompare) > 0 Then
          strAbbrev = "No Cover Species"
        ElseIf InStr(1, strSpecies, "No Density", vbTextCompare) > 0 Then
          strAbbrev = "No Density Species"
        ElseIf StrComp(Trim(strSplit(1)), "Sp.", vbTextCompare) = 0 Then
          strAbbrev = UCase(Left(strSplit(0), 3)) & "SPP"
        Else
          strAbbrev = UCase(Left(strSplit(0), 3) & Left(strSplit(1), 3))
        End If
        strSpeciesData(0, lngSpeciesArrayIndex) = strSpecies
        strSpeciesData(1, lngSpeciesArrayIndex) = strAbbrev
        strSpeciesData(2, lngSpeciesArrayIndex) = strType
        strSpeciesData(3, lngSpeciesArrayIndex) = strComment
      End If

      Set pFeature = pFCursor.NextFeature
    Loop

    QuickSort.StringAscending_TwoDimensional strSpeciesData, 0, lngSpeciesArrayIndex, 0, 3
    For lngIndex = 0 To lngSpeciesArrayIndex
      pRowBuffer.Value(lngSpeciesIndex) = strSpeciesData(0, lngIndex)
      pRowBuffer.Value(lngAbbrevIndex) = strSpeciesData(1, lngIndex)
      pRowBuffer.Value(lngTypeIndex) = strSpeciesData(2, lngIndex)
      pRowBuffer.Value(lngCommentIndex) = strSpeciesData(3, lngIndex)
      pCursor.InsertRow pRowBuffer
    Next lngIndex

    pRowBuffer.Value(lngSpeciesIndex) = "Footnote on Type: 'Cover' = polygon feature for perennial graminoids " & _
        "or mat forming forb, while 'Density' = point feature for annual and perennial forbs, annual graminoids or tree seedlings"
    pRowBuffer.Value(lngAbbrevIndex) = ""
    pRowBuffer.Value(lngTypeIndex) = ""
    pRowBuffer.Value(lngCommentIndex) = ""
    pCursor.InsertRow pRowBuffer

    pCursor.Flush

    ProduceTabularAreaPerSpeciesTable pCoverFClass, lngCoverYearIndex, lngCoverPlotIndex, lngCoverSiteIndex, _
        lngCoverSpeciesIndex, pApp, pSBar, pProg, pCoverFClass.FeatureCount(Nothing), strNewAncillaryFolder, pNewWS, _
        strAbstract, pMxDoc

  Else
    Set pNewTable = pFeatWS.OpenTable("Vegetation_Species")
  End If
  MyGeneralOperations.ExportToCSV pNewTable, strNewAncillaryFolder & "\Plant_Species_List.csv", _
        True, False, False, True, , , True

  Dim pReplacements As Collection
  Dim pQuadratReplacements As Collection
  Set pReplacements = ReturnSubstituteNamesColl(pQuadratReplacements)

  If Not pWS2.NameExists(esriDTFeatureClass, "Overstory_Data_and_Quadrat_Locations") Then
    Set pFields = New esriSystem.varArray
    pFields.Add MyGeneralOperations.CreateNewField("Site", esriFieldTypeString, , 35)
    pFields.Add MyGeneralOperations.CreateNewField("Quadrat", esriFieldTypeString, , 15)
    pFields.Add MyGeneralOperations.CreateNewField("Tree_Perc_Canopy_Cover", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Tree_Basal_Area_per_Ha", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Parent_Material_Class", esriFieldTypeString, , 5)
    pFields.Add MyGeneralOperations.CreateNewField("Elevation_m", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Aspect", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Slope_Percent", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Easting_NAD_1983_UTM_12", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Northing_NAD_1983_UTM_12", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Year_Canopy_Measured", esriFieldTypeString, , 5)
    pFields.Add MyGeneralOperations.CreateNewField("Comment", esriFieldTypeString, , 55)

    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass2(pNewWS, "Overstory_Data_and_Quadrat_Locations", esriFTSimple, _
          pSpRef, esriGeometryPoint, pFields, , , , False, ENUM_FileGDB, , 102)

    strPurpose = "Summary of overstory plot vegetation, soil and topographic characteristics in 20m x 20m " & _
        "plots surrounding each quadrat."
    Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewFClass, strAbstract, strPurpose)

    lngSiteIndex = pNewFClass.FindField("Site")
    lngQuadIndex = pNewFClass.FindField("Quadrat")
    lngAspectIndex = pNewFClass.FindField("Aspect")
    lngSlopeIndex = pNewFClass.FindField("Slope_Percent")
    lngCanopyCoverIndex = pNewFClass.FindField("Tree_Perc_Canopy_Cover")
    lngBasalAreaIndex = pNewFClass.FindField("Tree_Basal_Area_per_Ha")
    lngSoilIndex = pNewFClass.FindField("Parent_Material_Class")
    lngElevIndex = pNewFClass.FindField("Elevation_m")
    lngNorthingIndex = pNewFClass.FindField("Northing_NAD_1983_UTM_12")
    lngEastingIndex = pNewFClass.FindField("Easting_NAD_1983_UTM_12")
    lngYearIndex = pNewFClass.FindField("Year_Canopy_Measured")
    lngCommentIndex = pNewFClass.FindField("Comment")

    Set pFCursor = pNewFClass.Insert(True)
    Set pFBuffer = pNewFClass.CreateFeatureBuffer

    For lngIndex = 0 To UBound(strVegDataElevNames)
      varArray = pVegDataAndElevations.Item(strVegDataElevNames(lngIndex))
      Set pFBuffer.Shape = varArray(7)
      pFBuffer.Value(lngSiteIndex) = PickCorrectName(CStr(varArray(0)), pReplacements) ' varArray(0)
      pFBuffer.Value(lngQuadIndex) = PickCorrectName(strVegDataElevNames(lngIndex), pQuadratReplacements)
      pFBuffer.Value(lngElevIndex) = varArray(1)
      pFBuffer.Value(lngAspectIndex) = varArray(2)
      pFBuffer.Value(lngSlopeIndex) = varArray(3)
      pFBuffer.Value(lngCanopyCoverIndex) = varArray(4)
      pFBuffer.Value(lngBasalAreaIndex) = varArray(5)
      pFBuffer.Value(lngSoilIndex) = varArray(6)
      pFBuffer.Value(lngNorthingIndex) = varArray(8)
      pFBuffer.Value(lngEastingIndex) = varArray(9)
      pFBuffer.Value(lngYearIndex) = varArray(10)
      pFBuffer.Value(lngCommentIndex) = ""
      pFCursor.InsertFeature pFBuffer

      If pFBuffer.Value(lngQuadIndex) = "30750" Then
        Set pFBuffer.Shape = New Point
        pFBuffer.Value(lngSiteIndex) = "FS 9009H"
        pFBuffer.Value(lngQuadIndex) = "494"
        pFBuffer.Value(lngElevIndex) = Null
        pFBuffer.Value(lngAspectIndex) = Null
        pFBuffer.Value(lngSlopeIndex) = Null
        pFBuffer.Value(lngCanopyCoverIndex) = Null
        pFBuffer.Value(lngBasalAreaIndex) = Null
        pFBuffer.Value(lngSoilIndex) = "Bas"
        pFBuffer.Value(lngNorthingIndex) = Null
        pFBuffer.Value(lngEastingIndex) = Null
        pFBuffer.Value(lngYearIndex) = Null
        pFBuffer.Value(lngCommentIndex) = "Located in 2016; 20x20 m tree plot not established yet"
        pFCursor.InsertFeature pFBuffer

        Set pFBuffer.Shape = New Point
        pFBuffer.Value(lngSiteIndex) = "FS 9009H"
        pFBuffer.Value(lngQuadIndex) = "498"
        pFBuffer.Value(lngElevIndex) = Null
        pFBuffer.Value(lngAspectIndex) = Null
        pFBuffer.Value(lngSlopeIndex) = Null
        pFBuffer.Value(lngCanopyCoverIndex) = Null
        pFBuffer.Value(lngBasalAreaIndex) = Null
        pFBuffer.Value(lngSoilIndex) = "Bas"
        pFBuffer.Value(lngNorthingIndex) = Null
        pFBuffer.Value(lngEastingIndex) = Null
        pFBuffer.Value(lngYearIndex) = Null
        pFBuffer.Value(lngCommentIndex) = "Located in 2016; 20x20 m tree plot not established yet"
        pFCursor.InsertFeature pFBuffer
      ElseIf pFBuffer.Value(lngQuadIndex) = "105" Then
        Set pFBuffer.Shape = New Point
        pFBuffer.Value(lngSiteIndex) = "Wild Bill"
        pFBuffer.Value(lngQuadIndex) = "106"
        pFBuffer.Value(lngElevIndex) = Null
        pFBuffer.Value(lngAspectIndex) = Null
        pFBuffer.Value(lngSlopeIndex) = Null
        pFBuffer.Value(lngCanopyCoverIndex) = Null
        pFBuffer.Value(lngBasalAreaIndex) = Null
        pFBuffer.Value(lngSoilIndex) = "Bas"
        pFBuffer.Value(lngNorthingIndex) = Null
        pFBuffer.Value(lngEastingIndex) = Null
        pFBuffer.Value(lngYearIndex) = Null
        pFBuffer.Value(lngCommentIndex) = "Located in 2016; 20x20 m tree plot not established yet"
        pFCursor.InsertFeature pFBuffer
      End If

    Next lngIndex
    Debug.Print "Done with Plot Data FClass..."
    pFCursor.Flush
  Else
    Set pNewFClass = pFeatWS.OpenFeatureClass("Overstory_Data_and_Quadrat_Locations")
  End If

  MyGeneralOperations.ExportToCSV pNewFClass, strNewAncillaryFolder & "\Overstory_Data_and_Quadrat_Locations.csv", _
        True, False, False, True, , , True

  Dim pNAD83 As ISpatialReference
  Set pNAD83 = MyGeneralOperations.CreateSpatialReferenceNAD83

        strExclosure, strNote, strComment, strComment2, dblElev), strPlot

  Dim strQuadratSortList() As String
  ReDim strQuadratSortList(1, UBound(strFinalQuadratList))
  For lngIndex = 0 To UBound(strFinalQuadratList)
    strQuadrat = strFinalQuadratList(lngIndex)
    varArray = pFullQuadratData.Item(strQuadrat)
    strSite = varArray(2)
    strQuadratSortList(0, lngIndex) = strSite & "_" & varArray(3) & "_" & strQuadrat
    strQuadratSortList(1, lngIndex) = strQuadrat
  Next lngIndex

  Dim lngAKAIndex As Long
  Dim lngExclosureIndex As Long
  Dim lngNoteIndex As Long
  Dim lngComment2Index As Long
  Dim lngSiteSpecificIndex As Long
  Dim strQuadratComment As String

  QuickSort.StringAscending_TwoDimensional strQuadratSortList, 0, UBound(strQuadratSortList, 2), 0, 1

  If Not pWS2.NameExists(esriDTFeatureClass, "Quadrat_Locations_and_Data") Then
    Set pFields = New esriSystem.varArray
    pFields.Add MyGeneralOperations.CreateNewField("Site", esriFieldTypeString, , 35)
    pFields.Add MyGeneralOperations.CreateNewField("Site_Specific", esriFieldTypeString, , 80)
    pFields.Add MyGeneralOperations.CreateNewField("Quadrat", esriFieldTypeString, , 15)
    pFields.Add MyGeneralOperations.CreateNewField("AKA", esriFieldTypeString, , 15)
    pFields.Add MyGeneralOperations.CreateNewField("Easting_NAD_1983_UTM_12", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Northing_NAD_1983_UTM_12", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Latitude_NAD_1983", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Longitude_NAD_1983", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Elevation_m", esriFieldTypeDouble)
    pFields.Add MyGeneralOperations.CreateNewField("Exclosure", esriFieldTypeString, , 15)
    pFields.Add MyGeneralOperations.CreateNewField("Note", esriFieldTypeString, , 150)
    pFields.Add MyGeneralOperations.CreateNewField("Comment", esriFieldTypeString, , 150)

    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass2(pNewWS, "Quadrat_Locations_and_Data", esriFTSimple, _
          pNAD83, esriGeometryPoint, pFields, , , , False, ENUM_FileGDB, , 102)

    strPurpose = "Quadrat locations, in UTM Zone 12 and Geographic coordinates (both in North American Datum of 1983)."
    Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewFClass, strAbstract, strPurpose)

    lngSiteIndex = pNewFClass.FindField("Site")
    lngSiteSpecificIndex = pNewFClass.FindField("Site_Specific")
    lngQuadIndex = pNewFClass.FindField("Quadrat")
    lngAKAIndex = pNewFClass.FindField("AKA")
    lngEastingIndex = pNewFClass.FindField("Easting_NAD_1983_UTM_12")
    lngNorthingIndex = pNewFClass.FindField("Northing_NAD_1983_UTM_12")
    lngLatitudeIndex = pNewFClass.FindField("Latitude_NAD_1983")
    lngLongitudeIndex = pNewFClass.FindField("Longitude_NAD_1983")
    lngElevIndex = pNewFClass.FindField("Elevation_m")
    lngExclosureIndex = pNewFClass.FindField("Exclosure")
    lngNoteIndex = pNewFClass.FindField("Note")
    lngCommentIndex = pNewFClass.FindField("Comment")

    Set pFCursor = pNewFClass.Insert(True)
    Set pFBuffer = pNewFClass.CreateFeatureBuffer
    Dim pGeoPoint As IPoint

    For lngIndex = 0 To UBound(strQuadratSortList, 2)

            strExclosure, strNote, strComment, strComment2, dblElev), strPlot
      strQuadrat = strQuadratSortList(1, lngIndex)
      varArray = pFullQuadratData.Item(strQuadrat)
      Set pGeoPoint = varArray(10)
      Set pFBuffer.Shape = pGeoPoint

      pFBuffer.Value(lngSiteIndex) = varArray(2)
      pFBuffer.Value(lngSiteSpecificIndex) = varArray(3)
      pFBuffer.Value(lngQuadIndex) = strQuadrat
      pFBuffer.Value(lngAKAIndex) = varArray(4)
      pFBuffer.Value(lngEastingIndex) = varArray(0)
      pFBuffer.Value(lngNorthingIndex) = varArray(1)
      pFBuffer.Value(lngLatitudeIndex) = pGeoPoint.Y
      pFBuffer.Value(lngLongitudeIndex) = pGeoPoint.x
      pFBuffer.Value(lngElevIndex) = varArray(9)
      pFBuffer.Value(lngExclosureIndex) = varArray(5)
      pFBuffer.Value(lngNoteIndex) = FixQuadratComment(Trim(varArray(6)))
      strQuadratComment = FixQuadratComment(Trim(Replace(varArray(8), """", "")))
      pFBuffer.Value(lngCommentIndex) = strQuadratComment

      pFCursor.InsertFeature pFBuffer
    Next lngIndex
    Debug.Print "Done with Quadrat Locations FClass..."
    pFCursor.Flush
  Else
    Set pNewFClass = pFeatWS.OpenFeatureClass("Quadrat_Locations_and_Data")
  End If

  MyGeneralOperations.ExportToCSV pNewFClass, strNewAncillaryFolder & "\Quadrat_Locations_and_Data.csv", _
        True, False, False, True, , , True

  MyGeneralOperations.ExportToCSV_SpecialCases pCoverFClass, strNewAncillaryFolder & "\Cover_Species_Tabular_Version.csv", _
        True, False, False, True, Array("species", "Site", "Quadrat", "Year"), pApp, True, True, pPlotLocColl
  MyGeneralOperations.ExportToCSV_SpecialCases pDensityFClass, strNewAncillaryFolder & "\Density_Species_Tabular_Version.csv", _
        True, False, False, True, Array("species", "Site", "Quadrat", "Year"), pApp, True, True, pPlotLocColl

  FileCopy strSetFolder & "\Summarize_Quadrats_by_Year.csv", strNewAncillaryFolder & "\Summarize_Quadrats_by_Year.csv"
  FileCopy strSetFolder & "\Summarize_by_Site.csv", strNewAncillaryFolder & "\Summarize_by_Site.csv"

  Debug.Print "Done..."

  pSBar.HideProgressBar
  pProg.position = 0
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pFinalTable = Nothing
  Set pAddTable = Nothing
  Erase strPlotLocNames
  Set pPlotLocColl = Nothing
  Erase strPlotDataNames
  Set pPlotDataColl = Nothing
  Erase strQuadratNames
  Set pQuadratColl = Nothing
  Erase varSites
  Erase varSiteSpecifics
  Erase varArray
  Set pPlotToQuadratColl = Nothing
  Set pQuadratToPlotColl = Nothing
  Set pVegDataAndElevations = Nothing
  Erase strVegDataElevNames
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pNewWSFact = Nothing
  Set pSrcWS = Nothing
  Set pNewWS = Nothing
  Set pSrcCoverFClass = Nothing
  Set pSrcDensFClass = Nothing
  Set pTopoOp = Nothing
  Erase strItem
  Set pDatasetEnum = Nothing
  Set pWS = Nothing
  Erase strNameSplit
  Set pWStoUpdate = Nothing
  Set pEnumDataset = Nothing
  Set pUpdateDataset = Nothing
  Set pFClass = Nothing
  Set pWS2 = Nothing
  Set pNewFClass = Nothing
  Set pFields = Nothing
  Set pFCursor = Nothing
  Set pFBuffer = Nothing

End Sub

Public Function FixQuadratComment(strText As String) As String

  Dim strReturn As String
  strReturn = Replace(strText, ", tag?", "")
  strReturn = Replace(strReturn, "; tag?", "")
  strReturn = Replace(strReturn, "; rebar?", "")
  strReturn = Replace(strReturn, ":  Email from June 9, 2020", "")

  FixQuadratComment = strReturn

End Function

Public Function PickCorrectName(strName As String, pSubstitutes As Collection) As String

  If MyGeneralOperations.CheckCollectionForKey(pSubstitutes, strName) Then
    PickCorrectName = pSubstitutes.Item(strName)
  Else
    PickCorrectName = strName
  End If

End Function

Public Function ReturnSubstituteNamesColl(pQuadratReplacements As Collection) As Collection

  Set pQuadratReplacements = New Collection
  pQuadratReplacements.Add "16 / 30716", "16"
  pQuadratReplacements.Add "18 / 30718", "18"
  pQuadratReplacements.Add "10 / 30710", "10"
  pQuadratReplacements.Add "8 / 30708", "8"

  Dim pReturn As New Collection
  pReturn.Add "Fort Valley COC-S1A", "S1A"
  pReturn.Add "Fort Valley COC-S1B", "S1B"
  pReturn.Add "Fort Valley COC-S2A", "S2A"
  pReturn.Add "Fort Valley COC-S2B", "S2B"
  pReturn.Add "Fort Valley COC-S3A", "S3A"

  Set ReturnSubstituteNamesColl = pReturn

End Function

Public Sub ProduceTabularAreaPerSpeciesTable(pCoverFClass As IFeatureClass, lngCoverYearIndex As Long, _
    lngCoverPlotIndex As Long, lngCoverSiteIndex As Long, lngCoverSpeciesIndex As Long, pApp As IApplication, _
    pSBar As IStatusBar, pProg As IStepProgressor, lngCount As Long, strNewAncillaryFolder As String, _
    pNewWS As IWorkspace, strAbstract As String, pMxDoc As IMxDocument)

  Dim lngCounter As Long

  pSBar.ShowProgressBar "Cover Basal Area Pass 1 of 2...", 0, lngCount, 1, True
  pProg.position = 0

  lngCounter = 0
  Dim strSortArray() As String
  Dim lngSortCounter As Long
  Dim pDoneColl As New Collection
  lngSortCounter = -1
  Dim strSpecies As String
  Dim strSite As String
  Dim strYear As String
  Dim strQuadrat As String

  Dim strPrefix As String
  Dim strSuffix As String
  Dim strKey As String
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pCoverFClass, strPrefix, strSuffix)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature

  Dim strType As String
  strType = "Cover"
  Dim strBaseQuery As String

  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSite = Trim(pFeature.Value(lngCoverSiteIndex))
    strQuadrat = Trim(pFeature.Value(lngCoverPlotIndex))
    strKey = strSite & ":" & strQuadrat
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strKey) Then
      pDoneColl.Add True, strKey
      lngSortCounter = lngSortCounter + 1
      ReDim Preserve strSortArray(2, lngSortCounter)
      strSortArray(0, lngSortCounter) = strSite
      strSortArray(1, lngSortCounter) = strQuadrat
      strSortArray(2, lngSortCounter) = strPrefix & "Site" & strSuffix & " = '" & strSite & "' AND " & _
          strPrefix & "Quadrat" & strSuffix & " = '" & strQuadrat & "'"
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  QuickSort.StringAscending_TwoDimensional strSortArray, 0, lngSortCounter, 0, 2
  Dim lngIndex As Long

  Dim strYears() As String
  Dim strSpeciesArray() As String
  Dim lngYearIndex As Long
  Dim lngSpeciesIndex As Long
  Dim strBaseYearQuery As String
  Dim strBaseSpeciesQuery As String
  Dim dblTotalArea As Double
  Dim lngObservationCount As Long

  Dim pFields As esriSystem.IVariantArray
  Set pFields = New esriSystem.varArray
  pFields.Add MyGeneralOperations.CreateNewField("Site", esriFieldTypeString, , 35)
  pFields.Add MyGeneralOperations.CreateNewField("Quadrat", esriFieldTypeString, , 15)
  pFields.Add MyGeneralOperations.CreateNewField("Year", esriFieldTypeString, , 5)
  pFields.Add MyGeneralOperations.CreateNewField("Type", esriFieldTypeString, , 5)
  pFields.Add MyGeneralOperations.CreateNewField("Species", esriFieldTypeString, , 35)
  pFields.Add MyGeneralOperations.CreateNewField("Number_Observations", esriFieldTypeInteger)
  pFields.Add MyGeneralOperations.CreateNewField("Area_Sq_Cm", esriFieldTypeDouble)
  pFields.Add MyGeneralOperations.CreateNewField("Proportion_Quadrat", esriFieldTypeString, , 15)

  Dim pNewTable As ITable
  Set pNewTable = MyGeneralOperations.CreateGDBTable(pNewWS, "Basal_Cover_per_Species_by_Quadrat_and_Year", pFields)

  Dim lngNewSiteIndex As Long
  Dim lngNewQuadratIndex As Long
  Dim lngNewYearIndex As Long
  Dim lngNewTypeIndex As Long
  Dim lngNewSpeciesIndex As Long
  Dim lngNewObsCountIndex As Long
  Dim lngNewAreaIndex As Long
  Dim lngNewProportionIndex As Long

  lngNewSiteIndex = pNewTable.FindField("Site")
  lngNewQuadratIndex = pNewTable.FindField("Quadrat")
  lngNewYearIndex = pNewTable.FindField("Year")
  lngNewTypeIndex = pNewTable.FindField("Type")
  lngNewSpeciesIndex = pNewTable.FindField("Species")
  lngNewObsCountIndex = pNewTable.FindField("Number_Observations")
  lngNewAreaIndex = pNewTable.FindField("Area_Sq_Cm")
  lngNewProportionIndex = pNewTable.FindField("Proportion_Quadrat")

  Dim strPurpose As String
  strPurpose = "List of all species observed in all quadrats over all years."

  Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewTable, strAbstract, strPurpose)

  Dim pRowBuffer As IRowBuffer
  Dim pCursor As ICursor
  Set pCursor = pNewTable.Insert(True)
  Set pRowBuffer = pNewTable.CreateRowBuffer

  pSBar.ShowProgressBar "Cover Basal Area Pass 2 of 2...", 0, lngSortCounter, 1, True
  pProg.position = 0

  For lngIndex = 0 To lngSortCounter
    pProg.Step
    DoEvents
    strSite = strSortArray(0, lngIndex)
    strQuadrat = strSortArray(1, lngIndex)
    strBaseQuery = strSortArray(2, lngIndex)

    strYears = ReturnArrayOfValues(pCoverFClass, lngCoverYearIndex, strBaseQuery)
    For lngYearIndex = 0 To UBound(strYears, 2)
      strYear = strYears(0, lngYearIndex)
      strBaseYearQuery = strBaseQuery & " AND " & strPrefix & "Year" & strSuffix & " = '" & strYear & "'"

      strSpeciesArray = ReturnArrayOfValues(pCoverFClass, lngCoverSpeciesIndex, strBaseYearQuery)
      For lngSpeciesIndex = 0 To UBound(strSpeciesArray, 2)
        strSpecies = strSpeciesArray(0, lngSpeciesIndex)
        strBaseSpeciesQuery = strBaseYearQuery & " AND " & strPrefix & "Species" & strSuffix & " = '" & strSpecies & "'"

        FillCountAndAreaForSpecies pCoverFClass, strBaseSpeciesQuery, lngObservationCount, dblTotalArea

        pRowBuffer.Value(lngNewSiteIndex) = strSite
        pRowBuffer.Value(lngNewQuadratIndex) = strQuadrat
        pRowBuffer.Value(lngNewYearIndex) = strYear
        pRowBuffer.Value(lngNewTypeIndex) = strType
        pRowBuffer.Value(lngNewSpeciesIndex) = strSpecies
        pRowBuffer.Value(lngNewObsCountIndex) = lngObservationCount
        pRowBuffer.Value(lngNewAreaIndex) = dblTotalArea
        pRowBuffer.Value(lngNewProportionIndex) = Format(dblTotalArea / 10000, "0.00%")
        pCursor.InsertRow pRowBuffer

      Next lngSpeciesIndex
    Next lngYearIndex
    pCursor.Flush
  Next lngIndex

  pCursor.Flush

  MyGeneralOperations.ExportToCSV pNewTable, strNewAncillaryFolder & "\Basal_Cover_per_Species_by_Quadrat_and_Year.csv", _
        True, False, False, True, , pApp, True

ClearMemory:
  Erase strSortArray
  Set pDoneColl = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strYears
  Erase strSpeciesArray
  Set pFields = Nothing
  Set pNewTable = Nothing
  Set pRowBuffer = Nothing
  Set pCursor = Nothing

End Sub

Public Sub FillCountAndAreaForSpecies(pCoverFClass As IFeatureClass, strQueryString As String, lngObservationCount As Long, _
    dblTotalSqCm As Double)

  Dim pQueryFilt As IQueryFilter
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPoly As IPolygon
  Dim pArea As IArea
  Dim dblCumulativeArea As Double
  Dim lngCounter As Long

  lngCounter = 0
  dblCumulativeArea = 0

  Set pQueryFilt = New QueryFilter
  pQueryFilt.WhereClause = strQueryString
  Set pFCursor = pCoverFClass.Search(pQueryFilt, True)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pPoly = pFeature.Shape
    Set pArea = pPoly
    dblCumulativeArea = dblCumulativeArea + (pArea.Area * 10000)
    lngCounter = lngCounter + 1
    Set pFeature = pFCursor.NextFeature
  Loop

  dblTotalSqCm = dblCumulativeArea
  lngObservationCount = lngCounter

ClearMemory:
  Set pQueryFilt = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pPoly = Nothing
  Set pArea = Nothing

End Sub

Public Function ReturnArrayOfValues(pCoverFClass As IFeatureClass, lngFieldIndex As Long, strQueryString As String) As String()

  Dim pQueryFilt As IQueryFilter
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strReturn() As String
  Dim lngArrayIndex As Long
  Dim pDoneColl As New Collection
  Dim lngCounter As Long
  Dim strValue As String

  lngArrayIndex = -1
  Set pQueryFilt = New QueryFilter
  pQueryFilt.WhereClause = strQueryString
  Set pFCursor = pCoverFClass.Search(pQueryFilt, True)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strValue = pFeature.Value(lngFieldIndex)
    If MyGeneralOperations.CheckCollectionForKey(pDoneColl, strValue) Then
      lngCounter = pDoneColl.Item(strValue)
      pDoneColl.Remove strValue
    Else
      lngCounter = 0
      lngArrayIndex = lngArrayIndex + 1
      ReDim Preserve strReturn(1, lngArrayIndex)
      strReturn(0, lngArrayIndex) = strValue
    End If
    lngCounter = lngCounter + 1
    pDoneColl.Add lngCounter, strValue
    Set pFeature = pFCursor.NextFeature
  Loop

  Dim lngIndex As Long
  For lngIndex = 0 To lngArrayIndex
    strValue = strReturn(0, lngArrayIndex)
    strReturn(1, lngArrayIndex) = pDoneColl.Item(strValue)
  Next lngIndex

  QuickSort.StringAscending_TwoDimensional strReturn, 0, lngArrayIndex, 0, 1

  ReturnArrayOfValues = strReturn

ClearMemory:
  Set pQueryFilt = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strReturn
  Set pDoneColl = Nothing

End Function

Public Sub DeclareWorkspaces(strOrigShapefiles As String, Optional strModifiedRoot As String, _
    Optional strShiftedRoot As String, Optional strExportBase As String, Optional strRecreatedModifiedRoot As String, _
    Optional strSetFolder As String, Optional strExtractShapefileFolder As String, Optional strFinalFolder As String)

  Dim booUseCurrentDate As Boolean
  booUseCurrentDate = True

  Dim strSpecifiedDate As String
  strSpecifiedDate = "2024_12_22"

  Dim strDate As String
  Dim strDateSplit() As String
  Dim strCurrentDate As String

  If booUseCurrentDate Then
    strDate = Replace(Format(Now, "yyyy_mm_dd"), "Sep_", "Sept_")   ' "2021_06_08"
    strCurrentDate = Replace(Format(Now, "mmm_d_yyyy"), "Sep_", "Sept_")
  Else
    strDate = strSpecifiedDate
    strDateSplit = Split(strDate, "_")
    strCurrentDate = Format(DateSerial(CInt(strDateSplit(0)), CInt(strDateSplit(1)), CInt(strDateSplit(2))), "mmm_d_yyyy")
    strCurrentDate = Replace(strCurrentDate, "Sep_", "Sept_")
  End If

  strOrigShapefiles = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\contemporary_data_" & strDate
  strModifiedRoot = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Modified_Data_" & strDate
  strRecreatedModifiedRoot = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Cleaned_Data_" & strDate
  strShiftedRoot = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Cleaned_Data_" & strDate & "_Shift"
  strExportBase = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Map_Exports_" & strDate
  strSetFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate
  strExtractShapefileFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Species_Shapefile_Extractions"
  strFinalFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Final_Datasets_" & strCurrentDate

End Sub

Public Sub OrganizeData()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pRedigitizeColl As Collection
  Set pRedigitizeColl = ReturnReplacementColl

  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Dim lngCount As Long
  Dim lngIndex As Long
  Dim strPath As String
  Dim strModPath As String
  Dim lngCounter As Long

  Dim strQuadrat As String
  Dim strReplaceName As String

  Dim strExt As String
  Dim booTransfer As Boolean
  Dim strFilename As String
  Dim strNewDir As String

  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim strCombinePath As String
  Dim strSetFolder As String
  Call DeclareWorkspaces(strCombinePath, , , , , strSetFolder)

  If Not aml_func_mod.ExistFileDir(strCombinePath) Then
    MyGeneralOperations.CreateNestedFoldersByPath strCombinePath
  End If
  If Not aml_func_mod.ExistFileDir(strSetFolder & "\Description_of_Analysis.docx") Then
    CopyFile "D:\arcGIS_stuff\consultation\Margaret_Moore\Data_to_include_in_publication\Description_of_Analysis.docx", _
      strSetFolder & "\Description_of_Analysis.docx", 0
  End If

  Dim strDir As String
  Dim pAllPaths As esriSystem.IStringArray
  Dim varCheckArray() As Variant
  Dim strCheckPathReport
  Dim pDataset As IDataset

  Dim strSourcePath1 As String
  strSourcePath1 = "D:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data - Original"

  Dim strSourcePath2 As String
  strSourcePath2 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_Aug_24_2018\" & _
      "2017 Hill Wild Bill digitized 1m2 quadrats\Hill_Wild_Bill_Contemporary"

  Dim strSourcePath3 As String
  strSourcePath3 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_May_29_2019\Hill-WildBill_2018"

  Dim strSourcePath4 As String
  strSourcePath4 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_Aug_24_2018\" & _
      "2017 Hill Wild Bill digitized 1m2 quadrats\Wild Bill"

  Dim strSourcePath5 As String
  strSourcePath5 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_As_of_May_14_2020"

  Dim strSourcePath6 As String
  strSourcePath6 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_from_2020\Final"

  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath1, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  Dim strSourcePath7 As String
  strSourcePath7 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_June_18_2021\"

  Dim strSourcePath8 As String
  strSourcePath8 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_Jan_18_2022\Final"

  Dim strSourcePath9 As String
  strSourcePath9 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_Nov_2022\Final"

  Dim strSourcePath10 As String
  strSourcePath10 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_March_23_2024\Final"

  Dim strSourcePath11 As String
  strSourcePath11 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_Oct_11_2024"

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 1: " & Format(lngCount, "#,##0") & " paths found..."
  Dim pCopyFClass As IFeatureClass
  Dim pDoneColl As New Collection
  Dim pUnknownSpRef As IUnknownCoordinateSystem
  Set pUnknownSpRef = New UnknownCoordinateSystem
  Dim pGeoDataset As IGeoDataset

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To pAllPaths.Count - 1
      pProg.Step
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        strModPath = Replace(strPath, strSourcePath1, strCombinePath, , , vbTextCompare)
            "D:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data - Original", , , vbTextCompare)

        If InStr(1, strModPath, "_2008_", vbTextCompare) = 0 Then

          If Not aml_func_mod.ExistFileDir(strModPath) Then

            strFilename = aml_func_mod.ReturnFilename2(strPath)
            strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
            strFilename = aml_func_mod.ClipExtension2(strFilename)

            If MyGeneralOperations.CheckCollectionForKey(pRedigitizeColl, strFilename) Then
              UpdateCheckArray varCheckArray, strPath
              If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strFilename) Then
                Set pDataset = pRedigitizeColl.Item(strFilename)
                Set pCopyFClass = CopyFeatureClassToShapefile(pDataset, strModPath)
                Debug.Print "...Using redigitized feature class '" & pDataset.BrowseName & "..."
                pDoneColl.Add True, strFilename

                UpdateCheckArray varCheckArray, strPath
              Else
                Debug.Print "...Already copied over '" & strFilename & "..."
              End If

            Else

              strDir = aml_func_mod.ReturnDir3(strModPath, False)
              If Not aml_func_mod.ExistFileDir(strDir) Then
                MyGeneralOperations.CreateNestedFoldersByPath strDir
              End If
              lngCounter = lngCounter + 1
              CopyFile strPath, strModPath, True

              UpdateCheckArray varCheckArray, strPath
            End If
          Else
            UpdateCheckArray varCheckArray, strPath
          End If
        End If   ' END EXCLUDING 2008
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Dim pDataObj As New MSForms.DataObject
      pDataObj.SetText strCheckPathReport
      pDataObj.PutInClipboard
      Set pDataObj = Nothing
      Debug.Print "Original Data: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2017 As Collection
  Dim pConvertNames2017ToOld As Collection
  Dim varNameLinks() As Variant
  Call FillNameConverters(varNameLinks, pConvertNames2017ToOld, pConvertNamesOldTo2017)

  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath2, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 2: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)

      If InStr(1, strPath, "WB123_2017_", vbTextCompare) = 0 Then  ' Special case to exclude

        pProg.Step
        If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
          If lngIndex Mod 100 = 0 Then
            DoEvents
          End If
          If InStr(1, strPath, "VBA", vbTextCompare) > 0 Then
            DoEvents
          End If
          strExt = aml_func_mod.GetExtensionText(strPath)
          booTransfer = False

          If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
              StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
              StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
              StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
              StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
              StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
              StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
              StrComp(strExt, "atx", vbTextCompare) = 0 Then
            booTransfer = True
          ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
            booTransfer = True
            strExt = ".shp.xml"
          End If

          If booTransfer Then
            strFilename = aml_func_mod.ReturnFilename2(strPath)
            strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
            strFilename = aml_func_mod.ClipExtension2(strFilename)
            strReplaceName = Replace(strFilename, "_C", "", , , vbTextCompare)
            strReplaceName = Replace(strReplaceName, "_D", "", , , vbTextCompare)

            If StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0 Or _
                StrComp(Right(strFilename, 2), "_D", vbTextCompare) = 0 Then

              UpdateCheckArray varCheckArray, strPath

              If MyGeneralOperations.CheckCollectionForKey(pConvertNames2017ToOld, strReplaceName) Then
                strQuadrat = pConvertNames2017ToOld.Item(strReplaceName)
                strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2017" & _
                    IIf(StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0, "_C", "_D") & "." & strExt

                If Not aml_func_mod.ExistFileDir(strModPath) Then
                  strDir = aml_func_mod.ReturnDir3(strModPath, False)
                  If Not aml_func_mod.ExistFileDir(strDir) Then
                    Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                  Else
                    lngCounter = lngCounter + 1
                    CopyFile strPath, strModPath, True
                  End If
                End If

              Else
                Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                    "...Path = '" & strPath & "'..."
              End If
            End If
          End If
        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2017a: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath4, "WB123_2017_")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 2B: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)

      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "VBA", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_C", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D", "", , , vbTextCompare)

          If StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 2), "_D", vbTextCompare) = 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2017ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2017ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2017" & _
                  IIf(StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0, "_C", "_D") & "." & strExt

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2017b: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2018 As Collection
  Dim pConvertNames2018ToOld As Collection
  Dim varNameLinks_2018() As Variant
  Call FillNameConverters_2018(varNameLinks_2018, pConvertNames2018ToOld, pConvertNamesOldTo2018)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath3, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 3: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "VBA", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_C", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "2018C", "2018", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "2018D", "2018", , , vbTextCompare)

          If StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 2), "_D", vbTextCompare) = 0 Or _
              InStr(1, Right(strFilename, 6), "2018C", vbTextCompare) > 0 Or _
              InStr(1, Right(strFilename, 6), "2018D", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2018ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2018ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2018" & _
                  IIf(StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0 Or _
                      StrComp(Right(strFilename, 5), "2018C", vbTextCompare) = 0, "_C", "_D") & "." & strExt

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2020: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2019 As Collection
  Dim pConvertNames2019ToOld As Collection
  Dim varNameLinks_2019() As Variant
  Call FillNameConverters_2019(varNameLinks_2019, pConvertNames2019ToOld, pConvertNamesOldTo2019)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath5, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 4: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "Q_WB_114_2019", vbTextCompare) > 0 And InStr(1, strPath, ".shp", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2019ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2019ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2019" & _
                  IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                      (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2020: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2020 As Collection
  Dim pConvertNames2020ToOld As Collection
  Dim varNameLinks_2020() As Variant
  Call FillNameConverters_2020(varNameLinks_2020, pConvertNames2020ToOld, pConvertNamesOldTo2020)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath6, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 5: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "Q_WB_114_2020", vbTextCompare) > 0 And InStr(1, strPath, ".shp", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2020ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2020ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2020" & _
                  IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                      (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2020: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2021 As Collection
  Dim pConvertNames2021ToOld As Collection
  Dim varNameLinks_2021() As Variant
  Call FillNameConverters_2021(varNameLinks_2021, pConvertNames2021ToOld, pConvertNamesOldTo2021)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath8, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 6: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2021ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2021ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2021" & _
                  IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                      (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2021: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldToMissing2021 As Collection
  Dim pConvertNamesMissing2021ToOld As Collection
  Dim varNameLinks_Missing2021() As Variant
  Call FillNameConverters_Missing2021(varNameLinks_Missing2021, pConvertNamesMissing2021ToOld, pConvertNamesOldToMissing2021)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath7, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 7 [Missing Data]: " & Format(lngCount, "#,##0") & " paths found..."

  Dim strMissingYear As String

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If

        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_C", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D", "", , , vbTextCompare)

          If StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 2), "_D", vbTextCompare) = 0 Then

            UpdateCheckArray varCheckArray, strPath

            If InStr(1, strFilename, "_2019_", vbTextCompare) > 0 Then
              strMissingYear = "_2019"
            ElseIf InStr(1, strFilename, "_2016_", vbTextCompare) > 0 Then
              strMissingYear = "_2016"
            ElseIf InStr(1, strFilename, "_2007_", vbTextCompare) > 0 Then
              strMissingYear = "_2007"
            ElseIf InStr(1, strFilename, "_2009_", vbTextCompare) > 0 Then
              strMissingYear = "_2009"
            ElseIf InStr(1, strFilename, "_2010_", vbTextCompare) > 0 Then
              strMissingYear = "_2010"
            ElseIf InStr(1, strFilename, "_2011_", vbTextCompare) > 0 Then
              strMissingYear = "_2011"
            ElseIf InStr(1, strFilename, "_2017_", vbTextCompare) > 0 Then
              strMissingYear = "_2017"
            ElseIf InStr(1, strFilename, "_2018_", vbTextCompare) > 0 Then
              strMissingYear = "_2018"
            Else
              MsgBox "Problem with missing year..." & vbCrLf & strFilename
              DoEvents
            End If

            If MyGeneralOperations.CheckCollectionForKey(pConvertNamesMissing2021ToOld, strReplaceName) Then
              strQuadrat = pConvertNamesMissing2021ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & strMissingYear & _
                  IIf((InStr(1, strFilename, "_C", vbTextCompare) > 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "Missing2021: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2022 As Collection
  Dim pConvertNames2022ToOld As Collection
  Dim varNameLinks_2022() As Variant
  Call FillNameConverters_2022(varNameLinks_2022, pConvertNames2022ToOld, pConvertNamesOldTo2022)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath9, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 8 [2022]: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          If strFilename = "Q_FV_22241_2022_D_NO_SpF" Then strFilename = "Q_FV_22241_2022_DF" ' FOR NOV. 2022 DATA
          If strFilename = "Q_FV_22241_2022_C_NO_SpF" Then strFilename = "Q_FV_22241_2022_CF" ' FOR NOV. 2022 DATA
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_NO_SpF", "", , , vbTextCompare) ' FOR NOV. 2022 DATA
          strReplaceName = Replace(strReplaceName, "_NO_Sp", "", , , vbTextCompare) ' FOR NOV. 2022 DATA

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_NO_Sp", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2022ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2022ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2022" & _
                  IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                      (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2022: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2023 As Collection
  Dim pConvertNames2023ToOld As Collection
  Dim varNameLinks_2023() As Variant
  Call FillNameConverters_2023(varNameLinks_2023, pConvertNames2023ToOld, pConvertNamesOldTo2023)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath10, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 9 [2023]: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "Q_WB_29025", vbTextCompare) > 0 And InStr(1, strPath, ".shp", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_NO_SpF", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_NO_Spp", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_NO_SP", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_C", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_D", "", , , vbTextCompare) ' FOR Mar. 2024 DATA

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_NO_Sp", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2023ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2023ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2023" & _
                  IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                      (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2023: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim pConvertNamesOldTo2024 As Collection
  Dim pConvertNames2024ToOld As Collection
  Dim varNameLinks_2024() As Variant
  Call FillNameConverters_2024(varNameLinks_2024, pConvertNames2024ToOld, pConvertNamesOldTo2024)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath11, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 10 [2024]: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "Q_WB_29025", vbTextCompare) > 0 And InStr(1, strPath, ".shp", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_NO_SpF", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_NO_Spp", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_NO_SP", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_C", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_D", "", , , vbTextCompare) ' FOR Mar. 2024 DATA
          strReplaceName = Replace(strReplaceName, "_NoSp", "", , , vbTextCompare) ' FOR Oct. 2024 DATA

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_NO_Sp", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2024ToOld, strReplaceName) Then
              strQuadrat = pConvertNames2024ToOld.Item(strReplaceName)
              strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2024" & _
                  IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                      (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0) Or _
                      (StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0), "_C", "_D") & "." & strExt
              strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

              If Not aml_func_mod.ExistFileDir(strModPath) Then
                strDir = aml_func_mod.ReturnDir3(strModPath, False)
                If Not aml_func_mod.ExistFileDir(strDir) Then
                  Debug.Print "Failed to find folder '" & strDir & "'! ************************"
                Else
                  lngCounter = lngCounter + 1
                  CopyFile strPath, strModPath, True
                End If
              End If

            Else
              Debug.Print "Failed to find '" & strReplaceName & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2024: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim strQuadrats() As String
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim lngFeatCount As Long
  Dim pQuadData As Collection
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev(strQuadrats, pPlotToQuadratConversion, pQuadratToPlotConversion, _
      varSites, varSitesSpecific)

  Dim pSHPfiles As esriSystem.IStringArray
  Set pSHPfiles = MyGeneralOperations.ReturnFilesFromNestedFolders(strCombinePath, "shp")

  Debug.Print "pSHPfiles.Count = " & Format(pSHPfiles.Count, "0")

  Dim pDone1 As New Collection
  Dim strNames1() As String
  Dim lngNameIndex As Long
  Dim pWSFact As IWorkspaceFactory
  Dim pWS As IFeatureWorkspace

  Dim pDatasets As IEnumDataset
  Dim strName As String
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim lngSrcSpeciesNameIndex As Long
  Dim lngVerbSpeciesNameIndex As Long
  Dim lngVerbTypeIndex As Long
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pClone As IClone
  Dim pFeature As IFeature
  Dim strPrjPath As String

  pSBar.ShowProgressBar "Adding Verbatim Fields...", 0, pSHPfiles.Count, 1, True
  pProg.position = 0
  lngNameIndex = -1
  For lngIndex = 0 To pSHPfiles.Count - 1
    strPath = pSHPfiles.Element(lngIndex)
    strPrjPath = strPath
    strPrjPath = aml_func_mod.SetExtension(strPrjPath, "prj")
    If aml_func_mod.ExistFileDir(strPrjPath) Then
      Kill strPrjPath
    End If
    pProg.Step
    If lngIndex Mod 25 = 0 Then
      DoEvents
    End If
    Debug.Print MyGeneralOperations.SpacesInFrontOfText(Format(lngIndex, "#,##0"), 5) & "] " & aml_func_mod.ReturnFilename2(strPath)
    strDir = aml_func_mod.ReturnDir3(strPath, False)
    Set pWSFact = New ShapefileWorkspaceFactory
    Set pWS = pWSFact.OpenFromFile(strDir, 0)
    Set pFClass = pWS.OpenFeatureClass(aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath)))
    AddVerbatimFields pFClass, pQuadData

  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  Debug.Print "Done..."

ClearMemory:
  Set pRedigitizeColl = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pAllPaths = Nothing
  Erase varCheckArray
  Set pDataset = Nothing
  Set pCopyFClass = Nothing
  Set pDoneColl = Nothing
  Set pUnknownSpRef = Nothing
  Set pGeoDataset = Nothing
  Set pDataObj = Nothing
  Set pConvertNamesOldTo2017 = Nothing
  Set pConvertNames2017ToOld = Nothing
  Erase varNameLinks
  Set pConvertNamesOldTo2018 = Nothing
  Set pConvertNames2018ToOld = Nothing
  Erase varNameLinks_2018
  Set pConvertNamesOldTo2019 = Nothing
  Set pConvertNames2019ToOld = Nothing
  Erase varNameLinks_2019
  Set pConvertNamesOldTo2020 = Nothing
  Set pConvertNames2020ToOld = Nothing
  Erase varNameLinks_2020
  Set pConvertNamesOldTo2021 = Nothing
  Set pConvertNames2021ToOld = Nothing
  Erase varNameLinks_2021
  Set pConvertNamesOldToMissing2021 = Nothing
  Set pConvertNamesMissing2021ToOld = Nothing
  Erase varNameLinks_Missing2021
  Set pConvertNamesOldTo2022 = Nothing
  Set pConvertNames2022ToOld = Nothing
  Erase varNameLinks_2022
  Set pConvertNamesOldTo2023 = Nothing
  Set pConvertNames2023ToOld = Nothing
  Erase varNameLinks_2023
  Set pConvertNamesOldTo2024 = Nothing
  Set pConvertNames2024ToOld = Nothing
  Erase varNameLinks_2024
  Erase strQuadrats
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Set pQuadData = Nothing
  Erase varSites
  Erase varSitesSpecific
  Set pSHPfiles = Nothing
  Set pDone1 = Nothing
  Erase strNames1
  Set pWSFact = Nothing
  Set pWS = Nothing
  Set pDatasets = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pClone = Nothing
  Set pFeature = Nothing

End Sub

Public Function ReturnMissingShapefiles(varCheckArray() As Variant) As String
  Dim lngIndex As Long
  Dim strCheck As String
  Dim strReport As String
  For lngIndex = 0 To UBound(varCheckArray, 2)
    If varCheckArray(1, lngIndex) = False Then
      strCheck = varCheckArray(0, lngIndex)
      strReport = strReport & CStr(lngIndex) & "] " & strCheck & vbCrLf
    End If
  Next lngIndex

  ReturnMissingShapefiles = strReport
End Function

Public Sub UpdateCheckArray(varCheckArray() As Variant, strPath As String)
  Dim lngIndex As Long
  Dim strCheck As String
  For lngIndex = 0 To UBound(varCheckArray, 2)
    strCheck = varCheckArray(0, lngIndex)
    If InStr(1, strCheck, "Q_WB_29025", vbTextCompare) > 0 Then
      DoEvents
    End If
    If InStr(1, strPath, strCheck, vbTextCompare) > 0 Then
      varCheckArray(1, lngIndex) = True
    End If
  Next lngIndex

End Sub

Public Function BuildCheckArray(pAllPaths As esriSystem.IStringArray) As Variant()

  Dim varReturn() As Variant
  Dim lngCounter As Long

  lngCounter = -1

  Dim lngIndex As Long
  Dim strVal As String
  For lngIndex = 0 To pAllPaths.Count - 1
    strVal = pAllPaths.Element(lngIndex)
    If StrComp(Right(strVal, 4), ".dbf", vbTextCompare) = 0 Then
      lngCounter = lngCounter + 1
      ReDim Preserve varReturn(1, lngCounter)
      strVal = aml_func_mod.ReturnFilename2(strVal)
      strVal = aml_func_mod.ClipExtension2(strVal)
      varReturn(0, lngCounter) = strVal
      varReturn(1, lngCounter) = False
    End If
  Next lngIndex

  BuildCheckArray = varReturn

End Function

Public Sub FillNameConverters(varNameLinks() As Variant, p2017toOld As Collection, pOldTo2017 As Collection)

  ReDim varNameLinks(86)
  varNameLinks(0) = Array("BF11999_2017", "Q28")
  varNameLinks(1) = Array("BF12000_2017", "Q29")
  varNameLinks(2) = Array("BF2004_16_2017", "Q7")
  varNameLinks(3) = Array("BF30713_2017", "Q63")
  varNameLinks(4) = Array("BF30714_2017", "Q64")
  varNameLinks(5) = Array("BF30715_2017", "Q65")
  varNameLinks(6) = Array("BF30717_19", "Q66")
  varNameLinks(7) = Array("BF30717_2017", "Q66")
  varNameLinks(8) = Array("BF30719_2017", "Q67")
  varNameLinks(9) = Array("BF30720_2017", "Q68")
  varNameLinks(10) = Array("BS-26345-2017", "Q42")
  varNameLinks(11) = Array("BS-26346-2017", "Q43")
  varNameLinks(12) = Array("BS-26347-2017", "Q44")
  varNameLinks(13) = Array("BS-30743-2017", "Q91")
  varNameLinks(14) = Array("BS_2004_46_2017", "Q9")
  varNameLinks(15) = Array("BS_26348_2017", "Q45")
  varNameLinks(16) = Array("BS_30741_2017", "Q89")
  varNameLinks(17) = Array("BS_30742_2017", "Q90")
  varNameLinks(18) = Array("BS_30744_2017", "Q92")
  varNameLinks(19) = Array("BS_30747_2017", "Q95")
  varNameLinks(20) = Array("BS_30748_2017", "Q96")
  varNameLinks(21) = Array("BS_30749_2017", "Q97")
  varNameLinks(22) = Array("FP31_2017", "Q79")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(23) = Array("FP32_2017", "Q80")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(24) = Array("FP33_2017", "Q81")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(25) = Array("FP34_2017", "Q82")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(26) = Array("FP35_2017", "Q83")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(27) = Array("FP39_2017", "Q87")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(28) = Array("FP40_2017", "Q88")    ' ASSUMING FPxx = FRY PARK, 307xx
  varNameLinks(29) = Array("FS9009H-494_2017", "Q494")
  varNameLinks(30) = Array("FS9009H-498_2017", "Q498")
  varNameLinks(31) = Array("FV_21114_2017", "Q30")
  varNameLinks(32) = Array("FV_21174_2017", "Q31")
  varNameLinks(33) = Array("FV_21262_2017", "Q32")
  varNameLinks(34) = Array("FV_21269_2017", "Q33")
  varNameLinks(35) = Array("FV_22126_2017", "Q34")
  varNameLinks(36) = Array("FV_22156_2017", "Q35")
  varNameLinks(37) = Array("FV_22241_2017", "Q36")
  varNameLinks(38) = Array("FV_22244_2017", "Q37")
  varNameLinks(39) = Array("FV_23155_2017", "Q38")
  varNameLinks(40) = Array("FV_23159_2017", "Q39")
  varNameLinks(41) = Array("RL_26339_2017", "Q40")
  varNameLinks(42) = Array("RL_26369_2017", "Q46")
  varNameLinks(43) = Array("RL_26370_2017", "Q47")
  varNameLinks(44) = Array("RL_30721_2017", "Q69")
  varNameLinks(45) = Array("RL_30722_2017", "Q70")
  varNameLinks(46) = Array("RL_30723_2017", "Q71")
  varNameLinks(47) = Array("RL_30724_2017", "Q72")
  varNameLinks(48) = Array("RL_30725_2017", "Q73")
  varNameLinks(49) = Array("RL_30726_2017", "Q74")
  varNameLinks(50) = Array("RL_30728_2017", "Q76")
  varNameLinks(51) = Array("RL_30729_2017", "Q77")
  varNameLinks(52) = Array("RL_30730_2017", "Q78")
  varNameLinks(53) = Array("RT_2004_10_2017", "Q6")
  varNameLinks(54) = Array("RT_30701_2017", "Q53")
  varNameLinks(55) = Array("RT_30702_2017", "Q54")
  varNameLinks(56) = Array("RT_30703_2017", "Q55")
  varNameLinks(57) = Array("RT_30705_2017", "Q57")
  varNameLinks(58) = Array("RT_30706_2017", "Q58")
  varNameLinks(59) = Array("RT_30707_2017", "Q59")
  varNameLinks(60) = Array("RT_30707__2017", "Q59")
  varNameLinks(61) = Array("RT_30709_2017", "Q60")
  varNameLinks(62) = Array("WB101_2017", "Q10")
  varNameLinks(63) = Array("WB102_2017", "Q11")
  varNameLinks(64) = Array("WB103_2017", "Q12")
  varNameLinks(65) = Array("WB104_2017", "Q13")
  varNameLinks(66) = Array("WB105_2017", "Q14")
  varNameLinks(67) = Array("WB106_2017", "Q106")
  varNameLinks(68) = Array("WB107_2017", "Q15")
  varNameLinks(69) = Array("WB108_2017", "Q16")
  varNameLinks(70) = Array("WB109_2017", "Q17")
  varNameLinks(71) = Array("WB110_2017", "Q18")
  varNameLinks(72) = Array("WB114_2017", "Q19")
  varNameLinks(73) = Array("WB115_2017", "Q20")
  varNameLinks(74) = Array("WB119_2017", "Q21")
  varNameLinks(75) = Array("WB120_2017", "Q22")
  varNameLinks(76) = Array("WB121_2017", "Q23")
  varNameLinks(77) = Array("WB122_2017", "Q24")
  varNameLinks(78) = Array("WB123_2017", "Q25")
  varNameLinks(79) = Array("WB124_2017", "Q26")
  varNameLinks(80) = Array("WB125_2017", "Q27")
  varNameLinks(81) = Array("WB29003_2017", "Q48")
  varNameLinks(82) = Array("WB29004_2017", "Q49")
  varNameLinks(83) = Array("WB29016_2017", "Q50")
  varNameLinks(84) = Array("WB29017_2017", "Q51")
  varNameLinks(85) = Array("WB29025_2017", "Q52")
  varNameLinks(86) = Array("WB3_2017", "Q3")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2017 = New Collection
  Set p2017toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks)
    varSubArray = varNameLinks(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2017, CStr(varSubArray(1))) Then
      pOldTo2017.Add varSubArray(0), varSubArray(1)
    End If
    p2017toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2024(varNameLinks_2024() As Variant, p2024toOld As Collection, pOldTo2024 As Collection)

  ReDim varNameLinks_2024(107)

  varNameLinks_2024(0) = Array("Q_BF_11999_2024", "Q28")
  varNameLinks_2024(1) = Array("Q_BF_12000_2024", "Q29")
  varNameLinks_2024(2) = Array("Q_BF_12000_2024_NO_Sp", "Q29")
  varNameLinks_2024(3) = Array("Q_BF_16_2024", "Q7")
  varNameLinks_2024(4) = Array("Q_BF_30713_2024", "Q63")
  varNameLinks_2024(5) = Array("Q_BF_30714_2024", "Q64")
  varNameLinks_2024(6) = Array("Q_BF_30714_2024_NO_Sp", "Q64")
  varNameLinks_2024(7) = Array("Q_BF_30715_2024", "Q65")
  varNameLinks_2024(8) = Array("Q_BF_30715_2024_NO_Sp", "Q65")
  varNameLinks_2024(9) = Array("Q_BF_30717_2024", "Q66")
  varNameLinks_2024(10) = Array("Q_BF_30719_2024", "Q67")
  varNameLinks_2024(11) = Array("Q_BF_30720_2024", "Q68")
  varNameLinks_2024(12) = Array("Q_BS_26345_2024", "Q42")
  varNameLinks_2024(13) = Array("Q_BS_26346_2024", "Q43")
  varNameLinks_2024(14) = Array("Q_BS_26347_2024", "Q44")
  varNameLinks_2024(15) = Array("Q_BS_26347_2024_NO_SP", "Q44")
  varNameLinks_2024(16) = Array("Q_BS_26348_2024", "Q45")
  varNameLinks_2024(17) = Array("Q_BS_30741_2024", "Q89")
  varNameLinks_2024(18) = Array("Q_BS_30742_2024", "Q90")
  varNameLinks_2024(19) = Array("Q_BS_30743_2024", "Q91")
  varNameLinks_2024(20) = Array("Q_BS_30744_2024", "Q92")
  varNameLinks_2024(21) = Array("Q_BS_30747_2024", "Q95")
  varNameLinks_2024(22) = Array("Q_BS_30748_2024", "Q96")
  varNameLinks_2024(23) = Array("Q_BS_30749_2024", "Q97")
  varNameLinks_2024(24) = Array("Q_BS_46_2024", "Q9")
  varNameLinks_2024(25) = Array("Q_FP_30731_2024", "Q79")
  varNameLinks_2024(26) = Array("Q_FP_30732_2024", "Q80")
  varNameLinks_2024(27) = Array("Q_FP_30733_2024", "Q81")
  varNameLinks_2024(28) = Array("Q_FP_30734_2024", "Q82")
  varNameLinks_2024(29) = Array("Q_FP_30735_2024", "Q83")
  varNameLinks_2024(30) = Array("Q_FP_30739_2024", "Q87")
  varNameLinks_2024(31) = Array("Q_FP_30740_2024", "Q88")
  varNameLinks_2024(32) = Array("Q_FS9009H_494_2024", "Q494")
  varNameLinks_2024(33) = Array("Q_FS9009H_498_2024", "Q498")
  varNameLinks_2024(34) = Array("Q_FV_21114_2024", "Q30")
  varNameLinks_2024(35) = Array("Q_FV_21114_2024_NO_Sp", "Q30")
  varNameLinks_2024(36) = Array("Q_FV_21174_2024", "Q31")
  varNameLinks_2024(37) = Array("Q_FV_21262_2024", "Q32")
  varNameLinks_2024(38) = Array("Q_FV_21262_2024_NO_Sp", "Q32")
  varNameLinks_2024(39) = Array("Q_FV_21269_2024", "Q33")
  varNameLinks_2024(40) = Array("Q_FV_22126_2024", "Q34")
  varNameLinks_2024(41) = Array("Q_FV_22156_2024", "Q35")
  varNameLinks_2024(42) = Array("Q_FV_22241_2024_NO_Sp", "Q36")
  varNameLinks_2024(43) = Array("Q_FV_22244_2024", "Q37")
  varNameLinks_2024(44) = Array("Q_FV_23155_2024", "Q38")
  varNameLinks_2024(45) = Array("Q_FV_23155_2024_NO_Sp", "Q38")
  varNameLinks_2024(46) = Array("Q_FV_23159_2024", "Q39")
  varNameLinks_2024(47) = Array("Q_FV_23159_2024_NO_Sp", "Q39")
  varNameLinks_2024(48) = Array("Q_GK_101_2024", "Q10")
  varNameLinks_2024(49) = Array("Q_GK_102_2024", "Q11")
  varNameLinks_2024(50) = Array("Q_GK_102_2024_NO_Sp", "Q11")
  varNameLinks_2024(51) = Array("Q_GK_103_2024", "Q12")
  varNameLinks_2024(52) = Array("Q_GK_103_2024_NO_Sp", "Q12")
  varNameLinks_2024(53) = Array("Q_GK_104_2024", "Q13")
  varNameLinks_2024(54) = Array("Q_GK_104_2024_NO_Sp", "Q13")
  varNameLinks_2024(55) = Array("Q_GK_105_2024", "Q14")
  varNameLinks_2024(56) = Array("Q_GK_105_2024_NO_Sp", "Q14")
  varNameLinks_2024(57) = Array("Q_GK_106_2024", "Q106")
  varNameLinks_2024(58) = Array("Q_GK_106_2024_NO_Sp", "Q106")
  varNameLinks_2024(59) = Array("Q_KP_107_2024", "Q15")
  varNameLinks_2024(60) = Array("Q_KP_108_2024", "Q16")
  varNameLinks_2024(61) = Array("Q_KP_109_2024", "Q17")
  varNameLinks_2024(62) = Array("Q_KP_110_2024", "Q18")
  varNameLinks_2024(63) = Array("Q_RL_26339_2024", "Q40")
  varNameLinks_2024(64) = Array("Q_RL_26369_2024", "Q46")
  varNameLinks_2024(65) = Array("Q_RL_26370_2024", "Q47")
  varNameLinks_2024(66) = Array("Q_RL_30721_2024", "Q69")
  varNameLinks_2024(67) = Array("Q_RL_30722_2024", "Q70")
  varNameLinks_2024(68) = Array("Q_RL_30723_2024", "Q71")
  varNameLinks_2024(69) = Array("Q_RL_30724_2024", "Q72")
  varNameLinks_2024(70) = Array("Q_RL_30725_2024", "Q73")
  varNameLinks_2024(71) = Array("Q_RL_30726_2024", "Q74")
  varNameLinks_2024(72) = Array("Q_RL_30728_2024", "Q76")
  varNameLinks_2024(73) = Array("Q_RL_30729_2024", "Q77")
  varNameLinks_2024(74) = Array("Q_RL_30730_2024", "Q78")
  varNameLinks_2024(75) = Array("Q_RT_10_2024", "Q6")
  varNameLinks_2024(76) = Array("Q_RT_30701_2024", "Q53")
  varNameLinks_2024(77) = Array("Q_RT_30702_2024", "Q54")
  varNameLinks_2024(78) = Array("Q_RT_30703_2024", "Q55")
  varNameLinks_2024(79) = Array("Q_RT_30705_2024", "Q57")
  varNameLinks_2024(80) = Array("Q_RT_30706_2024", "Q58")
  varNameLinks_2024(81) = Array("Q_RT_30707_2024", "Q59")
  varNameLinks_2024(82) = Array("Q_RT_30709_2024", "Q60")
  varNameLinks_2024(83) = Array("Q_SI_115_2024", "Q20")
  varNameLinks_2024(84) = Array("Q_SI_120_2024", "Q22")
  varNameLinks_2024(85) = Array("Q_SI_121_2024_NO_Sp", "Q23")
  varNameLinks_2024(86) = Array("Q_SI_121_2024_NO_Spp", "Q23")
  varNameLinks_2024(87) = Array("Q_WB_114_2024_NO_SP", "Q19")
  varNameLinks_2024(88) = Array("Q_WB_119_2024", "Q21")
  varNameLinks_2024(89) = Array("Q_WB_122_2024", "Q24")
  varNameLinks_2024(90) = Array("Q_WB_123_2024", "Q25")
  varNameLinks_2024(91) = Array("Q_WB_124_2024", "Q26")
  varNameLinks_2024(92) = Array("Q_WB_125_2024", "Q27")
  varNameLinks_2024(93) = Array("Q_WB_29003_2024", "Q48")
  varNameLinks_2024(94) = Array("Q_WB_29004_2024", "Q49")
  varNameLinks_2024(95) = Array("Q_WB_29016_2024", "Q50")
  varNameLinks_2024(96) = Array("Q_WB_29017_2024", "Q51")
  varNameLinks_2024(97) = Array("Q_WB_29025_2024", "Q52")
  varNameLinks_2024(98) = Array("Q_W_1_2024", "Q1")
  varNameLinks_2024(99) = Array("Q_W_2_2024", "Q2")
  varNameLinks_2024(100) = Array("Q_W_3_2024_", "Q3")
  varNameLinks_2024(101) = Array("Q_W_3_2024", "Q3")
  varNameLinks_2024(102) = Array("Q_W_4_2024", "Q4")

  varNameLinks_2024(103) = Array("Q_WB_114_2024", "Q19")
  varNameLinks_2024(104) = Array("Q_SI_121_2024", "Q23")
  varNameLinks_2024(105) = Array("Q_FV_22241_2024", "Q36")

  varNameLinks_2024(106) = Array("Q_FP_30738_2024", "Q86") ' HASN'T BEEN MEASURED FOR SEVERAL YEARS
  varNameLinks_2024(107) = Array("Q_BS_2004_46_2024", "Q9")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2024 = New Collection
  Set p2024toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2024)
    varSubArray = varNameLinks_2024(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2024, CStr(varSubArray(1))) Then
      pOldTo2024.Add varSubArray(0), varSubArray(1)
    End If
    p2024toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2023(varNameLinks_2023() As Variant, p2023toOld As Collection, pOldTo2023 As Collection)

  ReDim varNameLinks_2023(105)

  varNameLinks_2023(0) = Array("Q_BF_11999_2023", "Q28")
  varNameLinks_2023(1) = Array("Q_BF_12000_2023", "Q29")
  varNameLinks_2023(2) = Array("Q_BF_12000_2023_NO_Sp", "Q29")
  varNameLinks_2023(3) = Array("Q_BF_16_2023", "Q7")
  varNameLinks_2023(4) = Array("Q_BF_30713_2023", "Q63")
  varNameLinks_2023(5) = Array("Q_BF_30714_2023", "Q64")
  varNameLinks_2023(6) = Array("Q_BF_30714_2023_NO_Sp", "Q64")
  varNameLinks_2023(7) = Array("Q_BF_30715_2023", "Q65")
  varNameLinks_2023(8) = Array("Q_BF_30715_2023_NO_Sp", "Q65")
  varNameLinks_2023(9) = Array("Q_BF_30717_2023", "Q66")
  varNameLinks_2023(10) = Array("Q_BF_30719_2023", "Q67")
  varNameLinks_2023(11) = Array("Q_BF_30720_2023", "Q68")
  varNameLinks_2023(12) = Array("Q_BS_26345_2023", "Q42")
  varNameLinks_2023(13) = Array("Q_BS_26346_2023", "Q43")
  varNameLinks_2023(14) = Array("Q_BS_26347_2023", "Q44")
  varNameLinks_2023(15) = Array("Q_BS_26347_2023_NO_SP", "Q44")
  varNameLinks_2023(16) = Array("Q_BS_26348_2023", "Q45")
  varNameLinks_2023(17) = Array("Q_BS_30741_2023", "Q89")
  varNameLinks_2023(18) = Array("Q_BS_30742_2023", "Q90")
  varNameLinks_2023(19) = Array("Q_BS_30743_2023", "Q91")
  varNameLinks_2023(20) = Array("Q_BS_30744_2023", "Q92")
  varNameLinks_2023(21) = Array("Q_BS_30747_2023", "Q95")
  varNameLinks_2023(22) = Array("Q_BS_30748_2023", "Q96")
  varNameLinks_2023(23) = Array("Q_BS_30749_2023", "Q97")
  varNameLinks_2023(24) = Array("Q_BS_46_2023", "Q9")
  varNameLinks_2023(25) = Array("Q_FP_30731_2023", "Q79")
  varNameLinks_2023(26) = Array("Q_FP_30732_2023", "Q80")
  varNameLinks_2023(27) = Array("Q_FP_30733_2023", "Q81")
  varNameLinks_2023(28) = Array("Q_FP_30734_2023", "Q82")
  varNameLinks_2023(29) = Array("Q_FP_30735_2023", "Q83")
  varNameLinks_2023(30) = Array("Q_FP_30739_2023", "Q87")
  varNameLinks_2023(31) = Array("Q_FP_30740_2023", "Q88")
  varNameLinks_2023(32) = Array("Q_FS9009H_494_2023", "Q494")
  varNameLinks_2023(33) = Array("Q_FS9009H_498_2023", "Q498")
  varNameLinks_2023(34) = Array("Q_FV_21114_2023", "Q30")
  varNameLinks_2023(35) = Array("Q_FV_21114_2023_NO_Sp", "Q30")
  varNameLinks_2023(36) = Array("Q_FV_21174_2023", "Q31")
  varNameLinks_2023(37) = Array("Q_FV_21262_2023", "Q32")
  varNameLinks_2023(38) = Array("Q_FV_21262_2023_NO_Sp", "Q32")
  varNameLinks_2023(39) = Array("Q_FV_21269_2023", "Q33")
  varNameLinks_2023(40) = Array("Q_FV_22126_2023", "Q34")
  varNameLinks_2023(41) = Array("Q_FV_22156_2023", "Q35")
  varNameLinks_2023(42) = Array("Q_FV_22241_2023_NO_Sp", "Q36")
  varNameLinks_2023(43) = Array("Q_FV_22244_2023", "Q37")
  varNameLinks_2023(44) = Array("Q_FV_23155_2023", "Q38")
  varNameLinks_2023(45) = Array("Q_FV_23155_2023_NO_Sp", "Q38")
  varNameLinks_2023(46) = Array("Q_FV_23159_2023", "Q39")
  varNameLinks_2023(47) = Array("Q_FV_23159_2023_NO_Sp", "Q39")
  varNameLinks_2023(48) = Array("Q_GK_101_2023", "Q10")
  varNameLinks_2023(49) = Array("Q_GK_102_2023", "Q11")
  varNameLinks_2023(50) = Array("Q_GK_102_2023_NO_Sp", "Q11")
  varNameLinks_2023(51) = Array("Q_GK_103_2023", "Q12")
  varNameLinks_2023(52) = Array("Q_GK_103_2023_NO_Sp", "Q12")
  varNameLinks_2023(53) = Array("Q_GK_104_2023", "Q13")
  varNameLinks_2023(54) = Array("Q_GK_104_2023_NO_Sp", "Q13")
  varNameLinks_2023(55) = Array("Q_GK_105_2023", "Q14")
  varNameLinks_2023(56) = Array("Q_GK_105_2023_NO_Sp", "Q14")
  varNameLinks_2023(57) = Array("Q_GK_106_2023", "Q106")
  varNameLinks_2023(58) = Array("Q_GK_106_2023_NO_Sp", "Q106")
  varNameLinks_2023(59) = Array("Q_KP_107_2023", "Q15")
  varNameLinks_2023(60) = Array("Q_KP_108_2023", "Q16")
  varNameLinks_2023(61) = Array("Q_KP_109_2023", "Q17")
  varNameLinks_2023(62) = Array("Q_KP_110_2023", "Q18")
  varNameLinks_2023(63) = Array("Q_RL_26339_2023", "Q40")
  varNameLinks_2023(64) = Array("Q_RL_26369_2023", "Q46")
  varNameLinks_2023(65) = Array("Q_RL_26370_2023", "Q47")
  varNameLinks_2023(66) = Array("Q_RL_30721_2023", "Q69")
  varNameLinks_2023(67) = Array("Q_RL_30722_2023", "Q70")
  varNameLinks_2023(68) = Array("Q_RL_30723_2023", "Q71")
  varNameLinks_2023(69) = Array("Q_RL_30724_2023", "Q72")
  varNameLinks_2023(70) = Array("Q_RL_30725_2023", "Q73")
  varNameLinks_2023(71) = Array("Q_RL_30726_2023", "Q74")
  varNameLinks_2023(72) = Array("Q_RL_30728_2023", "Q76")
  varNameLinks_2023(73) = Array("Q_RL_30729_2023", "Q77")
  varNameLinks_2023(74) = Array("Q_RL_30730_2023", "Q78")
  varNameLinks_2023(75) = Array("Q_RT_10_2023", "Q6")
  varNameLinks_2023(76) = Array("Q_RT_30701_2023", "Q53")
  varNameLinks_2023(77) = Array("Q_RT_30702_2023", "Q54")
  varNameLinks_2023(78) = Array("Q_RT_30703_2023", "Q55")
  varNameLinks_2023(79) = Array("Q_RT_30705_2023", "Q57")
  varNameLinks_2023(80) = Array("Q_RT_30706_2023", "Q58")
  varNameLinks_2023(81) = Array("Q_RT_30707_2023", "Q59")
  varNameLinks_2023(82) = Array("Q_RT_30709_2023", "Q60")
  varNameLinks_2023(83) = Array("Q_SI_115_2023", "Q20")
  varNameLinks_2023(84) = Array("Q_SI_120_2023", "Q22")
  varNameLinks_2023(85) = Array("Q_SI_121_2023_NO_Sp", "Q23")
  varNameLinks_2023(86) = Array("Q_SI_121_2023_NO_Spp", "Q23")
  varNameLinks_2023(87) = Array("Q_WB_114_2023_NO_SP", "Q19")
  varNameLinks_2023(88) = Array("Q_WB_119_2023", "Q21")
  varNameLinks_2023(89) = Array("Q_WB_122_2023", "Q24")
  varNameLinks_2023(90) = Array("Q_WB_123_2023", "Q25")
  varNameLinks_2023(91) = Array("Q_WB_124_2023", "Q26")
  varNameLinks_2023(92) = Array("Q_WB_125_2023", "Q27")
  varNameLinks_2023(93) = Array("Q_WB_29003_2023", "Q48")
  varNameLinks_2023(94) = Array("Q_WB_29004_2023", "Q49")
  varNameLinks_2023(95) = Array("Q_WB_29016_2023", "Q50")
  varNameLinks_2023(96) = Array("Q_WB_29017_2023", "Q51")
  varNameLinks_2023(97) = Array("Q_WB_29025_2023", "Q52")
  varNameLinks_2023(98) = Array("Q_W_1_2023", "Q1")
  varNameLinks_2023(99) = Array("Q_W_2_2023", "Q2")
  varNameLinks_2023(100) = Array("Q_W_3_2023_", "Q3")
  varNameLinks_2023(101) = Array("Q_W_3_2023", "Q3")
  varNameLinks_2023(102) = Array("Q_W_4_2023", "Q4")

  varNameLinks_2023(103) = Array("Q_WB_114_2023", "Q19")
  varNameLinks_2023(104) = Array("Q_SI_121_2023", "Q23")
  varNameLinks_2023(105) = Array("Q_FV_22241_2023", "Q36")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2023 = New Collection
  Set p2023toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2023)
    varSubArray = varNameLinks_2023(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2023, CStr(varSubArray(1))) Then
      pOldTo2023.Add varSubArray(0), varSubArray(1)
    End If
    p2023toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2022(varNameLinks_2022() As Variant, p2022toOld As Collection, pOldTo2022 As Collection)

  ReDim varNameLinks_2022(92)

  varNameLinks_2022(0) = Array("Q_BS_26345_2022", "Q42")
  varNameLinks_2022(1) = Array("Q_FP_30731_2022", "Q79")
  varNameLinks_2022(2) = Array("Q_FP_30732_2022", "Q80")
  varNameLinks_2022(3) = Array("Q_FP_30733_2022", "Q81")
  varNameLinks_2022(4) = Array("Q_FP_30734_2022", "Q82")
  varNameLinks_2022(5) = Array("Q_FP_30735_2022", "Q83")
  varNameLinks_2022(6) = Array("Q_FP_30739_2022", "Q87")
  varNameLinks_2022(7) = Array("Q_FP_30740_2022", "Q88")
  varNameLinks_2022(8) = Array("Q_FV_22126_2022", "Q34")
  varNameLinks_2022(9) = Array("Q_GK_101_2022", "Q10")
  varNameLinks_2022(10) = Array("Q_GK_102_2022", "Q11")
  varNameLinks_2022(11) = Array("Q_GK_103_2022", "Q12")
  varNameLinks_2022(12) = Array("Q_GK_104_2022", "Q13")
  varNameLinks_2022(13) = Array("Q_GK_105_2022", "Q14")
  varNameLinks_2022(14) = Array("Q_GK_106_2022", "Q106")
  varNameLinks_2022(15) = Array("Q_KP_109_2022", "Q17")
  varNameLinks_2022(16) = Array("Q_KP_110_2022", "Q18")
  varNameLinks_2022(17) = Array("Q_BF_11999_2022", "Q28")
  varNameLinks_2022(18) = Array("Q_BF_12000_2022", "Q29")
  varNameLinks_2022(19) = Array("Q_BF_200416_2022", "Q7")  ' CHANGED FROM PREVIOUS YEAR
  varNameLinks_2022(20) = Array("Q_BF_30713_2022", "Q63")
  varNameLinks_2022(21) = Array("Q_BF_30714_2022", "Q64")
  varNameLinks_2022(22) = Array("Q_BF_30715_2022", "Q65")
  varNameLinks_2022(23) = Array("Q_BF_30717_2022", "Q66")
  varNameLinks_2022(24) = Array("Q_BF_30719_2022", "Q67")
  varNameLinks_2022(25) = Array("Q_BF_30720_2022", "Q68")
  varNameLinks_2022(26) = Array("Q_BS_200446_2022", "Q9")
  varNameLinks_2022(27) = Array("Q_BS_26346_2022", "Q43")
  varNameLinks_2022(28) = Array("Q_BS_26347_2022", "Q44")
  varNameLinks_2022(29) = Array("Q_BS_26348_2022", "Q45")
  varNameLinks_2022(30) = Array("Q_BS_30741_2022", "Q89")
  varNameLinks_2022(31) = Array("Q_BS_30742_2022", "Q90")
  varNameLinks_2022(32) = Array("Q_BS_30743_2022", "Q91")
  varNameLinks_2022(33) = Array("Q_BS_30744_2022", "Q92")
  varNameLinks_2022(34) = Array("Q_BS_30747_2022", "Q95")
  varNameLinks_2022(35) = Array("Q_BS_30748_2022", "Q96")
  varNameLinks_2022(36) = Array("Q_BS_30749_2022", "Q97")
  varNameLinks_2022(37) = Array("Q_9009H_494_2022", "Q494")
  varNameLinks_2022(38) = Array("Q_9009H_498_2022", "Q498")
  varNameLinks_2022(39) = Array("Q_FV 21262_2022", "Q32")
  varNameLinks_2022(40) = Array("Q_FV_21114_2022", "Q30")
  varNameLinks_2022(41) = Array("Q_FV_21114_2022_NO_Sp", "Q30")
  varNameLinks_2022(42) = Array("Q_FV_21174_2022", "Q31")
  varNameLinks_2022(43) = Array("Q_FV_22156_2022", "Q35")
  varNameLinks_2022(44) = Array("Q_FV_21262_2022", "Q32")
  varNameLinks_2022(45) = Array("Q_FV_21269_2022", "Q33")
  varNameLinks_2022(46) = Array("Q_FV_22244_2022", "Q37")
  varNameLinks_2022(47) = Array("Q_FV_23155_2022", "Q38")
  varNameLinks_2022(48) = Array("Q_FV_23155_2022_NO_Sp", "Q38")
  varNameLinks_2022(49) = Array("Q_FV_23159_2022", "Q39")
  varNameLinks_2022(50) = Array("Q_FV_23159_2022_NO_Sp", "Q39")
  varNameLinks_2022(51) = Array("Q_KP_107_2022", "Q15")
  varNameLinks_2022(52) = Array("Q_KP_108_2022", "Q16")
  varNameLinks_2022(53) = Array("Q_RL_30723_2022", "Q71")
  varNameLinks_2022(54) = Array("Q_RL_30724_2022", "Q72")
  varNameLinks_2022(55) = Array("Q_RL_30725_2022", "Q73")
  varNameLinks_2022(56) = Array("Q_RL_30726_2022", "Q74")
  varNameLinks_2022(57) = Array("Q_RL_30729_2022", "Q77")
  varNameLinks_2022(58) = Array("Q_RL_30730_2022", "Q78")
  varNameLinks_2022(59) = Array("Q_RT_200410_2022", "Q6")
  varNameLinks_2022(60) = Array("Q_RT_30701_2022", "Q53")
  varNameLinks_2022(61) = Array("Q_RT_30702_2022", "Q54")
  varNameLinks_2022(62) = Array("Q_RT_30703_2022", "Q55")
  varNameLinks_2022(63) = Array("Q_RT_30705_2022", "Q57")
  varNameLinks_2022(64) = Array("Q_RT_30706_2022", "Q58")
  varNameLinks_2022(65) = Array("Q_RT_30707_2022", "Q59")
  varNameLinks_2022(66) = Array("Q_RT_30709_2022", "Q60")
  varNameLinks_2022(67) = Array("Q_SI_115_2022", "Q20")
  varNameLinks_2022(68) = Array("Q_SI_120_2022", "Q22")
  varNameLinks_2022(69) = Array("Q_SI_121_2022", "Q23")
  varNameLinks_2022(70) = Array("Q_WB_114_2022", "Q19")
  varNameLinks_2022(71) = Array("Q_WB_119_2022", "Q21")
  varNameLinks_2022(72) = Array("Q_Dispersed_122_2022", "Q24")
  varNameLinks_2022(73) = Array("Q_Dispersed_123_2022", "Q25")
  varNameLinks_2022(74) = Array("Q_Dispersed_124_2022", "Q26")
  varNameLinks_2022(75) = Array("Q_Dispersed_125_2022", "Q27")
  varNameLinks_2022(76) = Array("Q_Dispersed_29003_2022", "Q48")
  varNameLinks_2022(77) = Array("Q_Dispersed_29004_2022", "Q49")
  varNameLinks_2022(78) = Array("Q_Dispersed_29016_2022", "Q50")
  varNameLinks_2022(79) = Array("Q_Dispersed_29017_2022", "Q51")
  varNameLinks_2022(80) = Array("Q_WB_29017_2022", "Q51")
  varNameLinks_2022(81) = Array("Q_Dispersed_29025_2022", "Q52")
  varNameLinks_2022(82) = Array("Q_RL_26339_2022", "Q40")
  varNameLinks_2022(83) = Array("Q_RL_26369_2022", "Q46")
  varNameLinks_2022(84) = Array("Q_RL_26370_2022", "Q47")
  varNameLinks_2022(85) = Array("Q_RL_30721_2022", "Q69")
  varNameLinks_2022(86) = Array("Q_RL_30722_2022", "Q70")
  varNameLinks_2022(87) = Array("Q_RL_30728_2022", "Q76")
  varNameLinks_2022(88) = Array("Q_Willaha_1_2022", "Q1")
  varNameLinks_2022(89) = Array("Q_Willaha_2_2022", "Q2")
  varNameLinks_2022(90) = Array("Q_Willaha_3_2022", "Q3")
  varNameLinks_2022(91) = Array("Q_Willaha_4_2022", "Q4")
  varNameLinks_2022(92) = Array("Q_FV_22241_2022", "Q36")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2022 = New Collection
  Set p2022toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2022)
    varSubArray = varNameLinks_2022(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2022, CStr(varSubArray(1))) Then
      pOldTo2022.Add varSubArray(0), varSubArray(1)
    End If
    p2022toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2021(varNameLinks_2021() As Variant, p2021toOld As Collection, pOldTo2021 As Collection)

  ReDim varNameLinks_2021(93)

  varNameLinks_2021(0) = Array("Q_BS_26345_2021", "Q42")
  varNameLinks_2021(1) = Array("Q_FP_30731_2021", "Q79")
  varNameLinks_2021(2) = Array("Q_FP_30732_2021", "Q80")
  varNameLinks_2021(3) = Array("Q_FP_30733_2021", "Q81")
  varNameLinks_2021(4) = Array("Q_FP_30734_2021", "Q82")
  varNameLinks_2021(5) = Array("Q_FP_30735_2021", "Q83")
  varNameLinks_2021(6) = Array("Q_FP_30739_2021", "Q87")
  varNameLinks_2021(7) = Array("Q_FP_30740_2021", "Q88")
  varNameLinks_2021(8) = Array("Q_FV_22126_2021", "Q34")
  varNameLinks_2021(9) = Array("Q_GK_101_2021", "Q10")
  varNameLinks_2021(10) = Array("Q_GK_102_2021", "Q11")
  varNameLinks_2021(11) = Array("Q_GK_103_2021", "Q12")
  varNameLinks_2021(12) = Array("Q_GK_104_2021", "Q13")
  varNameLinks_2021(13) = Array("Q_GK_105_2021", "Q14")
  varNameLinks_2021(14) = Array("Q_GK_106_2021", "Q106")
  varNameLinks_2021(15) = Array("Q_KP_109_2021", "Q17")
  varNameLinks_2021(16) = Array("Q_KP_110_2021", "Q18")
  varNameLinks_2021(17) = Array("Q_BF_11999_2021", "Q28")
  varNameLinks_2021(18) = Array("Q_BF_12000_2021", "Q29")
  varNameLinks_2021(19) = Array("Q_BF_2004-16_2021", "Q7")
  varNameLinks_2021(20) = Array("Q_BF_30713_2021", "Q63")
  varNameLinks_2021(21) = Array("Q_BF_30714_2021", "Q64")
  varNameLinks_2021(22) = Array("Q_BF_30715_2021", "Q65")
  varNameLinks_2021(23) = Array("Q_BF_30717_2021", "Q66")
  varNameLinks_2021(24) = Array("Q_BF_30719_2021", "Q67")
  varNameLinks_2021(25) = Array("Q_BF_30720_2021", "Q68")
  varNameLinks_2021(26) = Array("Q_BS_2004-46_2021", "Q9")
  varNameLinks_2021(27) = Array("Q_BS_26346_2021", "Q43")
  varNameLinks_2021(28) = Array("Q_BS_26347_2021", "Q44")
  varNameLinks_2021(29) = Array("Q_BS_26348_2021", "Q45")
  varNameLinks_2021(30) = Array("Q_BS_30741_2021", "Q89")
  varNameLinks_2021(31) = Array("Q_BS_30742_2021", "Q90")
  varNameLinks_2021(32) = Array("Q_BS_30743_2021", "Q91")
  varNameLinks_2021(33) = Array("Q_BS_30744_2021", "Q92")
  varNameLinks_2021(34) = Array("Q_BS_30747_2021", "Q95")
  varNameLinks_2021(35) = Array("Q_BS_30748_2021", "Q96")
  varNameLinks_2021(36) = Array("Q_BS_30749_2021", "Q97")
  varNameLinks_2021(37) = Array("Q_9009H_494_2021", "Q494")
  varNameLinks_2021(38) = Array("Q_9009H_498_2021", "Q498")
  varNameLinks_2021(39) = Array("Q_FV 21262_2021", "Q32")
  varNameLinks_2021(40) = Array("Q_FV_21114_2021", "Q30")
  varNameLinks_2021(41) = Array("Q_FV_21114_2021_NO_Sp", "Q30")
  varNameLinks_2021(42) = Array("Q_FV_21174_2021", "Q31")
  varNameLinks_2021(43) = Array("Q_FV_22156_2021", "Q35")
  varNameLinks_2021(44) = Array("Q_FV_21262_2021", "Q32")
  varNameLinks_2021(45) = Array("Q_FV_21269_2021", "Q33")
  varNameLinks_2021(46) = Array("Q_FV_22244_2021", "Q37")
  varNameLinks_2021(47) = Array("Q_FV_23155_2021", "Q38")
  varNameLinks_2021(48) = Array("Q_FV_23155_2021_NO_Sp", "Q38")
  varNameLinks_2021(49) = Array("Q_FV_23159_2021", "Q39")
  varNameLinks_2021(50) = Array("Q_FV_23159_2021_NO_Sp", "Q39")
  varNameLinks_2021(51) = Array("Q_KP_107_2021", "Q15")
  varNameLinks_2021(52) = Array("Q_KP_108_2021", "Q16")
  varNameLinks_2021(53) = Array("Q_RL_30723_2021", "Q71")
  varNameLinks_2021(54) = Array("Q_RL_30724_2021", "Q72")
  varNameLinks_2021(55) = Array("Q_RL_30725_2021", "Q73")
  varNameLinks_2021(56) = Array("Q_RL_30726_2021", "Q74")
  varNameLinks_2021(57) = Array("Q_RL_30729_2021", "Q77")
  varNameLinks_2021(58) = Array("Q_RL_30730_2021", "Q78")
  varNameLinks_2021(59) = Array("Q_RT_2004-10_2021", "Q6")
  varNameLinks_2021(60) = Array("Q_RT_30701_2021", "Q53")
  varNameLinks_2021(61) = Array("Q_RT_30702_2021", "Q54")
  varNameLinks_2021(62) = Array("Q_RT_30703_2021", "Q55")
  varNameLinks_2021(63) = Array("Q_RT_30705_2021", "Q57")
  varNameLinks_2021(64) = Array("Q_RT_30706_2021", "Q58")
  varNameLinks_2021(65) = Array("Q_RT_30707_2021", "Q59")
  varNameLinks_2021(66) = Array("Q_RT_30709_2021", "Q60")
  varNameLinks_2021(67) = Array("Q_SI_115_2021", "Q20")
  varNameLinks_2021(68) = Array("Q_SI_120_2021", "Q22")
  varNameLinks_2021(69) = Array("Q_SI_121_2021", "Q23")
  varNameLinks_2021(70) = Array("Q_WB_114_2021", "Q19")
  varNameLinks_2021(71) = Array("Q_WB_119_2021", "Q21")
  varNameLinks_2021(72) = Array("Q_WB_122_2021", "Q24")
  varNameLinks_2021(73) = Array("Q_WB_123_2021", "Q25")
  varNameLinks_2021(74) = Array("Q_WB_124_2021", "Q26")
  varNameLinks_2021(75) = Array("Q_WB_125_2021", "Q27")
  varNameLinks_2021(76) = Array("Q_WB_29003_2021", "Q48")
  varNameLinks_2021(77) = Array("Q_WB_29004_2021", "Q49")
  varNameLinks_2021(78) = Array("Q_WB_29016_2021", "Q50")
  varNameLinks_2021(79) = Array("Q_Dispersed_29017_2021", "Q51")
  varNameLinks_2021(80) = Array("Q_WB_29017_2021", "Q51")
  varNameLinks_2021(81) = Array("Q_WB_29025_2021", "Q52")
  varNameLinks_2021(82) = Array("Q_RL_26339_2021", "Q40")
  varNameLinks_2021(83) = Array("Q_RL_26369_2021", "Q46")
  varNameLinks_2021(84) = Array("Q_RL_26370_2021", "Q47")
  varNameLinks_2021(85) = Array("Q_RL_30721_2021", "Q69")
  varNameLinks_2021(86) = Array("Q_RL_30722_2021", "Q70")
  varNameLinks_2021(87) = Array("Q_RL_30728_2021", "Q76")
  varNameLinks_2021(88) = Array("Q_Willaha_1_2021", "Q1")
  varNameLinks_2021(89) = Array("Q_Willaha_2_2021", "Q2")
  varNameLinks_2021(90) = Array("Q_Willaha_3_2021", "Q3")
  varNameLinks_2021(91) = Array("Q_Willaha_4_2021", "Q4")
  varNameLinks_2021(92) = Array("Q_FV_22241_2021_NOSPECS", "Q36")
  varNameLinks_2021(93) = Array("Q_FV_22241_2021", "Q36")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2021 = New Collection
  Set p2021toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2021)
    varSubArray = varNameLinks_2021(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2021, CStr(varSubArray(1))) Then
      pOldTo2021.Add varSubArray(0), varSubArray(1)
    End If
    p2021toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2020(varNameLinks_2020() As Variant, p2020toOld As Collection, pOldTo2020 As Collection)

  ReDim varNameLinks_2020(92)

  varNameLinks_2020(0) = Array("Q_BS_26345_2020", "Q42")
  varNameLinks_2020(1) = Array("Q_FP_30731_2020", "Q79")
  varNameLinks_2020(2) = Array("Q_FP_30732_2020", "Q80")
  varNameLinks_2020(3) = Array("Q_FP_30733_2020", "Q81")
  varNameLinks_2020(4) = Array("Q_FP_30734_2020", "Q82")
  varNameLinks_2020(5) = Array("Q_FP_30735_2020", "Q83")
  varNameLinks_2020(6) = Array("Q_FP_30739_2020", "Q87")
  varNameLinks_2020(7) = Array("Q_FP_30740_2020", "Q88")
  varNameLinks_2020(8) = Array("Q_FV_22126_2020", "Q34")
  varNameLinks_2020(9) = Array("Q_GK_101_2020", "Q10")
  varNameLinks_2020(10) = Array("Q_GK_102_2020", "Q11")
  varNameLinks_2020(11) = Array("Q_GK_103_2020", "Q12")
  varNameLinks_2020(12) = Array("Q_GK_104_2020", "Q13")
  varNameLinks_2020(13) = Array("Q_GK_105_2020", "Q14")
  varNameLinks_2020(14) = Array("Q_GK_106_2020", "Q106")
  varNameLinks_2020(15) = Array("Q_Kendrick_109_2020", "Q17")
  varNameLinks_2020(16) = Array("Q_Kendrick_110_2020", "Q18")
  varNameLinks_2020(17) = Array("Q_BF_11999_2020", "Q28")
  varNameLinks_2020(18) = Array("Q_BF_12000_2020", "Q29")
  varNameLinks_2020(19) = Array("Q_BF_2004-16_2020", "Q7")
  varNameLinks_2020(20) = Array("Q_BF_30713_2020", "Q63")
  varNameLinks_2020(21) = Array("Q_BF_30714_2020", "Q64")
  varNameLinks_2020(22) = Array("Q_BF_30715_2020", "Q65")
  varNameLinks_2020(23) = Array("Q_BF_30717_2020", "Q66")
  varNameLinks_2020(24) = Array("Q_BF_30719_2020", "Q67")
  varNameLinks_2020(25) = Array("Q_BF_30720_2020", "Q68")
  varNameLinks_2020(26) = Array("Q_BS_200446_2020", "Q9")
  varNameLinks_2020(27) = Array("Q_BS_26346_2020", "Q43")
  varNameLinks_2020(28) = Array("Q_BS_26347_2020", "Q44")
  varNameLinks_2020(29) = Array("Q_BS_26348_2020", "Q45")
  varNameLinks_2020(30) = Array("Q_BS_30741_2020", "Q89")
  varNameLinks_2020(31) = Array("Q_BS_30742_2020", "Q90")
  varNameLinks_2020(32) = Array("Q_BS_30743_2020", "Q91")
  varNameLinks_2020(33) = Array("Q_BS_30744_2020", "Q92")
  varNameLinks_2020(34) = Array("Q_BS_30747_2020", "Q95")
  varNameLinks_2020(35) = Array("Q_BS_30748_2020", "Q96")
  varNameLinks_2020(36) = Array("Q_BS_30749_2020", "Q97")
  varNameLinks_2020(37) = Array("Q_9009H_494_2020", "Q494")
  varNameLinks_2020(38) = Array("Q_9009H_498_2020", "Q498")
  varNameLinks_2020(39) = Array("Q_FV 21262_2020", "Q32")
  varNameLinks_2020(40) = Array("Q_FV_21114_2020", "Q30")
  varNameLinks_2020(41) = Array("Q_FV_21114_2020_NO_Sp", "Q30")
  varNameLinks_2020(42) = Array("Q_FV_21174_2020", "Q31")
  varNameLinks_2020(43) = Array("Q_FV_22156_2020", "Q35")
  varNameLinks_2020(44) = Array("Q_FV_21262_2020", "Q32")
  varNameLinks_2020(45) = Array("Q_FV_21269_2020", "Q33")
  varNameLinks_2020(46) = Array("Q_FV_22244_2020", "Q37")
  varNameLinks_2020(47) = Array("Q_FV_23155_2020", "Q38")
  varNameLinks_2020(48) = Array("Q_FV_23155_2020_NO_Sp", "Q38")
  varNameLinks_2020(49) = Array("Q_FV_23159_2020", "Q39")
  varNameLinks_2020(50) = Array("Q_FV_23159_2020_NO_Sp", "Q39")
  varNameLinks_2020(51) = Array("Q_Kendrick_107_2020", "Q15")
  varNameLinks_2020(52) = Array("Q_Kendrick_108_2020", "Q16")
  varNameLinks_2020(53) = Array("Q_RL_30723_2020", "Q71")
  varNameLinks_2020(54) = Array("Q_RL_30724_2020", "Q72")
  varNameLinks_2020(55) = Array("Q_RL_30725_2020", "Q73")
  varNameLinks_2020(56) = Array("Q_RL_30726_2020", "Q74")
  varNameLinks_2020(57) = Array("Q_RL_30729_2020", "Q77")
  varNameLinks_2020(58) = Array("Q_RL_30730_2020", "Q78")
  varNameLinks_2020(59) = Array("Q_RT_2004-10_2020", "Q6")
  varNameLinks_2020(60) = Array("Q_RT_30701_2020", "Q53")
  varNameLinks_2020(61) = Array("Q_RT_30702_2020", "Q54")
  varNameLinks_2020(62) = Array("Q_RT_30703_2020", "Q55")
  varNameLinks_2020(63) = Array("Q_RT_30705_2020", "Q57")
  varNameLinks_2020(64) = Array("Q_RT_30706_2020", "Q58")
  varNameLinks_2020(65) = Array("Q_RT_30707_2020", "Q59")
  varNameLinks_2020(66) = Array("Q_RT_30709_2020", "Q60")
  varNameLinks_2020(67) = Array("Q_SI_115_2020", "Q20")
  varNameLinks_2020(68) = Array("Q_SI_120_2020", "Q22")
  varNameLinks_2020(69) = Array("Q_SI_121_2020", "Q23")
  varNameLinks_2020(70) = Array("Q_WB_114_2020_NoSpec", "Q19")
  varNameLinks_2020(71) = Array("Q_WB_119_2020", "Q21")
  varNameLinks_2020(72) = Array("Q_Dispersed_122_2020", "Q24")
  varNameLinks_2020(73) = Array("Q_Dispersed_123_2020", "Q25")
  varNameLinks_2020(74) = Array("Q_Dispersed_124_2020", "Q26")
  varNameLinks_2020(75) = Array("Q_Dispersed_125_2020", "Q27")
  varNameLinks_2020(76) = Array("Q_Dispersed_29003_2020", "Q48")
  varNameLinks_2020(77) = Array("Q_Dispersed_29004_2020", "Q49")
  varNameLinks_2020(78) = Array("Q_Dispersed_29016_2020", "Q50")
  varNameLinks_2020(79) = Array("Q_Dispersed_29017_2020", "Q51")
  varNameLinks_2020(80) = Array("Q_WB_29017_2020", "Q51")
  varNameLinks_2020(81) = Array("Q_Dispersed_29025_2020", "Q52")
  varNameLinks_2020(82) = Array("Q_RL_26339_2020", "Q40")
  varNameLinks_2020(83) = Array("Q_RL_26369_2020", "Q46")
  varNameLinks_2020(84) = Array("Q_RL_26370_2020", "Q47")
  varNameLinks_2020(85) = Array("Q_RL_30721_2020", "Q69")
  varNameLinks_2020(86) = Array("Q_RL_30722_2020", "Q70")
  varNameLinks_2020(87) = Array("Q_RL_30728_2020", "Q76")
  varNameLinks_2020(88) = Array("Q_Willaha_1_2020", "Q1")
  varNameLinks_2020(89) = Array("Q_Willaha_2_2020", "Q2")
  varNameLinks_2020(90) = Array("Q_Willaha_3_2020", "Q3")
  varNameLinks_2020(91) = Array("Q_Willaha_4_2020", "Q4")
  varNameLinks_2020(92) = Array("Q_FV_22241_2020_NOSPECS", "Q36")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2020 = New Collection
  Set p2020toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2020)
    varSubArray = varNameLinks_2020(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2020, CStr(varSubArray(1))) Then
      pOldTo2020.Add varSubArray(0), varSubArray(1)
    End If
    p2020toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_Missing2021(varNameLinks_Missing2021() As Variant, _
      pMiss2021toOld As Collection, pOldToMiss2021 As Collection)

  ReDim varNameLinks_Missing2021(7)

  varNameLinks_Missing2021(0) = Array("Q_RL_30722_2016", "Q70")  ' <-- only one with plants observed
  varNameLinks_Missing2021(1) = Array("FV_22241_2019", "Q36")
  varNameLinks_Missing2021(2) = Array("Q_WB_114_2007", "Q19")
  varNameLinks_Missing2021(3) = Array("Q_WB_114_2009", "Q19")
  varNameLinks_Missing2021(4) = Array("Q_WB_114_2010", "Q19")
  varNameLinks_Missing2021(5) = Array("Q_WB_114_2011", "Q19")
  varNameLinks_Missing2021(6) = Array("Q_WB_114_2018", "Q19")
  varNameLinks_Missing2021(7) = Array("Willaha_1_2017", "Q1")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldToMiss2021 = New Collection
  Set pMiss2021toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_Missing2021)
    varSubArray = varNameLinks_Missing2021(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldToMiss2021, CStr(varSubArray(1))) Then
      pOldToMiss2021.Add varSubArray(0), varSubArray(1)
    End If
    pMiss2021toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2019(varNameLinks_2019() As Variant, p2019toOld As Collection, pOldTo2019 As Collection)

  ReDim varNameLinks_2019(93)

  varNameLinks_2019(0) = Array("BS_26345_2019", "Q42")
  varNameLinks_2019(1) = Array("FP_30731_2019", "Q79")
  varNameLinks_2019(2) = Array("FP_30732_2019", "Q80")
  varNameLinks_2019(3) = Array("FP_30733_2019", "Q81")
  varNameLinks_2019(4) = Array("FP_30734_2019", "Q82")
  varNameLinks_2019(5) = Array("FP_30735_2019", "Q83")
  varNameLinks_2019(6) = Array("FP_30739_2019", "Q87")
  varNameLinks_2019(7) = Array("FP_30740_2019", "Q88")
  varNameLinks_2019(8) = Array("FV_22126_2019", "Q34")
  varNameLinks_2019(9) = Array("GK_101_2019", "Q10")
  varNameLinks_2019(10) = Array("GK_102_2019", "Q11")
  varNameLinks_2019(11) = Array("GK_103_2019", "Q12")
  varNameLinks_2019(12) = Array("GK_104_2019", "Q13")
  varNameLinks_2019(13) = Array("GK_105_2019", "Q14")
  varNameLinks_2019(14) = Array("GK_106_2019", "Q106")
  varNameLinks_2019(15) = Array("KP_109_2019", "Q17")
  varNameLinks_2019(16) = Array("KP_110_2019", "Q18")
  varNameLinks_2019(17) = Array("Q_BF_11999_2019", "Q28")
  varNameLinks_2019(18) = Array("Q_BF_12000_2019", "Q29")
  varNameLinks_2019(19) = Array("Q_BF_2004-16_2019", "Q7")
  varNameLinks_2019(20) = Array("Q_BF_30713_2019", "Q63")
  varNameLinks_2019(21) = Array("Q_BF_30714_2019", "Q64")
  varNameLinks_2019(22) = Array("Q_BF_30715_2019", "Q65")
  varNameLinks_2019(23) = Array("Q_BF_30717_2019", "Q66")
  varNameLinks_2019(24) = Array("Q_BF_30719_2019", "Q67")
  varNameLinks_2019(25) = Array("Q_BF_30720_2019", "Q68")
  varNameLinks_2019(26) = Array("Q_BS_200446_2019", "Q9")
  varNameLinks_2019(27) = Array("Q_BS_26346_2019", "Q43")
  varNameLinks_2019(28) = Array("Q_BS_26347_2019", "Q44")
  varNameLinks_2019(29) = Array("Q_BS_26348_2019", "Q45")
  varNameLinks_2019(30) = Array("Q_BS_30741_2019", "Q89")
  varNameLinks_2019(31) = Array("Q_BS_30742_2019", "Q90")
  varNameLinks_2019(32) = Array("Q_BS_30743_2019", "Q91")
  varNameLinks_2019(33) = Array("Q_BS_30744_2019", "Q92")
  varNameLinks_2019(34) = Array("Q_BS_30747_2019", "Q95")
  varNameLinks_2019(35) = Array("Q_BS_30748_2019", "Q96")
  varNameLinks_2019(36) = Array("Q_BS_30749_2019", "Q97")
  varNameLinks_2019(37) = Array("Q_FS9009H_494_2019", "Q494")
  varNameLinks_2019(38) = Array("Q_FS9009_498_2019", "Q498")
  varNameLinks_2019(39) = Array("Q_FV 21262_2019", "Q32")
  varNameLinks_2019(40) = Array("Q_FV_21114_2019", "Q30")
  varNameLinks_2019(41) = Array("Q_FV_21114_2019_NO_Sp", "Q30")
  varNameLinks_2019(42) = Array("Q_FV_21174_2019", "Q31")
  varNameLinks_2019(43) = Array("Q_FV_21256_2019", "Q35")
  varNameLinks_2019(44) = Array("Q_FV_21262_2019", "Q32")
  varNameLinks_2019(45) = Array("Q_FV_21269_2019", "Q33")
  varNameLinks_2019(46) = Array("Q_FV_22244_2019", "Q37")
  varNameLinks_2019(47) = Array("Q_FV_23155_2019", "Q38")
  varNameLinks_2019(48) = Array("Q_FV_23155_2019_NO_Sp", "Q38")
  varNameLinks_2019(49) = Array("Q_FV_23159_2019", "Q39")
  varNameLinks_2019(50) = Array("Q_FV_23159_2019_NO_Sp", "Q39")
  varNameLinks_2019(51) = Array("Q_KP_107_2019", "Q15")
  varNameLinks_2019(52) = Array("Q_KP_108_2019", "Q16")
  varNameLinks_2019(53) = Array("Q_RL_30723_2019", "Q71")
  varNameLinks_2019(54) = Array("Q_RL_30724_2019", "Q72")
  varNameLinks_2019(55) = Array("Q_RL_30725_2019", "Q73")
  varNameLinks_2019(56) = Array("Q_RL_30726_2019", "Q74")
  varNameLinks_2019(57) = Array("Q_RL_30729_2019", "Q77")
  varNameLinks_2019(58) = Array("Q_RL_30730_2019", "Q78")
  varNameLinks_2019(59) = Array("Q_RT_2004-10_19", "Q6")
  varNameLinks_2019(60) = Array("Q_RT_30701_2019", "Q53")
  varNameLinks_2019(61) = Array("Q_RT_30702_2019", "Q54")
  varNameLinks_2019(62) = Array("Q_RT_30703_2019", "Q55")
  varNameLinks_2019(63) = Array("Q_RT_30705_2019", "Q57")
  varNameLinks_2019(64) = Array("Q_RT_30706_2019", "Q58")
  varNameLinks_2019(65) = Array("Q_RT_30707_2019", "Q59")
  varNameLinks_2019(66) = Array("Q_RT_30709_2019", "Q60")
  varNameLinks_2019(67) = Array("Q_SI_115_2019", "Q20")
  varNameLinks_2019(68) = Array("Q_SI_120_2019", "Q22")
  varNameLinks_2019(69) = Array("Q_SI_121_2019", "Q23")
  varNameLinks_2019(70) = Array("Q_WB_114_2019_DF_NO_Sp", "Q19")
  varNameLinks_2019(92) = Array("Q_WB_114_2019_CF_NO_Sp", "Q19")
  varNameLinks_2019(93) = Array("Q_WB_114_2019_NO_Sp", "Q19")
  varNameLinks_2019(71) = Array("Q_WB_119_2019", "Q21")
  varNameLinks_2019(72) = Array("Q_WB_122_2019", "Q24")
  varNameLinks_2019(73) = Array("Q_WB_123_2019", "Q25")
  varNameLinks_2019(74) = Array("Q_WB_124_2019", "Q26")
  varNameLinks_2019(75) = Array("Q_WB_125_2019", "Q27")
  varNameLinks_2019(76) = Array("Q_WB_29003_2019", "Q48")
  varNameLinks_2019(77) = Array("Q_WB_29004_2019", "Q49")
  varNameLinks_2019(78) = Array("Q_WB_29016_2019", "Q50")
  varNameLinks_2019(79) = Array("Q_WB_29017 -2019", "Q51")
  varNameLinks_2019(80) = Array("Q_WB_29017_2019", "Q51")
  varNameLinks_2019(81) = Array("Q_WB_29025_2019", "Q52")
  varNameLinks_2019(82) = Array("RL_26339_2019", "Q40")
  varNameLinks_2019(83) = Array("RL_26369_2019", "Q46")
  varNameLinks_2019(84) = Array("RL_26370_2019", "Q47")
  varNameLinks_2019(85) = Array("RL_30721_2019", "Q69")
  varNameLinks_2019(86) = Array("RL_30722_2019", "Q70")
  varNameLinks_2019(87) = Array("RL_30728_2019", "Q76")
  varNameLinks_2019(88) = Array("Willaha_1_2019", "Q1")
  varNameLinks_2019(89) = Array("Willaha_2_2019", "Q2")
  varNameLinks_2019(90) = Array("Willaha_3_2019", "Q3")
  varNameLinks_2019(91) = Array("Willaha_4_2019", "Q4")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2019 = New Collection
  Set p2019toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2019)
    varSubArray = varNameLinks_2019(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2019, CStr(varSubArray(1))) Then
      pOldTo2019.Add varSubArray(0), varSubArray(1)
    End If
    p2019toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2018(varNameLinks_2018() As Variant, p2018toOld As Collection, pOldTo2018 As Collection)

  ReDim varNameLinks_2018(87)
  varNameLinks_2018(0) = Array("QBF_11999_2018", "Q28")
  varNameLinks_2018(1) = Array("QBF_12000_2018", "Q29")
  varNameLinks_2018(2) = Array("QBF_2004-16_2018", "Q7")
  varNameLinks_2018(3) = Array("QBF_30713_2018", "Q63")
  varNameLinks_2018(4) = Array("QBF_30714_2018", "Q64")
  varNameLinks_2018(5) = Array("QBF_30715_2018", "Q65")
  varNameLinks_2018(6) = Array("QBF_30717_2018", "Q66")
  varNameLinks_2018(7) = Array("QBF_30719_2018", "Q67")
  varNameLinks_2018(8) = Array("QBF_30720_2018", "Q68")
  varNameLinks_2018(9) = Array("QFS_494_2018", "Q494")
  varNameLinks_2018(10) = Array("QFS_498_2018", "Q498")
  varNameLinks_2018(11) = Array("QFV_21114_2018", "Q30")
  varNameLinks_2018(12) = Array("QFV_21174_2018", "Q31")
  varNameLinks_2018(13) = Array("QFV_21262_2018", "Q32")
  varNameLinks_2018(14) = Array("QFV_21269_2018", "Q33")
  varNameLinks_2018(15) = Array("QFV_22126_2018", "Q34")
  varNameLinks_2018(16) = Array("QFV_22156_2018", "Q35")
  varNameLinks_2018(17) = Array("QFV_22241_2018_NO_SP", "Q36")
  varNameLinks_2018(18) = Array("QFV_22244_2018", "Q37")
  varNameLinks_2018(19) = Array("QFV_23155_2018", "Q38")
  varNameLinks_2018(20) = Array("QFV_23159_2018", "Q39")
  varNameLinks_2018(21) = Array("QRT_2004-10_2018", "Q6")
  varNameLinks_2018(22) = Array("QRT_30701_2018", "Q53")
  varNameLinks_2018(23) = Array("QRT_30702_2018", "Q54")
  varNameLinks_2018(24) = Array("QRT_30703_2018", "Q55")
  varNameLinks_2018(25) = Array("QRT_30705_2018", "Q57")
  varNameLinks_2018(26) = Array("QRT_30706_2018", "Q58")
  varNameLinks_2018(27) = Array("QRT_30707_2018", "Q59")
  varNameLinks_2018(28) = Array("QRT_30709_2018", "Q60")
  varNameLinks_2018(29) = Array("QWB_101_2018", "Q10")
  varNameLinks_2018(30) = Array("QWB_102_2018", "Q11")
  varNameLinks_2018(31) = Array("QWB_103_2018", "Q12")
  varNameLinks_2018(32) = Array("QWB_104_2018", "Q13")
  varNameLinks_2018(33) = Array("QWB_105_2018", "Q14")
  varNameLinks_2018(34) = Array("QWB_106_2018", "Q106")
  varNameLinks_2018(35) = Array("QWB_107_2018", "Q15")
  varNameLinks_2018(36) = Array("QWB_108_2018", "Q16")
  varNameLinks_2018(37) = Array("QWB_109_2018", "Q17")
  varNameLinks_2018(38) = Array("QWB_110_2018", "Q18")
  varNameLinks_2018(39) = Array("QWB_115_2018", "Q20")
  varNameLinks_2018(40) = Array("QWB_119_2018", "Q21")
  varNameLinks_2018(41) = Array("QWB_120_2018", "Q22")
  varNameLinks_2018(42) = Array("QWB_121_2018", "Q23")
  varNameLinks_2018(43) = Array("QWB_122_2018", "Q24")
  varNameLinks_2018(44) = Array("QWB_123_2018", "Q25")
  varNameLinks_2018(45) = Array("QWB_124_2018", "Q26")
  varNameLinks_2018(46) = Array("QWB_125_2018", "Q27")
  varNameLinks_2018(47) = Array("QWB_1_2018", "Q1")
  varNameLinks_2018(48) = Array("QWB_29003_2018", "Q48")
  varNameLinks_2018(49) = Array("QWB_29004_2018", "Q49")
  varNameLinks_2018(50) = Array("QWB_29016_2018", "Q50")
  varNameLinks_2018(51) = Array("QWB_29017_2018", "Q51")
  varNameLinks_2018(52) = Array("QWB_29025_2018", "Q52")
  varNameLinks_2018(53) = Array("QWB_2_2018", "Q2")
  varNameLinks_2018(54) = Array("QWB_3_2018", "Q3")
  varNameLinks_2018(55) = Array("Q_BS_2004_46_2018", "Q9")
  varNameLinks_2018(56) = Array("Q_BS_26345_2018", "Q42")
  varNameLinks_2018(57) = Array("Q_BS_26346_2018C", "Q43")
  varNameLinks_2018(58) = Array("Q_BS_26346_2018D", "Q43")
  varNameLinks_2018(87) = Array("Q_BS_26346_2018", "Q43")
  varNameLinks_2018(59) = Array("Q_BS_26347_2018", "Q44")
  varNameLinks_2018(60) = Array("Q_BS_26348_2018", "Q45")
  varNameLinks_2018(61) = Array("Q_BS_30741_2018", "Q89")
  varNameLinks_2018(62) = Array("Q_BS_30742_2018", "Q90")
  varNameLinks_2018(63) = Array("Q_BS_30743_2018", "Q91")
  varNameLinks_2018(64) = Array("Q_BS_30744_2018", "Q92")
  varNameLinks_2018(65) = Array("Q_BS_30747_2018", "Q95")
  varNameLinks_2018(66) = Array("Q_BS_30748_2018", "Q96")
  varNameLinks_2018(67) = Array("Q_BS_30749_2018", "Q97")
  varNameLinks_2018(68) = Array("Q_FP_30731_2018", "Q79")
  varNameLinks_2018(69) = Array("Q_FP_30732_2018", "Q80")
  varNameLinks_2018(70) = Array("Q_FP_30733_2018", "Q81")
  varNameLinks_2018(71) = Array("Q_FP_30734_2018", "Q82")
  varNameLinks_2018(72) = Array("Q_FP_30735_2018", "Q83")
  varNameLinks_2018(73) = Array("Q_FP_30739_2018", "Q87")
  varNameLinks_2018(74) = Array("Q_FP_30740_2018", "Q88")
  varNameLinks_2018(75) = Array("Q_RL_26339_2018", "Q40")
  varNameLinks_2018(76) = Array("Q_RL_26369_2018", "Q46")
  varNameLinks_2018(77) = Array("Q_RL_26370_2018", "Q47")
  varNameLinks_2018(78) = Array("Q_RL_30721_2018", "Q69")
  varNameLinks_2018(79) = Array("Q_RL_30722_2018", "Q70")
  varNameLinks_2018(80) = Array("Q_RL_30723_2018", "Q71")
  varNameLinks_2018(81) = Array("Q_RL_30724_2018", "Q72")
  varNameLinks_2018(82) = Array("Q_RL_30725_2018", "Q73")
  varNameLinks_2018(83) = Array("Q_RL_30726_2018", "Q74")
  varNameLinks_2018(84) = Array("Q_RL_30728_2018", "Q76")
  varNameLinks_2018(85) = Array("Q_RL_30729_2018", "Q77")
  varNameLinks_2018(86) = Array("Q_RL_30730_2018", "Q78")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2018 = New Collection
  Set p2018toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2018)
    varSubArray = varNameLinks_2018(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2018, CStr(varSubArray(1))) Then
      pOldTo2018.Add varSubArray(0), varSubArray(1)
    End If
    p2018toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub ConvertPointShapefiles()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pCoverCollection As New Collection
  Dim pDensityCollection As New Collection

  Dim pCoverToDensity As Collection
  Dim pDensityToCover As Collection
  Dim strCoverToDensityQuery As String
  Dim strDensityToCoverQuery As String
  Dim pCoverShouldChangeColl As Collection
  Dim pDensityShouldChangeColl As Collection
  Dim pRotateColl As Collection

  Debug.Print "---------------------"
  Call FillCollections(pCoverCollection, pDensityCollection, pCoverToDensity, pDensityToCover, _
    strCoverToDensityQuery, strDensityToCoverQuery, pCoverShouldChangeColl, pDensityShouldChangeColl)

  Set pRotateColl = FillRotateColl
  Dim pRotator As ITransform2D
  Dim strRotateBy As String
  Dim dblRotateVal As Double
  Dim pCollByQuadrat As Collection
  Dim varRotateElements() As Variant
  Dim pMidPoint As IPoint
  Set pMidPoint = New Point
  pMidPoint.PutCoords 0.5, 0.5

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String

  Dim strNewRoot As String
  Dim strExportPath As String

  Call DeclareWorkspaces(strRoot, strNewRoot, , , , strContainingFolder)

  Dim strMissingSpeciesPath As String
  Dim lngFileNum As Long
  Dim strSpeciesArray() As String
  Dim strMissingSummaryPath As String
  strMissingSpeciesPath = strContainingFolder & "\Missing_Species.csv"
  strMissingSummaryPath = strContainingFolder & "\Missing_Species_Summary.csv"

  lngFileNum = FreeFile(0)
  Open strMissingSpeciesPath For Output As lngFileNum
  Close #lngFileNum

  lngFileNum = FreeFile(0)
  Open strMissingSummaryPath For Output As lngFileNum
  Print #lngFileNum, """Species"",""Quadrats"""
  Close #lngFileNum

  Dim pSpeciesSummaryColl As New Collection
  Dim pSubColl As Collection
  Dim strSubNames() As String
  Dim varSubArray() As Variant
  Dim strSpeciesLine As String

  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long
  Dim lngIndex2 As Long

  Dim pDataset As IDataset
  Dim booFoundShapefiles As Boolean
  Dim varDatasets() As Variant

  Dim strNames() As String
  Dim strName As String
  Dim lngDatasetIndex As Long
  Dim lngNameIndex As Long
  Dim lngNameCount As Long
  Dim booFoundNames As Boolean
  Dim lngRecCount As Long

  Dim strFullNames() As String
  Dim lngFullNameCounter As Long

  Dim lngShapefileCount As Long
  Dim lngAcceptSFCount As Long
  lngShapefileCount = 0
  lngRecCount = 0
  lngAcceptSFCount = 0

  lngFullNameCounter = -1
  Dim pNameColl As New Collection
  Dim strHexify As String
  Dim strCorrect As String
  Dim pCheckCollection As Collection
  Dim strReport As String
  Dim booMadeChanges As Boolean
  Dim strEditReport As String
  Dim strExcelReport As String
  Dim strExcelFullReport As String
  Dim pFClass As IFeatureClass
  Dim strBase As String
  Dim strSplit() As String

  Dim strFolderName As String
  Dim booFoundPolys As Boolean
  Dim booFoundPoints As Boolean
  Dim pRepPointFClass As IFeatureClass
  Dim pRepPolyFClass As IFeatureClass
  Dim strNewFolder As String
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewWSFact As IWorkspaceFactory
  Dim pField As iField
  Dim pNewFields As esriSystem.IVariantArray

  Dim pNewDensityFClass As IFeatureClass
  Dim varDensityFieldIndexArray() As Variant
  Dim strNewDensityFClassName As String
  Dim booDensityHasFields As Boolean
  Dim lngDensityFClassIndex As Long
  Dim lngDensityQuadratIndex As Long
  Dim lngDensityYearIndex As Long
  Dim lngDensityTypeIndex As Long
  Dim lngDensityOrigFIDIndex As Long
  Dim lngDensityRotationIndex As Long

  Dim pNewGDBDensityFClass As IFeatureClass
  Dim varGDBDensityFieldIndexArray() As Variant
  Dim strGDBNewDensityFClassName As String
  Dim booGDBDensityHasFields As Boolean
  Dim lngGDBDensityFClassIndex As Long
  Dim lngGDBDensityQuadratIndex As Long
  Dim lngGDBDensityYearIndex As Long
  Dim lngGDBDensityTypeIndex As Long
  Dim lngGDBDensityOrigFIDIndex As Long
  Dim lngGDBDensityRotationIndex As Long

  Dim pNewCoverFClass As IFeatureClass
  Dim varCoverFieldIndexArray() As Variant
  Dim strNewCoverFClassName As String
  Dim booCoverHasFields As Boolean
  Dim lngCoverFClassIndex As Long
  Dim lngCoverQuadratIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverTypeIndex As Long
  Dim lngCoverOrigFIDIndex As Long
  Dim lngCoverRotationIndex As Long

  Dim pNewGDBCoverFClass As IFeatureClass
  Dim varGDBCoverFieldIndexArray() As Variant
  Dim strGDBNewCoverFClassName As String
  Dim booGDBCoverHasFields As Boolean
  Dim lngGDBCoverFClassIndex As Long
  Dim lngGDBCoverQuadratIndex As Long
  Dim lngGDBCoverYearIndex As Long
  Dim lngGDBCoverTypeIndex As Long
  Dim lngGDBCoverOrigFIDIndex As Long
  Dim lngGDBCoverRotationIndex As Long

  Dim strYear As String
  Dim strQuadrat As String
  Dim strFClassName As String
  Dim strType As String

  Dim pSrcFCursor As IFeatureCursor
  Dim pSrcFeature As IFeature
  Dim pDestFCursor As IFeatureCursor
  Dim pDestFBuffer As IFeatureBuffer
  Dim pDestGDBFCursor As IFeatureCursor
  Dim pDestGDBFBuffer As IFeatureBuffer

  Dim pDestFClass As IFeatureClass
  Dim pDestGDBFClass As IFeatureClass
  Dim pPoint As IPoint
  Dim pPolygon As IPolygon
  Dim pClone As IClone
  Dim varIndexArray() As Variant
  Dim varGDBIndexArray() As Variant
  Dim lngFClassIndex As Long
  Dim lngQuadratIndex As Long
  Dim lngYearIndex As Long
  Dim lngTypeIndex As Long
  Dim lngOrigFIDIndex As Long
  Dim lngIsEmptyIndex As Long
  Dim lngRotationIndex As Long
  Dim lngGDBFClassIndex As Long
  Dim lngGDBQuadratIndex As Long
  Dim lngGDBYearIndex As Long
  Dim lngGDBTypeIndex As Long
  Dim lngGDBOrigFIDIndex As Long
  Dim lngGDBIsEmptyIndex As Long
  Dim lngGDBRotationIndex As Long

  Dim varCoverIndexArray() As Variant
  Dim varCoverGDBIndexArray() As Variant
  Dim varDensityIndexArray() As Variant
  Dim varDensityGDBIndexArray() As Variant

  Dim pNewCombinedDensityFClass As IFeatureClass
  Dim varCombinedDensityFieldIndexArray() As Variant
  Dim strNewCombinedDensityFClassName As String
  Dim booCombinedDensityHasFields As Boolean
  Dim lngCombinedDensityFClassIndex As Long
  Dim lngCombinedDensityQuadratIndex As Long
  Dim lngCombinedDensityYearIndex As Long
  Dim lngCombinedDensityTypeIndex As Long
  Dim lngCombinedDensityOrigFIDIndex As Long
  Dim lngCombinedDensityRotationIndex As Long

  Dim pNewCombinedCoverFClass As IFeatureClass
  Dim varCombinedCoverFieldIndexArray() As Variant
  Dim strNewCombinedCoverFClassName As String
  Dim booCombinedCoverHasFields As Boolean
  Dim lngCombinedCoverFClassIndex As Long
  Dim lngCombinedCoverQuadratIndex As Long
  Dim lngCombinedCoverYearIndex As Long
  Dim lngCombinedCoverTypeIndex As Long
  Dim lngCombinedCoverOrigFIDIndex As Long
  Dim lngCombinedCoverRotationIndex As Long

  Dim pCombinedDestFClass As IFeatureClass
  Dim varCombinedIndexArray() As Variant
  Dim lngCombinedFClassIndex As Long
  Dim lngCombinedQuadratIndex As Long
  Dim lngCombinedYearIndex As Long
  Dim lngCombinedTypeIndex As Long
  Dim lngCombinedOrigFIDIndex As Long
  Dim lngCombinedIsEmptyIndex As Long
  Dim lngCombinedRotationIndex As Long

  Dim pCombinedFCursor As IFeatureCursor
  Dim pCombinedFBuffer As IFeatureBuffer
  Dim pCombinedDensityFCursor As IFeatureCursor
  Dim pCombinedDensityFBuffer As IFeatureBuffer
  Dim pCombinedCoverFCursor As IFeatureCursor
  Dim pCombinedCoverFBuffer As IFeatureBuffer

  Dim strIsEmpty As String
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.000001

  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pTopoOp As ITopologicalOperator4

  Dim pTempDataset As IDataset
  Dim pTempFClass As IFeatureClass
  Dim strCoverType As String
  Dim strDensityType As String
  Dim strAltType As String
  Dim pAltDestFClass As IFeatureClass
  Dim varAltIndexArray() As Variant
  Dim lngAltFClassIndex As Long
  Dim lngAltQuadratIndex As Long
  Dim lngAltYearIndex As Long
  Dim lngAltIsEmptyIndex As Long
  Dim lngAltTypeIndex As Long
  Dim lngAltRotationIndex As Long

  Dim pAltDestGDBFClass As IFeatureClass
  Dim varAltGDBIndexArray() As Variant
  Dim lngAltGDBFClassIndex As Long
  Dim lngAltGDBQuadratIndex As Long
  Dim lngAltGDBYearIndex As Long
  Dim lngAltGDBTypeIndex As Long
  Dim lngAltGDBIsEmptyIndex As Long
  Dim lngAltGDBRotationIndex As Long

  Dim pAltCombinedDestFClass As IFeatureClass
  Dim varAltCombinedIndexArray() As Variant
  Dim lngAltCombinedFClassIndex As Long
  Dim lngAltCombinedQuadratIndex As Long
  Dim lngAltCombinedYearIndex As Long
  Dim lngAltCombinedTypeIndex As Long
  Dim lngAltCombinedIsEmptyIndex As Long
  Dim lngAltCombinedRotationIndex As Long

  Dim pAltCombinedFCursor As IFeatureCursor
  Dim pAltCombinedFBuffer As IFeatureBuffer

  Dim pAltDestFCursor As IFeatureCursor
  Dim pAltDestFBuffer As IFeatureBuffer
  Dim pAltDestGDBFCursor As IFeatureCursor
  Dim pAltDestGDBFBuffer As IFeatureBuffer

  Dim var_C_to_D_IndexArray() As Variant
  Dim var_D_to_C_IndexArray() As Variant

  Dim strSpecies As String
  Dim lngSpeciesIndex As Long
  Dim strHexSpecies As String
  Dim booShouldChange As Boolean
  Dim varPoints() As Variant
  Dim pTestPolygon As IPolygon
  Dim pTestPoint As IPoint
  Dim lngConvertIndex As Long

  Dim pQuadrat As IPolygon
  Set pQuadrat = ReturnQuadratPolygon(pSpRef)
  Dim pNewPoly As IPolygon

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Set pNewFGDBWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strNewRoot & "\Combined_by_Site")
  Set pNewFeatFGDBWS = pNewFGDBWS

  For lngIndex = 0 To pFolders.Count - 1
    DoEvents
    strFolder = pFolders.Element(lngIndex)

    varDatasets = ReturnFeatureClassesOrNothing(strFolder, booFoundShapefiles, booFoundPolys, booFoundPoints, _
        pRepPointFClass, pRepPolyFClass)

    booFoundPolys = True
    booFoundPoints = True

    Debug.Print CStr(lngIndex + 1) & " of " & CStr(pFolders.Count) & "] " & strFolder
    If booFoundShapefiles Then
      Debug.Print "--> Found Shapefiles = " & CStr(booFoundShapefiles) & " [n = " & CStr(UBound(varDatasets) + 1) & "]"

      strFolderName = aml_func_mod.ReturnFilename(strFolder)
      strNewFolder = strNewRoot & "\Shapefiles\" & strFolderName

      If Not aml_func_mod.ExistFileDir(strNewFolder) Then
        MyGeneralOperations.CreateNestedFoldersByPath strNewFolder
      End If
      Set pNewWS = pNewWSFact.OpenFromFile(strNewFolder, 0)
      Set pNewFeatWS = pNewWS

      If booFoundPoints Then
        Set pDataset = pRepPointFClass
        strSplit = Split(pDataset.BrowseName, "_")
        strNewDensityFClassName = strSplit(0) & "_Density"

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewDensityFClassName) Then
          Set pDataset = pNewFeatWS.OpenFeatureClass(strNewDensityFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varDensityFieldIndexArray
        Set pNewDensityFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPointFClass, pNewWS, _
            varDensityFieldIndexArray, strNewDensityFClassName, booDensityHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewDensityFClass, lngDensityFClassIndex, lngDensityQuadratIndex, _
            lngDensityYearIndex, lngDensityTypeIndex, lngDensityOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewDensityFClass, strAbstract, strPurpose)
        DoEvents

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, strNewDensityFClassName) Then
          Set pDataset = pNewFeatFGDBWS.OpenFeatureClass(strNewDensityFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varGDBDensityFieldIndexArray
        Set pNewGDBDensityFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPointFClass, pNewFGDBWS, _
            varGDBDensityFieldIndexArray, strNewDensityFClassName, booGDBDensityHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewGDBDensityFClass, lngGDBDensityFClassIndex, lngGDBDensityQuadratIndex, _
            lngGDBDensityYearIndex, lngGDBDensityTypeIndex, lngGDBDensityOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewGDBDensityFClass, strAbstract, strPurpose)
        DoEvents

        If pNewCombinedDensityFClass Is Nothing Then
          If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, "Density_All") Then
            Set pDataset = pNewFeatFGDBWS.OpenFeatureClass("Density_All")
            pDataset.DELETE
            Set pDataset = Nothing
          End If

          Erase varCombinedDensityFieldIndexArray
          Set pNewCombinedDensityFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPointFClass, pNewFGDBWS, _
              varCombinedDensityFieldIndexArray, "Density_All", booCombinedDensityHasFields, esriGeometryPolygon)

          Call CreateNewFields(pNewCombinedDensityFClass, lngCombinedDensityFClassIndex, lngCombinedDensityQuadratIndex, _
              lngCombinedDensityYearIndex, lngCombinedDensityTypeIndex, lngCombinedDensityOrigFIDIndex)
          Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewCombinedDensityFClass, strAbstract, strPurpose)
          DoEvents

          Set pCombinedDensityFCursor = pNewCombinedDensityFClass.Insert(True)
          Set pCombinedDensityFBuffer = pNewCombinedDensityFClass.CreateFeatureBuffer
        End If

      End If

      If booFoundPolys Then
        Set pDataset = pRepPolyFClass
        strSplit = Split(pDataset.BrowseName, "_")
        strNewCoverFClassName = strSplit(0) & "_Cover"

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewCoverFClassName) Then
          Set pDataset = pNewFeatWS.OpenFeatureClass(strNewCoverFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varCoverFieldIndexArray
        Set pNewCoverFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPolyFClass, pNewWS, _
            varCoverFieldIndexArray, strNewCoverFClassName, booCoverHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewCoverFClass, lngCoverFClassIndex, lngCoverQuadratIndex, _
            lngCoverYearIndex, lngCoverTypeIndex, lngCoverOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewCoverFClass, strAbstract, strPurpose)
        DoEvents

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, strNewCoverFClassName) Then
          Set pDataset = pNewFeatFGDBWS.OpenFeatureClass(strNewCoverFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varGDBCoverFieldIndexArray
        Set pNewGDBCoverFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPolyFClass, pNewFGDBWS, _
            varGDBCoverFieldIndexArray, strNewCoverFClassName, booGDBCoverHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewGDBCoverFClass, lngGDBCoverFClassIndex, lngGDBCoverQuadratIndex, _
            lngGDBCoverYearIndex, lngGDBCoverTypeIndex, lngGDBCoverOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewGDBCoverFClass, strAbstract, strPurpose)
        DoEvents

        If pNewCombinedCoverFClass Is Nothing Then
          If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, "Cover_All") Then
            Set pDataset = pNewFeatFGDBWS.OpenFeatureClass("Cover_All")
            pDataset.DELETE
            Set pDataset = Nothing
          End If

          Erase varCombinedCoverFieldIndexArray
          Set pNewCombinedCoverFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPolyFClass, pNewFGDBWS, _
              varCombinedCoverFieldIndexArray, "Cover_All", booCombinedCoverHasFields, esriGeometryPolygon)

          Call CreateNewFields(pNewCombinedCoverFClass, lngCombinedCoverFClassIndex, lngCombinedCoverQuadratIndex, _
              lngCombinedCoverYearIndex, lngCombinedCoverTypeIndex, lngCombinedCoverOrigFIDIndex)
          Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewCombinedCoverFClass, strAbstract, strPurpose)
          DoEvents

          Set pCombinedCoverFCursor = pNewCombinedCoverFClass.Insert(True)
          Set pCombinedCoverFBuffer = pNewCombinedCoverFClass.CreateFeatureBuffer

        End If
      End If

      lngShapefileCount = lngShapefileCount + UBound(varDatasets) + 1

      For lngDatasetIndex = 0 To UBound(varDatasets)
        DoEvents

        Set pDataset = varDatasets(lngDatasetIndex)
        strFClassName = pDataset.BrowseName

        Set pFClass = pDataset

        lngSpeciesIndex = pFClass.FindField("Species")
        If lngSpeciesIndex = -1 Then
          DoEvents
        End If
        strSplit = Split(pDataset.BrowseName, "_")

        Debug.Print "  --> Adding Dataset '" & pDataset.BrowseName & "'"

        strQuadrat = strSplit(0)
        strYear = strSplit(1)

        If strYear = 2023 And strQuadrat = "Q52" Then
          DoEvents
        End If

        If MyGeneralOperations.CheckCollectionForKey(pRotateColl, Replace(strQuadrat, "Q", "", , , vbTextCompare)) Then
          Set pCollByQuadrat = pRotateColl.Item(Replace(strQuadrat, "Q", "", , , vbTextCompare))
          varRotateElements = pCollByQuadrat.Item(strYear)
          strRotateBy = varRotateElements(3)
        Else
          strRotateBy = "0"
        End If

        If strFClassName = "Q67_2012_D" Then
          DoEvents
        End If

        If strSplit(2) = "C" Then
          strType = "Cover"
          Set pDestFClass = pNewCoverFClass
          varIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestFClass)
          lngFClassIndex = lngCoverFClassIndex
          lngQuadratIndex = lngCoverQuadratIndex
          lngYearIndex = lngCoverYearIndex
          lngTypeIndex = lngCoverTypeIndex
          lngIsEmptyIndex = pDestFClass.FindField("IsEmpty")
          lngRotationIndex = pDestFClass.FindField("Revise_Rtn")

          Set pDestGDBFClass = pNewGDBCoverFClass
          varGDBIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestGDBFClass)
          lngGDBFClassIndex = lngGDBCoverFClassIndex
          lngGDBQuadratIndex = lngGDBCoverQuadratIndex
          lngGDBYearIndex = lngGDBCoverYearIndex
          lngGDBTypeIndex = lngGDBCoverTypeIndex
          lngGDBIsEmptyIndex = pDestGDBFClass.FindField("IsEmpty")
          lngGDBRotationIndex = pDestGDBFClass.FindField("Revise_Rtn")

          Set pCombinedDestFClass = pNewCombinedCoverFClass
          varCombinedIndexArray = ReturnArrayOfFieldLinks(pFClass, pCombinedDestFClass)
          lngCombinedFClassIndex = lngCombinedCoverFClassIndex
          lngCombinedQuadratIndex = lngCombinedCoverQuadratIndex
          lngCombinedYearIndex = lngCombinedCoverYearIndex
          lngCombinedTypeIndex = lngCombinedCoverTypeIndex
          lngCombinedIsEmptyIndex = pCombinedDestFClass.FindField("IsEmpty")
          lngCombinedRotationIndex = pCombinedDestFClass.FindField("Revise_Rtn")
          Set pCombinedFCursor = pCombinedCoverFCursor
          Set pCombinedFBuffer = pCombinedCoverFBuffer

          strAltType = "Density"
          Set pAltDestFClass = pNewDensityFClass
          ReDim varAltIndexArray(3, 4)

          varAltIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestFClass)
          lngAltFClassIndex = lngDensityFClassIndex
          lngAltQuadratIndex = lngDensityQuadratIndex
          lngAltYearIndex = lngDensityYearIndex
          lngAltTypeIndex = lngDensityTypeIndex
          lngAltIsEmptyIndex = pAltDestFClass.FindField("IsEmpty")
          lngAltRotationIndex = pAltDestFClass.FindField("Revise_Rtn")

          Set pAltDestGDBFClass = pNewGDBDensityFClass
          varAltGDBIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestGDBFClass)
          lngAltGDBFClassIndex = lngGDBDensityFClassIndex
          lngAltGDBQuadratIndex = lngGDBDensityQuadratIndex
          lngAltGDBYearIndex = lngGDBDensityYearIndex
          lngAltGDBTypeIndex = lngGDBDensityTypeIndex
          lngAltGDBIsEmptyIndex = pAltDestGDBFClass.FindField("IsEmpty")
          lngAltGDBRotationIndex = pAltDestGDBFClass.FindField("Revise_Rtn")

          Set pAltCombinedDestFClass = pNewCombinedDensityFClass
          varAltCombinedIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltCombinedDestFClass)
          lngAltCombinedFClassIndex = lngCombinedDensityFClassIndex
          lngAltCombinedQuadratIndex = lngCombinedDensityQuadratIndex
          lngAltCombinedYearIndex = lngCombinedDensityYearIndex
          lngAltCombinedTypeIndex = lngCombinedDensityTypeIndex
          lngAltCombinedIsEmptyIndex = pAltCombinedDestFClass.FindField("IsEmpty")
          lngAltCombinedRotationIndex = pAltCombinedDestFClass.FindField("Revise_Rtn")
          Set pAltCombinedFCursor = pCombinedDensityFCursor
          Set pAltCombinedFBuffer = pCombinedDensityFBuffer

        Else
          strType = "Density"
          Set pDestFClass = pNewDensityFClass
          varIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestFClass)
          lngFClassIndex = lngDensityFClassIndex
          lngQuadratIndex = lngDensityQuadratIndex
          lngYearIndex = lngDensityYearIndex
          lngTypeIndex = lngDensityTypeIndex
          lngIsEmptyIndex = pDestFClass.FindField("IsEmpty")
          lngRotationIndex = pDestFClass.FindField("Revise_Rtn")

          Set pDestGDBFClass = pNewGDBDensityFClass
          varGDBIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestGDBFClass)
          lngGDBFClassIndex = lngGDBDensityFClassIndex
          lngGDBQuadratIndex = lngGDBDensityQuadratIndex
          lngGDBYearIndex = lngGDBDensityYearIndex
          lngGDBTypeIndex = lngGDBDensityTypeIndex
          lngGDBIsEmptyIndex = pDestGDBFClass.FindField("IsEmpty")
          lngGDBRotationIndex = pDestGDBFClass.FindField("Revise_Rtn")

          Set pCombinedDestFClass = pNewCombinedDensityFClass
          varCombinedIndexArray = ReturnArrayOfFieldLinks(pFClass, pCombinedDestFClass)
          lngCombinedFClassIndex = lngCombinedDensityFClassIndex
          lngCombinedQuadratIndex = lngCombinedDensityQuadratIndex
          lngCombinedYearIndex = lngCombinedDensityYearIndex
          lngCombinedTypeIndex = lngCombinedDensityTypeIndex
          lngCombinedIsEmptyIndex = pCombinedDestFClass.FindField("IsEmpty")
          lngCombinedRotationIndex = pCombinedDestFClass.FindField("Revise_Rtn")
          Set pCombinedFCursor = pCombinedDensityFCursor
          Set pCombinedFBuffer = pCombinedDensityFBuffer

          strAltType = "Cover"
          Set pAltDestFClass = pNewCoverFClass
          varAltIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestFClass)
          lngAltFClassIndex = lngCoverFClassIndex
          lngAltQuadratIndex = lngCoverQuadratIndex
          lngAltYearIndex = lngCoverYearIndex
          lngAltTypeIndex = lngCoverTypeIndex
          lngAltIsEmptyIndex = pAltDestFClass.FindField("IsEmpty")
          lngAltRotationIndex = pAltDestFClass.FindField("Revise_Rtn")

          Set pAltDestGDBFClass = pNewGDBCoverFClass
          varAltGDBIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestGDBFClass)
          lngAltGDBFClassIndex = lngGDBCoverFClassIndex
          lngAltGDBQuadratIndex = lngGDBCoverQuadratIndex
          lngAltGDBYearIndex = lngGDBCoverYearIndex
          lngAltGDBTypeIndex = lngGDBCoverTypeIndex
          lngAltGDBIsEmptyIndex = pAltDestGDBFClass.FindField("IsEmpty")
          lngAltGDBRotationIndex = pAltDestGDBFClass.FindField("Revise_Rtn")

          Set pAltCombinedDestFClass = pNewCombinedCoverFClass
          varAltCombinedIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltCombinedDestFClass)
          lngAltCombinedFClassIndex = lngCombinedCoverFClassIndex
          lngAltCombinedQuadratIndex = lngCombinedCoverQuadratIndex
          lngAltCombinedYearIndex = lngCombinedCoverYearIndex
          lngAltCombinedTypeIndex = lngCombinedCoverTypeIndex
          lngAltCombinedIsEmptyIndex = pAltCombinedDestFClass.FindField("IsEmpty")
          lngAltCombinedRotationIndex = pAltCombinedDestFClass.FindField("Revise_Rtn")
          Set pAltCombinedFCursor = pCombinedCoverFCursor
          Set pAltCombinedFBuffer = pCombinedCoverFBuffer
        End If

        Set pDestFCursor = pDestFClass.Insert(True)
        Set pDestFBuffer = pDestFClass.CreateFeatureBuffer
        Set pDestGDBFCursor = pDestGDBFClass.Insert(True)
        Set pDestGDBFBuffer = pDestGDBFClass.CreateFeatureBuffer

        Set pAltDestFCursor = pAltDestFClass.Insert(True)
        Set pAltDestFBuffer = pAltDestFClass.CreateFeatureBuffer
        Set pAltDestGDBFCursor = pAltDestGDBFClass.Insert(True)
        Set pAltDestGDBFBuffer = pAltDestGDBFClass.CreateFeatureBuffer

        If pFClass.FindField("Cover") > -1 Or pFClass.FindField("Species") > -1 Then

          Set pSrcFCursor = pFClass.Search(Nothing, False)
          Set pSrcFeature = pSrcFCursor.NextFeature
          Do Until pSrcFeature Is Nothing
            strSpecies = pSrcFeature.Value(lngSpeciesIndex)

            If strSpecies = "Muhlenbergia tricholepis" Then
              DoEvents
            End If

                strSpecies, strNoteOnChanges)

            If Trim(strSpecies) = "" And strType = "Density" Then strSpecies = "No Point Species"
            If Trim(strSpecies) = "" And strType = "Cover" Then strSpecies = "No Polygon Species"

            strHexSpecies = HexifyName(strSpecies)
            If strType = "Density" Then
              If Not MyGeneralOperations.CheckCollectionForKey(pDensityShouldChangeColl, strHexSpecies) Then
                Debug.Print "Failed to find '" & strSpecies & "'..."

                lngFileNum = FreeFile(0)
                Open strMissingSpeciesPath For Append As lngFileNum
                Print #lngFileNum, """" & pDataset.BrowseName & """,""" & strSpecies & """"
                Close #lngFileNum

                If Not MyGeneralOperations.CheckCollectionForKey(pSpeciesSummaryColl, strSpecies) Then
                  Set pSubColl = New Collection
                  pSubColl.Add True, pDataset.BrowseName
                  ReDim strSubNames(0)
                  strSubNames(0) = pDataset.BrowseName
                  varSubArray = Array(strSubNames, pSubColl)
                  pSpeciesSummaryColl.Add varSubArray, strSpecies

                  If Not IsDimmed(strSpeciesArray) Then
                    ReDim strSpeciesArray(0)
                  Else
                    ReDim Preserve strSpeciesArray(UBound(strSpeciesArray) + 1)
                  End If
                  strSpeciesArray(UBound(strSpeciesArray)) = strSpecies
                Else
                  varSubArray = pSpeciesSummaryColl.Item(strSpecies)
                  strSubNames = varSubArray(0)
                  Set pSubColl = varSubArray(1)
                  If Not MyGeneralOperations.CheckCollectionForKey(pSubColl, pDataset.BrowseName) Then
                    ReDim Preserve strSubNames(UBound(strSubNames) + 1)
                    strSubNames(UBound(strSubNames)) = pDataset.BrowseName
                    pSubColl.Add True, pDataset.BrowseName
                    varSubArray = Array(strSubNames, pSubColl)
                    pSpeciesSummaryColl.Remove strSpecies
                    pSpeciesSummaryColl.Add varSubArray, strSpecies
                  End If
                End If

                booShouldChange = False
              Else
                booShouldChange = pDensityShouldChangeColl.Item(strHexSpecies)
              End If
            Else
              If Not MyGeneralOperations.CheckCollectionForKey(pCoverShouldChangeColl, strHexSpecies) Then
                Debug.Print "Failed to find '" & strSpecies & "'..."

                lngFileNum = FreeFile(0)
                Open strMissingSpeciesPath For Append As lngFileNum
                Print #lngFileNum, """" & pDataset.BrowseName & """,""" & strSpecies & """"
                Close #lngFileNum

                If Not MyGeneralOperations.CheckCollectionForKey(pSpeciesSummaryColl, strSpecies) Then
                  Set pSubColl = New Collection
                  pSubColl.Add True, pDataset.BrowseName
                  ReDim strSubNames(0)
                  strSubNames(0) = pDataset.BrowseName
                  varSubArray = Array(strSubNames, pSubColl)
                  pSpeciesSummaryColl.Add varSubArray, strSpecies

                  If Not IsDimmed(strSpeciesArray) Then
                    ReDim strSpeciesArray(0)
                  Else
                    ReDim Preserve strSpeciesArray(UBound(strSpeciesArray) + 1)
                  End If
                  strSpeciesArray(UBound(strSpeciesArray)) = strSpecies
                Else
                  varSubArray = pSpeciesSummaryColl.Item(strSpecies)
                  strSubNames = varSubArray(0)
                  Set pSubColl = varSubArray(1)
                  If Not MyGeneralOperations.CheckCollectionForKey(pSubColl, pDataset.BrowseName) Then
                    ReDim Preserve strSubNames(UBound(strSubNames) + 1)
                    strSubNames(UBound(strSubNames)) = pDataset.BrowseName
                    pSubColl.Add True, pDataset.BrowseName
                    varSubArray = Array(strSubNames, pSubColl)
                    pSpeciesSummaryColl.Remove strSpecies
                    pSpeciesSummaryColl.Add varSubArray, strSpecies
                  End If
                End If

                booShouldChange = False
              Else
                booShouldChange = pCoverShouldChangeColl.Item(strHexSpecies)
              End If
            End If

            If strType = "Density" Then
              Set pPoint = pSrcFeature.ShapeCopy
              If pPoint.IsEmpty Then
                Set pPolygon = New Polygon
              Else
                Set pPolygon = ReturnCircleClippedToQuadrat(pPoint, 0.001, 30, pQuadrat)
              End If
            Else
              Set pPolygon = pSrcFeature.ShapeCopy
            End If

            strIsEmpty = CBool(pPolygon.IsEmpty)
            Set pPolygon.SpatialReference = pSpRef

            Select Case strRotateBy
              Case "", "0"
                dblRotateVal = 0
              Case "CW 90"
                Set pRotator = pPolygon
                pRotator.Rotate pMidPoint, MyGeometricOperations.DegToRad(-90)   ' ASSUMING MATHEMATICAL ANGLES
                dblRotateVal = -90
              Case "CCW 90"
                Set pRotator = pPolygon
                pRotator.Rotate pMidPoint, MyGeometricOperations.DegToRad(90)    ' ASSUMING MATHEMATICAL ANGLES
                dblRotateVal = 90
              Case "180"
                Set pRotator = pPolygon
                pRotator.Rotate pMidPoint, MyGeometricOperations.DegToRad(180)   ' ASSUMING MATHEMATICAL ANGLES
                dblRotateVal = 180
              Case Else
                MsgBox "Unexpected Rotation! [" & strRotateBy & "]"
            End Select

            Set pTopoOp = pPolygon
            pTopoOp.IsKnownSimple = False
            pTopoOp.Simplify
            Set pClone = pPolygon

            Erase varPoints
            If booShouldChange Then
              If strType = "Cover" Then
                If pPolygon.IsEmpty Then
                  ReDim varPoints(0)
                  Set pTestPolygon = New Polygon
                  Set pTestPolygon.SpatialReference = pSpRef
                  Set varPoints(0) = pTestPolygon
                Else
                  varPoints = Margaret_Functions.FillPolygonWithPointArray(pPolygon, 0.005)
                  For lngConvertIndex = 0 To UBound(varPoints)
                    Set pTestPoint = varPoints(lngConvertIndex)
                    Set pTestPoint.SpatialReference = pSpRef
                    Set pTestPolygon = ReturnCircleClippedToQuadrat(pTestPoint, 0.001, 30, pQuadrat, pPolygon)
                    Set pTestPolygon.SpatialReference = pSpRef
                    Set varPoints(lngConvertIndex) = pTestPolygon
                    If Not pTestPolygon.IsEmpty Then
                    End If
                  Next lngConvertIndex
                End If
              Else  ' IF STARTING AS DENSITY AND CONVERTING TO COVER; DECIDE IF WE WANT TO MAKE THIS A BIGGER POLYGON
                ReDim varPoints(0)
                Set varPoints(0) = pClone.Clone
              End If
            End If

            If booShouldChange Then
              For lngConvertIndex = 0 To UBound(varPoints)
                Set pClone = varPoints(lngConvertIndex)
                Set pNewPoly = pClone.Clone
                If Not pNewPoly.IsEmpty Then

                  Set pAltDestFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltIndexArray, 2)
                    pAltDestFBuffer.Value(varAltIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltDestFBuffer.Value(lngAltFClassIndex) = strFClassName
                  pAltDestFBuffer.Value(lngAltQuadratIndex) = strQuadrat
                  pAltDestFBuffer.Value(lngAltYearIndex) = strYear
                  pAltDestFBuffer.Value(lngAltTypeIndex) = strAltType
                  pAltDestFBuffer.Value(lngAltIsEmptyIndex) = strIsEmpty
                  pAltDestFBuffer.Value(lngAltRotationIndex) = dblRotateVal
                  pAltDestFCursor.InsertFeature pAltDestFBuffer

                  Set pAltDestGDBFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltGDBIndexArray, 2)
                    pAltDestGDBFBuffer.Value(varAltGDBIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltGDBIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltDestGDBFBuffer.Value(lngAltGDBFClassIndex) = strFClassName
                  pAltDestGDBFBuffer.Value(lngAltGDBQuadratIndex) = strQuadrat
                  pAltDestGDBFBuffer.Value(lngAltGDBYearIndex) = strYear
                  pAltDestGDBFBuffer.Value(lngAltGDBTypeIndex) = strAltType
                  pAltDestGDBFBuffer.Value(lngAltGDBIsEmptyIndex) = strIsEmpty
                  pAltDestGDBFBuffer.Value(lngAltGDBRotationIndex) = dblRotateVal
                  pAltDestGDBFCursor.InsertFeature pAltDestGDBFBuffer

                  Set pAltCombinedFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltCombinedIndexArray, 2)
                    pAltCombinedFBuffer.Value(varAltCombinedIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltCombinedIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltCombinedFBuffer.Value(lngAltCombinedFClassIndex) = strFClassName
                  pAltCombinedFBuffer.Value(lngAltCombinedQuadratIndex) = strQuadrat
                  pAltCombinedFBuffer.Value(lngAltCombinedYearIndex) = strYear
                  pAltCombinedFBuffer.Value(lngAltCombinedTypeIndex) = strAltType
                  pAltCombinedFBuffer.Value(lngAltCombinedIsEmptyIndex) = strIsEmpty
                  pAltCombinedFBuffer.Value(lngAltCombinedRotationIndex) = dblRotateVal
                  pAltCombinedFCursor.InsertFeature pAltCombinedFBuffer
                End If
              Next lngConvertIndex

            Else
              Set pDestFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varIndexArray, 2)
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                If IsNull(pSrcFeature.Value(varIndexArray(1, lngIndex2))) Then
                  If pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Type = esriFieldTypeString Then
                    pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = ""
                  Else
                    pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = 0
                  End If
                Else
                  If varIndexArray(3, lngIndex2) > -1 Then
                    pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = pSrcFeature.Value(varIndexArray(1, lngIndex2))
                  End If
                End If
              Next lngIndex2
              pDestFBuffer.Value(lngFClassIndex) = strFClassName
              pDestFBuffer.Value(lngQuadratIndex) = strQuadrat
              pDestFBuffer.Value(lngYearIndex) = strYear
              pDestFBuffer.Value(lngTypeIndex) = strType
              pDestFBuffer.Value(lngIsEmptyIndex) = strIsEmpty
              pDestFBuffer.Value(lngRotationIndex) = dblRotateVal
              pDestFCursor.InsertFeature pDestFBuffer

              Set pDestGDBFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varGDBIndexArray, 2)
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                If varGDBIndexArray(3, lngIndex2) <> -1 Then
                  If pDestGDBFBuffer.Fields.Field(varGDBIndexArray(3, lngIndex2)).Editable Then
                    pDestGDBFBuffer.Value(varGDBIndexArray(3, lngIndex2)) = pSrcFeature.Value(varGDBIndexArray(1, lngIndex2))
                  End If
                End If
              Next lngIndex2
              pDestGDBFBuffer.Value(lngGDBFClassIndex) = strFClassName
              pDestGDBFBuffer.Value(lngGDBQuadratIndex) = strQuadrat
              pDestGDBFBuffer.Value(lngGDBYearIndex) = strYear
              pDestGDBFBuffer.Value(lngGDBTypeIndex) = strType
              pDestGDBFBuffer.Value(lngGDBIsEmptyIndex) = strIsEmpty
              pDestGDBFBuffer.Value(lngGDBRotationIndex) = dblRotateVal
              pDestGDBFCursor.InsertFeature pDestGDBFBuffer

              Set pCombinedFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varCombinedIndexArray, 2)
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"

                If varCombinedIndexArray(3, lngIndex2) <> -1 Then
                  If pCombinedFBuffer.Fields.Field(varCombinedIndexArray(3, lngIndex2)).Editable Then
                    pCombinedFBuffer.Value(varCombinedIndexArray(3, lngIndex2)) = pSrcFeature.Value(varCombinedIndexArray(1, lngIndex2))
                  End If
                End If
              Next lngIndex2
              pCombinedFBuffer.Value(lngCombinedFClassIndex) = strFClassName
              pCombinedFBuffer.Value(lngCombinedQuadratIndex) = strQuadrat
              pCombinedFBuffer.Value(lngCombinedYearIndex) = strYear
              pCombinedFBuffer.Value(lngCombinedTypeIndex) = strType
              pCombinedFBuffer.Value(lngCombinedIsEmptyIndex) = strIsEmpty
              pCombinedFBuffer.Value(lngCombinedRotationIndex) = dblRotateVal
              pCombinedFCursor.InsertFeature pCombinedFBuffer
            End If

            Set pSrcFeature = pSrcFCursor.NextFeature
          Loop

          pDestFCursor.Flush
          pDestGDBFCursor.Flush
          pCombinedFCursor.Flush

        End If
      Next lngDatasetIndex

      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Plot")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Plot")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Plot")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Plot")

    End If

  Next lngIndex

  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Plot")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Plot")

  If IsDimmed(strSpeciesArray) Then
    QuickSort.StringsAscending strSpeciesArray, 0, UBound(strSpeciesArray)
    lngFileNum = FreeFile(0)
    Open strMissingSummaryPath For Append As lngFileNum
    For lngIndex = 0 To UBound(strSpeciesArray)
      strSpeciesLine = ""
      strSpecies = strSpeciesArray(lngIndex)
      varSubArray = pSpeciesSummaryColl.Item(strSpecies)
      strSubNames = varSubArray(0)
      If IsDimmed(strSubNames) Then
        For lngIndex2 = 0 To UBound(strSubNames)
          strSpeciesLine = strSpeciesLine & strSubNames(lngIndex2) & IIf(lngIndex2 = UBound(strSubNames), "", ", ")
        Next lngIndex2
        Print #lngFileNum, """" & strSpecies & """,""" & strSpeciesLine & """"
      End If
    Next lngIndex
    Close lngFileNum
  End If

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pCoverCollection = Nothing
  Set pDensityCollection = Nothing
  Set pCoverToDensity = Nothing
  Set pDensityToCover = Nothing
  Set pCoverShouldChangeColl = Nothing
  Set pDensityShouldChangeColl = Nothing
  Set pRotateColl = Nothing
  Set pRotator = Nothing
  Set pCollByQuadrat = Nothing
  Erase varRotateElements
  Set pMidPoint = Nothing
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Erase strSpeciesArray
  Set pSpeciesSummaryColl = Nothing
  Set pSubColl = Nothing
  Erase strSubNames
  Erase varSubArray
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pCheckCollection = Nothing
  Set pFClass = Nothing
  Erase strSplit
  Set pRepPointFClass = Nothing
  Set pRepPolyFClass = Nothing
  Set pNewWS = Nothing
  Set pNewFeatWS = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewWSFact = Nothing
  Set pField = Nothing
  Set pNewFields = Nothing
  Set pNewDensityFClass = Nothing
  Erase varDensityFieldIndexArray
  Set pNewGDBDensityFClass = Nothing
  Erase varGDBDensityFieldIndexArray
  Set pNewCoverFClass = Nothing
  Erase varCoverFieldIndexArray
  Set pNewGDBCoverFClass = Nothing
  Erase varGDBCoverFieldIndexArray
  Set pSrcFCursor = Nothing
  Set pSrcFeature = Nothing
  Set pDestFCursor = Nothing
  Set pDestFBuffer = Nothing
  Set pDestGDBFCursor = Nothing
  Set pDestGDBFBuffer = Nothing
  Set pDestFClass = Nothing
  Set pDestGDBFClass = Nothing
  Set pPoint = Nothing
  Set pPolygon = Nothing
  Set pClone = Nothing
  Erase varIndexArray
  Erase varGDBIndexArray
  Erase varCoverIndexArray
  Erase varCoverGDBIndexArray
  Erase varDensityIndexArray
  Erase varDensityGDBIndexArray
  Set pNewCombinedDensityFClass = Nothing
  Erase varCombinedDensityFieldIndexArray
  Set pNewCombinedCoverFClass = Nothing
  Erase varCombinedCoverFieldIndexArray
  Set pCombinedDestFClass = Nothing
  Erase varCombinedIndexArray
  Set pCombinedFCursor = Nothing
  Set pCombinedFBuffer = Nothing
  Set pCombinedDensityFCursor = Nothing
  Set pCombinedDensityFBuffer = Nothing
  Set pCombinedCoverFCursor = Nothing
  Set pCombinedCoverFBuffer = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pTopoOp = Nothing
  Set pTempDataset = Nothing
  Set pTempFClass = Nothing
  Set pAltDestFClass = Nothing
  Erase varAltIndexArray
  Set pAltDestGDBFClass = Nothing
  Erase varAltGDBIndexArray
  Set pAltCombinedDestFClass = Nothing
  Erase varAltCombinedIndexArray
  Set pAltCombinedFCursor = Nothing
  Set pAltCombinedFBuffer = Nothing
  Set pAltDestFCursor = Nothing
  Set pAltDestFBuffer = Nothing
  Set pAltDestGDBFCursor = Nothing
  Set pAltDestGDBFBuffer = Nothing
  Erase var_C_to_D_IndexArray
  Erase var_D_to_C_IndexArray
  Erase varPoints
  Set pTestPolygon = Nothing
  Set pTestPoint = Nothing
  Set pQuadrat = Nothing
  Set pNewPoly = Nothing

End Sub

Public Sub AddEmptyFeaturesAndFeatureClassesToCleaned()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim strNewFolder As String
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewWSFact As IWorkspaceFactory

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Set pApp = Application
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String
  Dim strRecreatedFolder As String

  Dim strExportPath As String

  Call DeclareWorkspaces(strRoot, , , , strRecreatedFolder, strContainingFolder)

  Set pNewWSFact = New ShapefileWorkspaceFactory
  Set pNewWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Shapefiles", 0)
  Set pNewFeatWS = pNewWS
  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pNewFGDBWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Combined_by_Site.gdb", 0)
  Set pNewFeatFGDBWS = pNewFGDBWS

  Dim strEmptyFeatureReport As String
  Dim strEmptyKey As String
  Dim pDoneEmptyFeaturesColl As Collection
  Dim strEmptyYear As String
  Dim strQuad As String
  Dim pInsertCursor As IFeatureCursor
  Dim pInsertBuffer As IFeatureBuffer
  Dim pEmptyPolygon As IPolygon
  Dim strItems() As String
  Dim strSite As String
  Dim strPlot As String
  Dim lngStartYear As Long
  Dim lngEndYear As Long
  Dim booSurveyedThisYear As Boolean
  Dim lngEmptyYearIndex As Long
  Dim pQueryFilt As IQueryFilter
  Dim strShapefilePrefix As String
  Dim strShapefileSuffix As String
  Dim strGDBPrefix As String
  Dim strGDBSuffix As String
  Dim pYearsSiteSurveyed As Collection

  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewWS, strShapefilePrefix, strShapefileSuffix)
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewFGDBWS, strGDBPrefix, strGDBSuffix)

  lngStartYear = 2002
  lngEndYear = 2022
  Dim pSitesSurveyedByYearColl As Collection
  Set pSitesSurveyedByYearColl = More_Margaret_Functions.ReturnCollectionOfYearsSurveyedByQuadrat(lngStartYear, lngEndYear)
  Set pQueryFilt = New QueryFilter
  Set pDoneEmptyFeaturesColl = New Collection
  strEmptyFeatureReport = """Quadrat""" & vbTab & """Site""" & vbTab & """Plot""" & vbTab & _
      """Year""" & vbTab & """Type""" & vbCrLf

  Dim pQuadData As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim strQuadratNames() As String
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev(strQuadratNames, pPlotToQuadratConversion, _
      pQuadratToPlotConversion, varSites, varSitesSpecific)

  Dim lngIndex As Long
  Dim strQuadrat As String
  Dim pNewCoverFClass As IFeatureClass
  Dim pNewDensityFClass As IFeatureClass
  Dim pSpRef As ISpatialReference
  Dim strFClassName As String
  Dim pNewGDBCoverFClass As IFeatureClass
  Dim pNewGDBDensityFClass As IFeatureClass
  Dim pNewGDBCoverAllFClass As IFeatureClass
  Dim pNewGDBDensityAllFClass As IFeatureClass
  Dim pGeoDataset As IGeoDataset

  Set pNewGDBCoverAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Cover_All")
  Set pNewGDBDensityAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Density_All")
  Set pGeoDataset = pNewGDBCoverAllFClass
  Set pSpRef = pGeoDataset.SpatialReference

  pSBar.ShowProgressBar "Adding Empty Features...", 0, UBound(strQuadratNames) + 1, 1, True
  pProg.position = 0
  Dim lngIndex2 As Long
  Dim strNewFClassName As String

  For lngIndex = 0 To UBound(strQuadratNames)
    DoEvents
    pProg.Step
    strQuadrat = "Q" & strQuadratNames(lngIndex)
    strQuad = Replace(strQuadrat, "Q", "", , , vbTextCompare)

    If strQuad <> "496" Then
      strItems = pQuadData.Item(strQuad)
      strSite = Trim(strItems(1))
      If strSite = "" Then
        strSite = Trim(strItems(0))
      End If
      strPlot = Trim(strItems(2))

      strNewFClassName = ReplaceBadChars(strSite, True, True, True, True)
      Do Until InStr(1, strNewFClassName, "__", vbTextCompare) = 0
        strNewFClassName = Replace(strNewFClassName, "__", "_")
      Loop
      Debug.Print "Reviewing Quadrat " & strQuadrat & " [from " & strNewFClassName & "]"

      Set pYearsSiteSurveyed = pSitesSurveyedByYearColl.Item(strQuadrat)

      Set pNewWSFact = New ShapefileWorkspaceFactory
      Set pNewWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Shapefiles", 0)
      Set pNewFeatWS = pNewWS
      Set pNewCoverFClass = pNewFeatWS.OpenFeatureClass(strNewFClassName & "_Cover")
      Set pNewDensityFClass = pNewFeatWS.OpenFeatureClass(strNewFClassName & "_Density")

      Set pNewGDBCoverFClass = pNewFeatFGDBWS.OpenFeatureClass(strNewFClassName & "_Cover")
      Set pNewGDBDensityFClass = pNewFeatFGDBWS.OpenFeatureClass(strNewFClassName & "_Density")

      For lngEmptyYearIndex = lngStartYear To lngEndYear
        strEmptyYear = Format(lngEmptyYearIndex, "0")
        strFClassName = strQuadrat & "_" & strEmptyYear & "_"
        booSurveyedThisYear = pYearsSiteSurveyed.Item(strEmptyYear)

        If strEmptyYear = "2004" And strQuadrat = "Q84" Then
          DoEvents
        End If

        If booSurveyedThisYear Then
          If pNewCoverFClass.FindField("Year") = -1 And pNewCoverFClass.FindField("z_Year") > 0 Then
            pQueryFilt.WhereClause = strShapefilePrefix & "Quadrat" & strShapefileSuffix & " = '" & strQuadrat & "' AND " & _
                strShapefilePrefix & "z_Year" & strShapefileSuffix & " = '" & strEmptyYear & "'"
          Else
            pQueryFilt.WhereClause = strShapefilePrefix & "Quadrat" & strShapefileSuffix & " = '" & strQuadrat & "' AND " & _
                strShapefilePrefix & "Year" & strShapefileSuffix & " = '" & strEmptyYear & "'"
          End If

          If pNewCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewCoverFClass.Insert(True)
            Set pInsertBuffer = pNewCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("z_Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewDensityFClass.Insert(True)
            Set pInsertBuffer = pNewDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("z_Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Quadrat" & strGDBSuffix & " = '" & strQuadrat & "' AND " & _
               strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Quadrat" & strGDBSuffix & " = '" & strQuadrat & "' AND " & _
              strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

        End If
      Next lngEmptyYearIndex
    End If
  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  strEmptyFeatureReport = Replace(strEmptyFeatureReport, vbTab, ",", , , vbTextCompare)

  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainingFolder & "\Empty_Features_Added_to_Cleaned_dataset.csv")
  MyGeneralOperations.WriteTextFile strExportPath, strEmptyFeatureReport, False, False

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pNewWS = Nothing
  Set pNewFeatWS = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewWSFact = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pFolders = Nothing
  Set pDoneEmptyFeaturesColl = Nothing
  Set pInsertCursor = Nothing
  Set pInsertBuffer = Nothing
  Set pEmptyPolygon = Nothing
  Erase strItems
  Set pQueryFilt = Nothing
  Set pYearsSiteSurveyed = Nothing
  Set pSitesSurveyedByYearColl = Nothing
  Set pQuadData = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Erase strQuadratNames
  Set pNewCoverFClass = Nothing
  Set pNewDensityFClass = Nothing
  Set pSpRef = Nothing
  Set pNewGDBCoverFClass = Nothing
  Set pNewGDBDensityFClass = Nothing
  Set pNewGDBCoverAllFClass = Nothing
  Set pNewGDBDensityAllFClass = Nothing
  Set pGeoDataset = Nothing

End Sub

Public Sub AddEmptyFeaturesAndFeatureClasses(Optional booDoRecreated As Boolean = False)

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim strNewFolder As String
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewWSFact As IWorkspaceFactory

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Set pApp = Application
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String
  Dim strRecreatedFolder As String

  Dim strNewRoot As String
  Dim strExportPath As String

  Call DeclareWorkspaces(strRoot, strNewRoot, , , strRecreatedFolder, strContainingFolder)

  If booDoRecreated Then
    Set pNewWSFact = New ShapefileWorkspaceFactory
    Set pNewWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Shapefiles", 0)
    Set pNewFeatWS = pNewWS
    Set pNewWSFact = New FileGDBWorkspaceFactory
    Set pNewFGDBWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Combined_by_Site.gdb", 0)
    Set pNewFeatFGDBWS = pNewFGDBWS
  Else
    Set pNewWSFact = New ShapefileWorkspaceFactory
    Set pNewWS = pNewWSFact.OpenFromFile(strNewRoot & "\Shapefiles", 0)
    Set pNewFeatWS = pNewWS
    Set pNewWSFact = New FileGDBWorkspaceFactory
    Set pNewFGDBWS = pNewWSFact.OpenFromFile(strNewRoot & "\Combined_by_Site.gdb", 0)
    Set pNewFeatFGDBWS = pNewFGDBWS
  End If

  Dim strEmptyFeatureReport As String
  Dim strEmptyKey As String
  Dim pDoneEmptyFeaturesColl As Collection
  Dim strEmptyYear As String
  Dim strQuad As String
  Dim pInsertCursor As IFeatureCursor
  Dim pInsertBuffer As IFeatureBuffer
  Dim pEmptyPolygon As IPolygon
  Dim strItems() As String
  Dim strSite As String
  Dim strPlot As String
  Dim lngStartYear As Long
  Dim lngEndYear As Long
  Dim booSurveyedThisYear As Boolean
  Dim lngEmptyYearIndex As Long
  Dim pQueryFilt As IQueryFilter
  Dim strShapefilePrefix As String
  Dim strShapefileSuffix As String
  Dim strGDBPrefix As String
  Dim strGDBSuffix As String
  Dim pYearsSiteSurveyed As Collection

  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewWS, strShapefilePrefix, strShapefileSuffix)
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewFGDBWS, strGDBPrefix, strGDBSuffix)

  lngStartYear = 2002
  lngEndYear = 2022
  Dim pSitesSurveyedByYearColl As Collection
  Set pSitesSurveyedByYearColl = More_Margaret_Functions.ReturnCollectionOfYearsSurveyedByQuadrat(lngStartYear, lngEndYear)
  Set pQueryFilt = New QueryFilter
  Set pDoneEmptyFeaturesColl = New Collection
  strEmptyFeatureReport = """Quadrat""" & vbTab & """Site""" & vbTab & """Plot""" & vbTab & _
      """Year""" & vbTab & """Type""" & vbCrLf

  Dim pQuadData As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim strQuadratNames() As String
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev(strQuadratNames, pPlotToQuadratConversion, _
      pQuadratToPlotConversion, varSites, varSitesSpecific)

  Dim lngIndex As Long
  Dim strQuadrat As String
  Dim pNewCoverFClass As IFeatureClass
  Dim pNewDensityFClass As IFeatureClass
  Dim pSpRef As ISpatialReference
  Dim strFClassName As String
  Dim pNewGDBCoverFClass As IFeatureClass
  Dim pNewGDBDensityFClass As IFeatureClass
  Dim pNewGDBCoverAllFClass As IFeatureClass
  Dim pNewGDBDensityAllFClass As IFeatureClass
  Dim pGeoDataset As IGeoDataset

  Set pNewGDBCoverAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Cover_All")
  Set pNewGDBDensityAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Density_All")
  Set pGeoDataset = pNewGDBCoverAllFClass
  Set pSpRef = pGeoDataset.SpatialReference

  pSBar.ShowProgressBar "Adding Empty Features...", 0, UBound(strQuadratNames) + 1, 1, True
  pProg.position = 0

  For lngIndex = 0 To UBound(strQuadratNames)
    DoEvents
    pProg.Step
    strQuadrat = "Q" & strQuadratNames(lngIndex)
    strQuad = Replace(strQuadrat, "Q", "", , , vbTextCompare)

    If strQuad <> "496" Then
      strItems = pQuadData.Item(strQuad)
      strSite = Trim(strItems(1))
      If strSite = "" Then
        strSite = Trim(strItems(0))
      End If
      strPlot = Trim(strItems(2))
      Debug.Print "Reviewing Quadrat " & strQuadrat; ""

      Set pYearsSiteSurveyed = pSitesSurveyedByYearColl.Item(strQuadrat)

      Set pNewWSFact = New ShapefileWorkspaceFactory
      Set pNewWS = pNewWSFact.OpenFromFile(strNewRoot & "\Shapefiles\" & strQuadrat, 0)
      Set pNewFeatWS = pNewWS
      Set pNewCoverFClass = pNewFeatWS.OpenFeatureClass(strQuadrat & "_Cover")
      Set pNewDensityFClass = pNewFeatWS.OpenFeatureClass(strQuadrat & "_Density")

      Set pNewGDBCoverFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "_Cover")
      Set pNewGDBDensityFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "_Density")

      For lngEmptyYearIndex = lngStartYear To lngEndYear
        strEmptyYear = Format(lngEmptyYearIndex, "0")
        strFClassName = strQuadrat & "_" & strEmptyYear & "_"
        booSurveyedThisYear = pYearsSiteSurveyed.Item(strEmptyYear)
        If booSurveyedThisYear Then
          pQueryFilt.WhereClause = strShapefilePrefix & "Year" & strShapefileSuffix & " = '" & strEmptyYear & "'"

          If pNewCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewCoverFClass.Insert(True)
            Set pInsertBuffer = pNewCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewDensityFClass.Insert(True)
            Set pInsertBuffer = pNewDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Quadrat" & strGDBSuffix & " = '" & strQuadrat & "' AND " & _
              strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

        End If
      Next lngEmptyYearIndex
    End If
  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  strEmptyFeatureReport = Replace(strEmptyFeatureReport, vbTab, ",", , , vbTextCompare)

  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainingFolder & "\Empty_Features_Added.csv")
  MyGeneralOperations.WriteTextFile strExportPath, strEmptyFeatureReport, False, False

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pNewWS = Nothing
  Set pNewFeatWS = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewWSFact = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pFolders = Nothing
  Set pDoneEmptyFeaturesColl = Nothing
  Set pInsertCursor = Nothing
  Set pInsertBuffer = Nothing
  Set pEmptyPolygon = Nothing
  Erase strItems
  Set pQueryFilt = Nothing
  Set pYearsSiteSurveyed = Nothing
  Set pSitesSurveyedByYearColl = Nothing
  Set pQuadData = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Erase strQuadratNames
  Set pNewCoverFClass = Nothing
  Set pNewDensityFClass = Nothing
  Set pSpRef = Nothing
  Set pNewGDBCoverFClass = Nothing
  Set pNewGDBDensityFClass = Nothing
  Set pNewGDBCoverAllFClass = Nothing
  Set pNewGDBDensityAllFClass = Nothing
  Set pGeoDataset = Nothing

End Sub

Public Function FillRotateColl() As Collection

  Debug.Print "  --> Extracting Rotation Info..."
  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument

  Dim strNewSource As String
  strNewSource = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Source_Files_March_2018\HillPlotQC_Laughlin.xlsx"

  Dim strFolder As String
  Dim lngIndex As Long

  Dim strPlotLocNames() As String
  Dim pPlotLocColl As Collection

  Dim strPlotDataNames() As String
  Dim pPlotDataColl As Collection

  Call ReturnQuadratVegSoilData(pPlotDataColl, strPlotDataNames)
  Call ReturnQuadratCoordsAndNames(pPlotLocColl, strPlotLocNames)
  Set pQuadratColl = FillQuadratNameColl_Rev(strQuadratNames, pPlotToQuadratConversion, pQuadratToPlotConversion, _
       varSites, varSitesSpecific)

  Dim pWSFact As IWorkspaceFactory
  Dim pWS As IFeatureWorkspace
  Set pWSFact = New ExcelWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strNewSource, 0)

  Dim pTable As ITable
  Set pTable = pWS.OpenTable("For_ArcGIS$")

  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim pReturn As New Collection
  Dim lngSiteIndex As Long
  Dim lngPlotIndex As Long
  Dim lngYearIndex As Long
  Dim lngTurnIndex As Long
  Dim lngNotesIndex As Long
  Dim lngExtraNotesIndex As Long

  lngSiteIndex = pTable.FindField("Site")
  lngPlotIndex = pTable.FindField("Quadrat")
  lngYearIndex = pTable.FindField("Year")
  lngTurnIndex = pTable.FindField("Turn_quadrat")
  lngNotesIndex = pTable.FindField("Notes")
  lngExtraNotesIndex = pTable.FindField("Extra_Notes")

  Dim strSite As String
  Dim strPlot As String
  Dim strQuadrat As String
  Dim strYear As String
  Dim strTurn As String
  Dim strNotes As String
  Dim strExtra As String
  Dim varElement() As Variant
  Dim varVal As Variant
  Dim pQuadratByYearColl As Collection

  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    varVal = pRow.Value(lngSiteIndex)
    If IsNull(varVal) Then
      strSite = ""
    Else
      strSite = Trim(CStr(varVal))
    End If

    varVal = pRow.Value(lngPlotIndex)
    If IsNull(varVal) Then
      strPlot = ""
    Else
      strPlot = Trim(CStr(varVal))
    End If
    strQuadrat = pPlotToQuadratConversion.Item(strPlot)

    varVal = pRow.Value(lngYearIndex)
    If IsNull(varVal) Then
      strYear = ""
    Else
      strYear = Trim(CStr(varVal))
    End If
    varVal = pRow.Value(lngTurnIndex)
    If IsNull(varVal) Then
      strTurn = ""
    Else
      strTurn = Trim(CStr(varVal))
    End If
    varVal = pRow.Value(lngNotesIndex)
    If IsNull(varVal) Then
      strNotes = ""
    Else
      strNotes = Trim(CStr(varVal))
    End If
    varVal = pRow.Value(lngExtraNotesIndex)
    If IsNull(varVal) Then
      strExtra = ""
    Else
      strExtra = Trim(CStr(varVal))
    End If

    If Not MyGeneralOperations.CheckCollectionForKey(pReturn, strQuadrat) Then
      Set pQuadratByYearColl = ReturnEmptyYearColl
    Else
      Set pQuadratByYearColl = pReturn.Item(strQuadrat)
      pReturn.Remove strQuadrat
    End If

    If strYear <> "" And strTurn <> "" And strTurn <> "0" Then ' ONLY WORRY ABOUT CASES WHERE ROTATION IS DESIGNATED...
      varElement = Array(strSite, strPlot, strYear, strTurn, strNotes, strExtra)
      pQuadratByYearColl.Remove strYear
      pQuadratByYearColl.Add varElement, strYear
      Debug.Print "Plot '" & strPlot & "' [Quadrat = '" & strQuadrat & "'], " & strYear & ": Rotate " & strTurn
    End If

    pReturn.Add pQuadratByYearColl, strQuadrat
    Set pRow = pCursor.NextRow
  Loop

  Set FillRotateColl = pReturn

ClearMemory:
  Erase strQuadratNames
  Set pQuadratColl = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Set pMxDoc = Nothing
  Erase strPlotLocNames
  Set pPlotLocColl = Nothing
  Erase strPlotDataNames
  Set pPlotDataColl = Nothing
  Set pWSFact = Nothing
  Set pWS = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pReturn = Nothing
  Erase varElement
  varVal = Null
  Set pQuadratByYearColl = Nothing

End Function

Public Function ReturnEmptyYearColl() As Collection
  Dim lngIndex As Long
  Dim pReturn As New Collection
  Dim varElement() As Variant

  For lngIndex = 1990 To 2032
    varElement = Array("", "", Format(lngIndex, "0"), "0", "", "")  ' PRESET ROTATION TO ZERO
    pReturn.Add varElement, Format(lngIndex, "0")
  Next lngIndex
  Set ReturnEmptyYearColl = pReturn

  Set pReturn = Nothing
  Erase varElement
End Function

Public Sub ShiftFinishedShapefilesToCoordinateSystem()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray

  Dim strNewSource As String
  strNewSource = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Source_Files_March_2018\HillPlotQC_Laughlin.xlsx"

  Dim strOrigRoot As String
  Dim strModRoot As String
  Dim strShiftRoot As String
  Call DeclareWorkspaces(strOrigRoot, , strShiftRoot, , strModRoot)

  Dim strFolder As String
  Dim lngIndex As Long

  Dim strPlotLocNames() As String
  Dim pPlotLocColl As Collection

  Dim strPlotDataNames() As String
  Dim pPlotDataColl As Collection

  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim varSites() As Variant
  Dim varSiteSpecifics() As Variant

  Call ReturnQuadratVegSoilData(pPlotDataColl, strPlotDataNames)
  Call ReturnQuadratCoordsAndNames(pPlotLocColl, strPlotLocNames)
  Set pQuadratColl = FillQuadratNameColl_Rev(strQuadratNames, , , varSites, varSiteSpecifics)

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Dim pNewWSFact As IWorkspaceFactory
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pSrcWS As IFeatureWorkspace
  Dim pNewWS As IFeatureWorkspace
  Dim pSrcCoverFClass As IFeatureClass
  Dim pSrcDensFClass As IFeatureClass
  Dim pTopoOp As ITopologicalOperator4
  Dim lngQuadIndex As Long

  Dim strQuadrat As String
  Dim strDestFolder As String
  Dim strItem() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strFileHeader As String
  Dim dblCentroidX As Double
  Dim dblCentroidY As Double

  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace
  Dim pCoverAll As IFeatureClass
  Dim pDensityAll As IFeatureClass
  Dim varCoverIndexes() As Variant
  Dim varDensityIndexes() As Variant

  Dim strFClassName As String
  Dim strNameSplit() As String

  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pSrcWS = pNewWSFact.OpenFromFile(strModRoot & "\Combined_by_Site.gdb", 0)
  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strShiftRoot & "\Combined_by_Site")

  Set pWS = pSrcWS
  Set pDatasetEnum = pWS.Datasets(esriDTFeatureClass)
  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    If strFClassName <> "Cover_All" And strFClassName <> "Density_All" Then
      If InStr(1, strFClassName, "Density", vbTextCompare) Then
        Debug.Print strFClassName
      End If

      ExportFGDBFClass_2 pNewWS, pDataset, pMxDoc, pPlotLocColl, pQuadratColl, pCoverAll, pDensityAll, _
          varCoverIndexes, varDensityIndexes, False
    End If
    Set pDataset = pDatasetEnum.Next
  Loop

  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, CStr(varCoverIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Plot")

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, CStr(varDensityIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Plot")

  If Not aml_func_mod.ExistFileDir(strShiftRoot & "\Shapefiles") Then
    MyGeneralOperations.CreateNestedFoldersByPath (strShiftRoot & "\Shapefiles")
  End If
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Set pNewWS = pNewWSFact.OpenFromFile(strShiftRoot & "\Shapefiles", 0)

  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    If strFClassName <> "Cover_All" And strFClassName <> "Density_All" Then
      Debug.Print strFClassName

      ExportFGDBFClass_2 pNewWS, pDataset, pMxDoc, pPlotLocColl, pQuadratColl, pCoverAll, pDensityAll, _
          varCoverIndexes, varDensityIndexes, True
    End If
    Set pDataset = pDatasetEnum.Next
  Loop

  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, CStr(varCoverIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Plot")

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, CStr(varDensityIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Plot")
  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Erase strPlotLocNames
  Set pPlotLocColl = Nothing
  Erase strPlotDataNames
  Set pPlotDataColl = Nothing
  Erase strQuadratNames
  Set pQuadratColl = Nothing
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pNewWSFact = Nothing
  Set pSrcWS = Nothing
  Set pNewWS = Nothing
  Set pSrcCoverFClass = Nothing
  Set pSrcDensFClass = Nothing
  Set pTopoOp = Nothing
  Erase strItem
  Set pDatasetEnum = Nothing
  Set pWS = Nothing
  Set pCoverAll = Nothing
  Set pDensityAll = Nothing
  Erase varCoverIndexes
  Erase varDensityIndexes
  Erase strNameSplit

End Sub

Public Sub ExportFGDBFClass_2(pDestWS As IFeatureWorkspace, pSrcFClass As IFeatureClass, _
    pMxDoc As IMxDocument, pPlotLocColl As Collection, pQuadratColl As Collection, pCoverAll As IFeatureClass, _
    pDensityAll As IFeatureClass, varCoverIndexes() As Variant, varDensityIndexes() As Variant, _
    booIsShapefile As Boolean)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngYearIndex As Long
  Dim pInsertFCursor As IFeatureCursor
  Dim pInsertFBuffer As IFeatureBuffer
  Dim pDestFClass As IFeatureClass
  Dim varIndexArray() As Variant
  Dim strNewName As String
  Dim lngIndex As Long
  Dim pDataset As IDataset
  Dim lngQuadratIndex As Long

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Dim pPolygon As IPolygon
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix As String
  Dim strSuffix As String

  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim lngCount As Long
  Dim lngCounter As Long
  lngCount = pSrcFClass.FeatureCount(Nothing)
  lngCounter = 0

  Set pDataset = pSrcFClass
  strNewName = pDataset.Name
  Debug.Print "  --> " & strNewName
  DoEvents
  If MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, strNewName) Then
    Set pDataset = pDestWS.OpenFeatureClass(strNewName)
    pDataset.DELETE
  End If

  Dim pDensityFCursor As IFeatureCursor
  Dim pDensityFBuffer As IFeatureBuffer
  Dim pCoverFCursor As IFeatureCursor
  Dim pCoverFBuffer As IFeatureBuffer
  Dim pClone As IClone
  Dim booDoCover As Boolean
  Dim booDoDensity As Boolean

  Set pDataset = pSrcFClass
  booDoCover = InStr(1, pDataset.BrowseName, "Cover", vbTextCompare)
  booDoDensity = InStr(1, pDataset.BrowseName, "Density", vbTextCompare)

  If booDoCover Then
    If Not MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, "Cover_All") Then
      Set pCoverAll = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pDestWS, varCoverIndexes, _
            "Cover_All", True)
      Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pCoverAll, strAbstract, strPurpose)
    Else
      For lngIndex = 0 To UBound(varCoverIndexes, 2)
        varCoverIndexes(1, lngIndex) = pSrcFClass.FindField(CStr(varCoverIndexes(0, lngIndex)))
      Next lngIndex
    End If
  End If

  If booDoDensity Then
    If Not MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, "Density_All") Then
      Set pDensityAll = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pDestWS, varDensityIndexes, _
            "Density_All", True)
      Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDensityAll, strAbstract, strPurpose)
    Else
      For lngIndex = 0 To UBound(varDensityIndexes, 2)
        varDensityIndexes(1, lngIndex) = pSrcFClass.FindField(CStr(varDensityIndexes(0, lngIndex)))
      Next lngIndex
    End If
  End If

  If booDoCover Then
    Set pCoverFCursor = pCoverAll.Insert(True)
    Set pCoverFBuffer = pCoverAll.CreateFeatureBuffer
  End If
  If booDoDensity Then
    Set pDensityFCursor = pDensityAll.Insert(True)
    Set pDensityFBuffer = pDensityAll.CreateFeatureBuffer
  End If

  Set pDestFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pDestWS, varIndexArray, strNewName, True)
  Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDestFClass, strAbstract, strPurpose)
  Set pInsertFCursor = pDestFClass.Insert(True)
  Set pInsertFBuffer = pDestFClass.CreateFeatureBuffer

  pSBar.ShowProgressBar "Exporting '" & pDataset.BrowseName & "'...", 0, lngCount, 1, True
  pProg.position = 0

  Dim dblCentroidX As Double
  Dim dblCentroidY As Double
  Dim strQuadrat As String
  Dim strItem() As String
  Dim strPlot As String

  Dim varSrcVal As Variant
  Dim lngVarType As Long
  Dim lngDestIndex As Long

  Dim lngSPCodeFieldIndex As Long
  lngSPCodeFieldIndex = pSrcFClass.FindField("SPCODE")

  lngQuadratIndex = pSrcFClass.FindField("Quadrat")

  Set pFCursor = pSrcFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
      pInsertFCursor.Flush
      If booDoCover Then pCoverFCursor.Flush
      If booDoDensity Then pDensityFCursor.Flush
    End If

    strQuadrat = pFeature.Value(lngQuadratIndex)
    strItem = pQuadratColl.Item(Replace(strQuadrat, "Q", ""))
    strPlot = strItem(2)
    FillQuadratCenter strPlot, pPlotLocColl, dblCentroidX, dblCentroidY

    Set pPolygon = pFeature.ShapeCopy
    Call Margaret_Functions.ShiftPolygon(pPolygon, dblCentroidX, dblCentroidY)
    Set pClone = pPolygon

    Set pInsertFBuffer.Shape = pClone.Clone
    For lngIndex = 0 To UBound(varIndexArray, 2)

      varSrcVal = pFeature.Value(varIndexArray(1, lngIndex))
      lngDestIndex = varIndexArray(3, lngIndex)
      If booIsShapefile Then
        If IsNull(varSrcVal) Then
          If pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
            varSrcVal = ""
          End If
        End If
      End If
      pInsertFBuffer.Value(lngDestIndex) = varSrcVal
    Next lngIndex
    pInsertFCursor.InsertFeature pInsertFBuffer

    If booDoDensity Then
      Set pDensityFBuffer.Shape = pClone.Clone
      For lngIndex = 0 To UBound(varDensityIndexes, 2)
        varSrcVal = pFeature.Value(varDensityIndexes(1, lngIndex))
        lngDestIndex = varDensityIndexes(3, lngIndex)

        If varDensityIndexes(1, lngIndex) = lngSPCodeFieldIndex Then ' if SPCODE field, which should be integer
          If IsNull(varSrcVal) Then
            If booIsShapefile Then
              pDensityFBuffer.Value(lngDestIndex) = -999
            Else
              pDensityFBuffer.Value(lngDestIndex) = Null
            End If
          Else
            If Trim(CStr(pFeature.Value(varDensityIndexes(1, lngIndex)))) = "" Then
              pDensityFBuffer.Value(lngDestIndex) = Null
            Else
              pDensityFBuffer.Value(lngDestIndex) = pFeature.Value(varDensityIndexes(1, lngIndex))
            End If
          End If
        Else

          If booIsShapefile Then
            If IsNull(varSrcVal) Then
              If pDensityFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
                varSrcVal = -999
              ElseIf pDensityFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
                varSrcVal = -999
              ElseIf pDensityFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
                varSrcVal = ""
              End If
            End If
            pDensityFBuffer.Value(lngDestIndex) = varSrcVal
          Else
            pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = pFeature.Value(varDensityIndexes(1, lngIndex))
          End If

        End If
      Next lngIndex
      pDensityFCursor.InsertFeature pDensityFBuffer
    End If

    If booDoCover Then
      Set pCoverFBuffer.Shape = pClone.Clone
      For lngIndex = 0 To UBound(varCoverIndexes, 2)
        varSrcVal = pFeature.Value(varCoverIndexes(1, lngIndex))
        lngDestIndex = varCoverIndexes(3, lngIndex)

        If varCoverIndexes(1, lngIndex) = lngSPCodeFieldIndex Then ' if SPCODE field, which should be integer
          If IsNull(varSrcVal) Then
            If booIsShapefile Then
              pCoverFBuffer.Value(lngDestIndex) = -999
            Else
              pCoverFBuffer.Value(lngDestIndex) = Null
            End If
          Else
            If Trim(CStr(pFeature.Value(varCoverIndexes(1, lngIndex)))) = "" Then
              pCoverFBuffer.Value(lngDestIndex) = Null
            Else
              pCoverFBuffer.Value(lngDestIndex) = pFeature.Value(varCoverIndexes(1, lngIndex))
            End If
          End If

        Else
          If booIsShapefile Then
            lngDestIndex = varCoverIndexes(3, lngIndex)
            varSrcVal = pFeature.Value(varCoverIndexes(1, lngIndex))
            If IsNull(varSrcVal) Then
              If pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
                varSrcVal = -999
              ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
                varSrcVal = -999
              ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
                varSrcVal = ""
              End If
            End If
            pCoverFBuffer.Value(lngDestIndex) = varSrcVal
          Else
            pCoverFBuffer.Value(varCoverIndexes(3, lngIndex)) = pFeature.Value(varCoverIndexes(1, lngIndex))
          End If
        End If
      Next lngIndex
      pCoverFCursor.InsertFeature pCoverFBuffer
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  pInsertFCursor.Flush
  If booDoCover Then pCoverFCursor.Flush
  If booDoDensity Then pDensityFCursor.Flush

  pSBar.ShowProgressBar "Building Indexes for '" & pDataset.BrowseName & "'...", 0, 8, 1, True
  pProg.position = 0

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "SPCODE")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "FClassName")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Seedling")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Species")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Quadrat")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, CStr(varIndexArray(2, 9))) ' Year
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Orig_FID")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Site")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Verb_Spcs")
  pProg.Step
  DoEvents

  pSBar.HideProgressBar
  pProg.position = 0

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pInsertFCursor = Nothing
  Set pInsertFBuffer = Nothing
  Set pDestFClass = Nothing
  Erase varIndexArray
  Set pDataset = Nothing
  Set pPolygon = Nothing
  Set pQueryFilt = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pDensityFCursor = Nothing
  Set pDensityFBuffer = Nothing
  Set pCoverFCursor = Nothing
  Set pCoverFBuffer = Nothing
  Set pClone = Nothing

End Sub

Public Sub ExportFinalFClass(pDestWS As IFeatureWorkspace, pSrcFClass As IFeatureClass, _
    pMxDoc As IMxDocument, booIsShapefile As Boolean)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngYearIndex As Long
  Dim pInsertFCursor As IFeatureCursor
  Dim pInsertFBuffer As IFeatureBuffer
  Dim pDestFClass As IFeatureClass
  Dim varIndexArray() As Variant
  Dim strNewName As String
  Dim lngIndex As Long
  Dim pDataset As IDataset
  Dim lngQuadratIndex As Long

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Dim pPolygon As IPolygon
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix As String
  Dim strSuffix As String

  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim lngCount As Long
  Dim lngCounter As Long
  lngCount = pSrcFClass.FeatureCount(Nothing)
  lngCounter = 0

  Set pDataset = pSrcFClass
  strNewName = pDataset.Name
  Debug.Print "  --> " & strNewName
  DoEvents
  If MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, strNewName) Then
    Set pDataset = pDestWS.OpenFeatureClass(strNewName)
    pDataset.DELETE
  End If

  Dim pClone As IClone

  Set pDataset = pSrcFClass

  Set pDestFClass = ReturnEmptyFClassWithSameSchema_SpecialCase(pSrcFClass, pDestWS, varIndexArray, strNewName, True)
  Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDestFClass, strAbstract, strPurpose)
  Set pInsertFCursor = pDestFClass.Insert(True)
  Set pInsertFBuffer = pDestFClass.CreateFeatureBuffer

  pSBar.ShowProgressBar "Exporting '" & pDataset.BrowseName & "'...", 0, lngCount, 1, True
  pProg.position = 0

  Dim dblCentroidX As Double
  Dim dblCentroidY As Double
  Dim strQuadrat As String
  Dim strItem() As String
  Dim strPlot As String

  Dim varSrcVal As Variant
  Dim lngVarType As Long
  Dim lngDestIndex As Long

  Dim lngSPCodeFieldIndex As Long
  lngSPCodeFieldIndex = pSrcFClass.FindField("SPCODE")

  lngQuadratIndex = pSrcFClass.FindField("Quadrat")

  Set pFCursor = pSrcFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
      pInsertFCursor.Flush
    End If

    Set pInsertFBuffer.Shape = pFeature.ShapeCopy ' pClone.Clone
    For lngIndex = 0 To UBound(varIndexArray, 2)

      varSrcVal = pFeature.Value(varIndexArray(1, lngIndex))
      lngDestIndex = varIndexArray(3, lngIndex)
      If booIsShapefile Then
        If IsNull(varSrcVal) Then
          If pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
            varSrcVal = ""
          End If
        End If
      End If
      pInsertFBuffer.Value(lngDestIndex) = varSrcVal
    Next lngIndex
    pInsertFCursor.InsertFeature pInsertFBuffer

    Set pFeature = pFCursor.NextFeature
  Loop

  pInsertFCursor.Flush

  pSBar.ShowProgressBar "Building Indexes for '" & pDataset.BrowseName & "'...", 0, 8, 1, True
  pProg.position = 0

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "SPCODE")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "FClassName")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Seedling")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Species")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Quadrat")
  pProg.Step
  DoEvents
  If pDestFClass.FindField("Year") > -1 Then
    Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Year") ' Year
  Else
    Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "z_Year") ' Year
  End If
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Orig_FID")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Site")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Verb_Spcs")
  pProg.Step
  DoEvents

  pSBar.HideProgressBar
  pProg.position = 0

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pInsertFCursor = Nothing
  Set pInsertFBuffer = Nothing
  Set pDestFClass = Nothing
  Erase varIndexArray
  Set pDataset = Nothing
  Set pPolygon = Nothing
  Set pQueryFilt = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pClone = Nothing
  Erase strItem
  varSrcVal = Null

End Sub

Public Function ReturnEmptyFClassWithSameSchema_SpecialCase(pFClass As IFeatureClass, pWS_NothingForInMemory As IWorkspace, _
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
      If pSrcField.Name <> "Plot" And pSrcField.Name <> "Verb_Spcs" And pSrcField.Name <> "Verb_Type" And _
           pSrcField.Name <> "Revise_Rtn" And pSrcField.Name <> "FClassName" And pSrcField.Name <> "Orig_FID" Then
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
          If pSrcField.Name = "Quadrat" Then
            .length = 25
          End If
        End With
        pNewFieldArray.Add pNewField

        lngCounter = lngCounter + 1
        ReDim Preserve varReturnArray(3, lngCounter)
        If pSrcField.Name = "Quadrat" Then
          varReturnArray(0, lngCounter) = "Plot"
          varReturnArray(1, lngCounter) = pFields.FindField("Plot")
          varReturnArray(2, lngCounter) = pNewField.Name
        Else
          varReturnArray(0, lngCounter) = pSrcField.Name
          varReturnArray(1, lngCounter) = lngIndex
          varReturnArray(2, lngCounter) = pNewField.Name
        End If
      End If

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
  Set ReturnEmptyFClassWithSameSchema_SpecialCase = pNewFClass

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

Public Sub SplitShapefiles(strDestFolder As String, pSrcFClass As IFeatureClass, pShapefileWSFact As IWorkspaceFactory, _
    strTypeLetter As String, strQuadrat As String, pMxDoc As IMxDocument, dblCentroidX As Double, _
    dblCentroidY As Double)

  If Not aml_func_mod.ExistFileDir(strDestFolder) Then
    MyGeneralOperations.CreateNestedFoldersByPath strDestFolder
  End If

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngYearIndex As Long
  Dim pInsertFCursor As IFeatureCursor
  Dim pInsertFBuffer As IFeatureBuffer
  Dim pDestFClass As IFeatureClass
  Dim varIndexArray() As Variant
  Dim strNewName As String
  Dim lngIndex As Long
  Dim pDataset As IDataset

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Dim pWS As IFeatureWorkspace
  Set pWS = pShapefileWSFact.OpenFromFile(strDestFolder, 0)

  Dim pPolygon As IPolygon
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix As String
  Dim strSuffix As String
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pSrcFClass, strPrefix, strSuffix)
  Set pQueryFilt = New QueryFilter

  For lngYearIndex = 1995 To 2025
    pQueryFilt.WhereClause = strPrefix & "Year" & strSuffix & " = '" & Format(lngYearIndex, "0") & "'"
    If pSrcFClass.FeatureCount(pQueryFilt) > 0 Then
      strNewName = strQuadrat & "_" & Format(lngYearIndex, "0") & "_" & strTypeLetter
      Debug.Print "  --> " & strNewName
      DoEvents
      If MyGeneralOperations.CheckIfFeatureClassExists(pWS, strNewName) Then
        Set pDataset = pWS.OpenFeatureClass(strNewName)
        pDataset.DELETE
      End If

      Set pDestFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pWS, varIndexArray, strNewName, True)
      Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDestFClass, strAbstract, strPurpose)
      Set pInsertFCursor = pDestFClass.Insert(True)
      Set pInsertFBuffer = pDestFClass.CreateFeatureBuffer

      Set pFCursor = pSrcFClass.Search(pQueryFilt, False)
      Set pFeature = pFCursor.NextFeature
      Do Until pFeature Is Nothing
        Set pPolygon = pFeature.ShapeCopy
        Call Margaret_Functions.ShiftPolygon(pPolygon, dblCentroidX, dblCentroidY)
        Set pInsertFBuffer.Shape = pPolygon
        For lngIndex = 0 To UBound(varIndexArray, 2)
          pInsertFBuffer.Value(varIndexArray(3, lngIndex)) = pFeature.Value(varIndexArray(1, lngIndex))
        Next lngIndex
        pInsertFCursor.InsertFeature pInsertFBuffer
        Set pFeature = pFCursor.NextFeature
      Loop

      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, CStr(varIndexArray(2, 9))) ' Year
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Orig_FID")

      pInsertFCursor.Flush
    End If
  Next lngYearIndex

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pInsertFCursor = Nothing
  Set pInsertFBuffer = Nothing
  Set pDestFClass = Nothing
  Erase varIndexArray
  Set pDataset = Nothing
  Set pWS = Nothing
  Set pPolygon = Nothing
  Set pQueryFilt = Nothing

End Sub

Public Function MakeGridFLayer() As IFeatureLayer

  Dim pFClass As IFeatureClass
  Dim pFLayer As IFeatureLayer
  Dim pFields As esriSystem.IVariantArray
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit

  Set pFields = New esriSystem.varArray
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Class"
    .Type = esriFieldTypeInteger
  End With
  pFields.Add pField

  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)  ' UTM ZONE 12, NAD 83
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Set pFClass = MyGeneralOperations.CreateInMemoryFeatureClass_Empty(pFields, "Grid", pSpRef, esriGeometryPolyline, _
      False, False)

  Dim lngClassIndex As Long
  lngClassIndex = pFClass.FindField("Class")
  Dim pFCursor As IFeatureCursor
  Dim pFBuffer As IFeatureBuffer

  Dim lngIndex As Long
  Dim pPtColl As IPointCollection
  Dim pPolyline As IPolyline
  Dim pStartPoint As IPoint
  Dim pEndPoint As IPoint

  Set pFCursor = pFClass.Insert(True)
  Set pFBuffer = pFClass.CreateFeatureBuffer

  Dim pClass0Array As esriSystem.IArray
  Dim pCLass1Array As esriSystem.IArray
  Dim pClass2Array As esriSystem.IArray
  Dim pClass3Array As esriSystem.IArray

  Set pClass0Array = New esriSystem.Array
  Set pCLass1Array = New esriSystem.Array
  Set pClass2Array = New esriSystem.Array
  Set pClass3Array = New esriSystem.Array

  For lngIndex = 0 To 100 Step 5
    Set pPolyline = New Polyline
    Set pPolyline.SpatialReference = pSpRef
    Set pPtColl = pPolyline
    Set pStartPoint = New Point
    Set pStartPoint.SpatialReference = pSpRef
    pStartPoint.PutCoords CDbl(lngIndex) / 100, 0
    pPtColl.AddPoint pStartPoint
    Set pEndPoint = New Point
    Set pEndPoint.SpatialReference = pSpRef
    pEndPoint.PutCoords CDbl(lngIndex) / 100, 1
    pPtColl.AddPoint pEndPoint
    If lngIndex = 0 Or lngIndex = 100 Then
      pClass0Array.Add pPolyline
    ElseIf lngIndex Mod 20 = 0 Then
      pCLass1Array.Add pPolyline
    Else
      If lngIndex Mod 10 = 0 Then
        pClass2Array.Add pPolyline
      Else
        pClass3Array.Add pPolyline
      End If
    End If

    Set pPolyline = New Polyline
    Set pPolyline.SpatialReference = pSpRef
    Set pPtColl = pPolyline
    Set pStartPoint = New Point
    Set pStartPoint.SpatialReference = pSpRef
    pStartPoint.PutCoords 0, CDbl(lngIndex) / 100
    pPtColl.AddPoint pStartPoint
    Set pEndPoint = New Point
    Set pEndPoint.SpatialReference = pSpRef
    pEndPoint.PutCoords 1, CDbl(lngIndex) / 100
    pPtColl.AddPoint pEndPoint
    If lngIndex = 0 Or lngIndex = 100 Then
      pClass0Array.Add pPolyline
    ElseIf lngIndex Mod 20 = 0 Then
      pCLass1Array.Add pPolyline
    Else
      If lngIndex Mod 10 = 0 Then
        pClass2Array.Add pPolyline
      Else
        pClass3Array.Add pPolyline
      End If
    End If
  Next lngIndex

  Dim lngIndex2 As Long
  Dim varArrays() As Variant
  Dim pArray As esriSystem.IArray
  varArrays = Array(pClass3Array, pClass2Array, pCLass1Array, pClass0Array)

  For lngIndex = 0 To UBound(varArrays)
    Set pArray = varArrays(lngIndex)
    For lngIndex2 = 0 To pArray.Count - 1
      Set pPolyline = pArray.Element(lngIndex2)

      Set pFBuffer.Shape = pPolyline
      pFBuffer.Value(lngClassIndex) = 3 - lngIndex
      pFCursor.InsertFeature pFBuffer
    Next lngIndex2
  Next lngIndex

  pFCursor.Flush

  Set pFLayer = New FeatureLayer
  Set pFLayer.FeatureClass = pFClass
  pFLayer.Name = "Sample Grid"

  Set MakeGridFLayer = pFLayer
  CreateAndApplyGridRenderer pFLayer, "Class"

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  pMxDoc.AddLayer pFLayer

ClearMemory:
  Set pFClass = Nothing
  Set pFLayer = Nothing
  Set pFields = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pFCursor = Nothing
  Set pFBuffer = Nothing
  Set pPtColl = Nothing
  Set pPolyline = Nothing
  Set pStartPoint = Nothing
  Set pEndPoint = Nothing

End Function

Sub CreateAndApplyGridRenderer(pLayer As IFeatureLayer, strFieldName As String)

  Dim pFLayer As IFeatureLayer
  Set pFLayer = pLayer
  Dim pLyr As IGeoFeatureLayer
  Set pLyr = pFLayer

  Dim pFClass As IFeatureClass
  Set pFClass = pFLayer.FeatureClass

  Dim pClass0Symbol As ISimpleLineSymbol
  Dim pClass1Symbol As ISimpleLineSymbol
  Dim pClass2Symbol As ISimpleLineSymbol
  Dim pClass3Symbol As ICartographicLineSymbol

  Dim pClass0Color As IRgbColor
  Dim pClass1Color As IRgbColor
  Dim pClass2Color As IRgbColor
  Dim pClass3Color As IRgbColor

  Set pClass0Color = New RgbColor
  Set pClass1Color = New RgbColor
  Set pClass2Color = New RgbColor
  Set pClass3Color = New RgbColor

  pClass0Color.RGB = RGB(0, 0, 0)
  pClass1Color.RGB = RGB(100, 100, 100)
  pClass2Color.RGB = RGB(150, 150, 255)
  pClass3Color.RGB = RGB(175, 155, 255)

  Set pClass0Symbol = New SimpleLineSymbol
  pClass0Symbol.Width = 2
  pClass0Symbol.Color = pClass0Color

  Set pClass1Symbol = New SimpleLineSymbol
  pClass1Symbol.Width = 0.75
  pClass1Symbol.Color = pClass1Color

  Set pClass2Symbol = New SimpleLineSymbol
  pClass2Symbol.Width = 0.2
  pClass2Symbol.Color = pClass2Color

  Set pClass3Symbol = New CartographicLineSymbol
  With pClass3Symbol

    .Cap = esriLCSButt
    .Join = esriLJSBevel
    .Color = pClass3Color
    .Width = 0.1
  End With

  Dim pTemplate As ITemplate
  Set pTemplate = New Template
  pTemplate.AddPatternElement 1, 1

  Dim pLineProps As ILineProperties
  Set pLineProps = pClass3Symbol
  Set pLineProps.Template = pTemplate

  Dim pRender As IUniqueValueRenderer, n As Long
  Set pRender = New UniqueValueRenderer

  pRender.FieldCount = 1
  pRender.Field(0) = strFieldName
  pRender.DefaultSymbol = pClass2Symbol
  pRender.UseDefaultSymbol = False

  pRender.AddValue 0, strFieldName, pClass0Symbol
  pRender.Label(CStr(0)) = "Outline"
  pRender.Symbol(CStr(0)) = pClass0Symbol

  pRender.AddValue 1, strFieldName, pClass1Symbol
  pRender.Label(CStr(1)) = "20cm"
  pRender.Symbol(CStr(1)) = pClass1Symbol

  pRender.AddValue 2, strFieldName, pClass2Symbol
  pRender.Label(CStr(2)) = "10cm"
  pRender.Symbol(CStr(2)) = pClass2Symbol

  pRender.AddValue 3, strFieldName, pClass3Symbol
  pRender.Label(CStr(3)) = "5cm"
  pRender.Symbol(CStr(3)) = pClass3Symbol

  pRender.ColorScheme = "Custom"
  pRender.fieldType(0) = True
  Set pLyr.Renderer = pRender
  pLyr.DisplayField = strFieldName

  Dim hx As IRendererPropertyPage
  Set hx = New UniqueValuePropertyPage
  pLyr.RendererPropertyPageClassID = hx.ClassID

ClearMemory:
  Set pFLayer = Nothing
  Set pLyr = Nothing
  Set pFClass = Nothing
  Set pClass0Symbol = Nothing
  Set pClass1Symbol = Nothing
  Set pClass2Symbol = Nothing
  Set pClass3Symbol = Nothing
  Set pClass0Color = Nothing
  Set pClass1Color = Nothing
  Set pClass2Color = Nothing
  Set pClass3Color = Nothing
  Set pTemplate = Nothing
  Set pLineProps = Nothing
  Set hx = Nothing

End Sub

Public Function ReturnArrayOfFieldLinks(pSrcFClass As IFeatureClass, pDestFClass As IFeatureClass) As Variant()

  Dim pSrcFields As IFields
  Dim pDestFields As IFields
  Set pSrcFields = pSrcFClass.Fields
  Set pDestFields = pDestFClass.Fields

  Dim pField As iField
  Dim lngIndex As Long
  Dim varReturn() As Variant
  Dim lngCounter As Long
  Dim strNewName As String

  lngCounter = -1
  For lngIndex = 0 To pSrcFields.FieldCount - 1
    Set pField = pSrcFields.Field(lngIndex)
    If pField.Type <> esriFieldTypeGeometry Then
      lngCounter = lngCounter + 1
      ReDim Preserve varReturn(3, lngCounter)
      varReturn(0, lngCounter) = pField.Name
      varReturn(1, lngCounter) = lngIndex
      If pField.Type = esriFieldTypeOID Then
        strNewName = "Orig_FID"
      ElseIf pField.Name = "SP_CODE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP_CPDE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP_" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP" Then
        strNewName = "SPCODE"
      Else
        strNewName = pField.Name
      End If
      varReturn(2, lngCounter) = strNewName
      varReturn(3, lngCounter) = pDestFields.FindField(strNewName)

          " | " & varReturn(2, lngCounter) & " | " & varReturn(3, lngCounter)

    End If
  Next lngIndex

  ReturnArrayOfFieldLinks = varReturn

  Set pSrcFields = Nothing
  Set pDestFields = Nothing
  Erase varReturn

End Function

Public Function ReturnArrayOfFieldCrossLinks(pSrcFClass As IFeatureClass, pDestFClass As IFeatureClass) As Variant()

  Dim pSrcFields As IFields
  Dim pDestFields As IFields
  Set pSrcFields = pSrcFClass.Fields
  Set pDestFields = pDestFClass.Fields

  Dim pField As iField
  Dim lngIndex As Long
  Dim varReturn() As Variant
  Dim lngCounter As Long
  Dim strNewName As String

  lngCounter = -1
  For lngIndex = 0 To pSrcFields.FieldCount - 1
    Set pField = pSrcFields.Field(lngIndex)
    If pField.Type <> esriFieldTypeGeometry Then
      If pField.Type = esriFieldTypeOID Then
        strNewName = "Orig_FID"
      ElseIf pField.Name = "SP_CODE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP_CPDE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP_" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "x" Then
        strNewName = "coords_x1"
      ElseIf pField.Name = "y" Then
        strNewName = "coords_x2"
      ElseIf pField.Name = "coords_x1" Then
        strNewName = "x"
      ElseIf pField.Name = "coords_x2" Then
        strNewName = "y"
      Else
        strNewName = pField.Name
      End If

      If pDestFields.FindField(strNewName) > -1 Then
        lngCounter = lngCounter + 1
        ReDim Preserve varReturn(3, lngCounter)
        varReturn(0, lngCounter) = pField.Name
        varReturn(1, lngCounter) = lngIndex
        varReturn(2, lngCounter) = strNewName
        varReturn(3, lngCounter) = pDestFields.FindField(strNewName)
      End If

          " | " & varReturn(2, lngCounter) & " | " & varReturn(3, lngCounter)

    End If
  Next lngIndex

  ReturnArrayOfFieldCrossLinks = varReturn

  Set pSrcFields = Nothing
  Set pDestFields = Nothing
  Erase varReturn

End Function

Public Sub CreateNewFields(pNewFClass As IFeatureClass, lngFClassNameIndex As Long, _
    lngQuadratIndex As Long, lngYearIndex As Long, lngTypeIndex As Long, lngOrigFIDIndex As Long)

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim lngIDIndex As Long
  Dim lngIsEmptyIndex As Long

  Dim lngSPCodeIndex As Long
  lngSPCodeIndex = pNewFClass.FindField("SP_CODE")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SP_CPDE")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SPP_")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SPP")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SP")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SPCODE")
  If lngSPCodeIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "SPCODE"
      .Type = esriFieldTypeInteger
    End With
    pNewFClass.AddField pField
    lngSPCodeIndex = pNewFClass.FindField("SPCODE")
  End If

  lngIDIndex = pNewFClass.FindField("Id")
  If lngIDIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Id"
      .Type = esriFieldTypeInteger
    End With
    pNewFClass.AddField pField
    lngIDIndex = pNewFClass.FindField("Id")
  End If

  lngFClassNameIndex = pNewFClass.FindField("FClassName")
  If lngFClassNameIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "FClassName"
      .Type = esriFieldTypeString
      .length = 25
    End With
    pNewFClass.AddField pField
    lngFClassNameIndex = pNewFClass.FindField("FClassName")
  End If

  lngQuadratIndex = pNewFClass.FindField("Quadrat")
  If lngQuadratIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Quadrat"
      .Type = esriFieldTypeString
      .length = 10
    End With
    pNewFClass.AddField pField
    lngQuadratIndex = pNewFClass.FindField("Quadrat")
  End If

  lngYearIndex = pNewFClass.FindField("Year")
  If lngYearIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Year"
      .Type = esriFieldTypeString
      .length = 10
    End With
    pNewFClass.AddField pField
    lngYearIndex = pNewFClass.FindField("Year")
  End If

  lngTypeIndex = pNewFClass.FindField("Type")
  If lngTypeIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Type"
      .Type = esriFieldTypeString
      .length = 10
    End With
    pNewFClass.AddField pField
    lngTypeIndex = pNewFClass.FindField("Type")
  End If

  lngOrigFIDIndex = pNewFClass.FindField("Orig_FID")
  If lngOrigFIDIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Orig_FID"
      .Type = esriFieldTypeInteger
    End With
    pNewFClass.AddField pField
    lngOrigFIDIndex = pNewFClass.FindField("Orig_FID")
  End If

  lngIsEmptyIndex = pNewFClass.FindField("IsEmpty")
  If lngIsEmptyIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "IsEmpty"
      .Type = esriFieldTypeString
      .length = 5
    End With
    pNewFClass.AddField pField
    lngIsEmptyIndex = pNewFClass.FindField("IsEmpty")
  End If

  Set pField = Nothing
  Set pFieldEdit = Nothing

End Sub

Public Function CheckSpeciesAgainstSpecialConversions(varSpecialConversions() As Variant, strQuadrat As String, _
    lngYear As Long, strSpecies As String, strNoteOnChanges As String) As String

  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long
  Dim strTestSpecies As String

  If InStr(1, strSpecies, "Muhlenbergia tricholepis", vbTextCompare) > 0 Then
    DoEvents
  End If

  CheckSpeciesAgainstSpecialConversions = Trim(strSpecies)
  strNoteOnChanges = ""

  For lngIndex = 0 To UBound(varSpecialConversions, 2)
    strTestQuadrat = varSpecialConversions(0, lngIndex)
    lngTestYear = varSpecialConversions(1, lngIndex)
    strTestSpecies = varSpecialConversions(2, lngIndex)
    If StrComp(Trim(strQuadrat), Trim(strTestQuadrat), vbTextCompare) = 0 Then
      If lngTestYear = lngYear Or lngTestYear = -999 Then
        If StrComp(Trim(strSpecies), Trim(strTestSpecies), vbTextCompare) = 0 Then
          CheckSpeciesAgainstSpecialConversions = Trim(CStr(varSpecialConversions(3, lngIndex)))
          strNoteOnChanges = Trim(CStr(varSpecialConversions(5, lngIndex)))
          Exit Function
        End If
      End If
    End If
  Next lngIndex

End Function

Public Function SpecialConversionExistsForYearQuadrat(varSpecialConversions() As Variant, strQuadrat As String, _
    lngYear As Long) As Boolean

  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long

  SpecialConversionExistsForYearQuadrat = False

  For lngIndex = 0 To UBound(varSpecialConversions, 2)
    strTestQuadrat = varSpecialConversions(0, lngIndex)
    lngTestYear = varSpecialConversions(1, lngIndex)
    If StrComp(Trim(strQuadrat), Trim(strTestQuadrat), vbTextCompare) = 0 Then
      If lngTestYear = lngYear Or lngTestYear = -999 Then
        SpecialConversionExistsForYearQuadrat = True
        Exit Function
      End If
    End If
  Next lngIndex

End Function

Public Function ReturnQueryStringFromSpecialConversions(strQuadrat As String, lngYear As Long, _
    booIsCover As Boolean, varSpecialConversions() As Variant, strInstructions As String, _
    lngSpecialIndex As Long, booYearQuadrat As Boolean) As String

  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long
  Dim strTestSpecies As String

  Dim strQueryString As String
  Dim varStrings() As Variant
  Dim varInstructions() As Variant

  ReturnQueryStringFromSpecialConversions = ""
  strInstructions = ""

  booYearQuadrat = False

  lngIndex = lngSpecialIndex
  strTestQuadrat = varSpecialConversions(0, lngIndex)
  lngTestYear = varSpecialConversions(1, lngIndex)
  If StrComp(Trim(strQuadrat), Trim(strTestQuadrat), vbTextCompare) = 0 Then
    If lngTestYear = lngYear Then
      booYearQuadrat = True
      varStrings = varSpecialConversions(4, lngIndex)
      varInstructions = varSpecialConversions(6, lngIndex)
      If booIsCover Then
        ReturnQueryStringFromSpecialConversions = varStrings(0)
        strInstructions = varInstructions(0)
      Else
        ReturnQueryStringFromSpecialConversions = varStrings(1)
        strInstructions = varInstructions(1)
      End If
      Exit Function
    End If
  End If

ClearMemory:
  Erase varStrings
  Erase varInstructions

End Function

Public Sub CopyFeaturesInFClassBasedOnQueryFilter(pFClass As IFeatureClass, _
    strQueryPair As String, strEditReport As String, strExcelReport As String, _
    booMadeEdits As Boolean, lngNameIndex As Long, strBase As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strOID As String

  Dim strSource As String
  Dim strQueryString As String
  Dim strQuerySplit() As String
  Dim pQueryFilt As IQueryFilter

  strQuerySplit = Split(strQueryPair, "|")
  strSource = strQuerySplit(0)
  strQueryString = strQuerySplit(1)
  Set pQueryFilt = New QueryFilter
  pQueryFilt.WhereClause = strQueryString

  Dim pWS As IFeatureWorkspace
  Dim pDataset As IDataset
  Dim pDonorFClass As IFeatureClass

  Set pDataset = pFClass
  Set pWS = pDataset.Workspace
  Set pDonorFClass = pWS.OpenFeatureClass(strSource)

  Dim lngLinks() As Long
  Dim lngLinkIndex As Long
  Dim pField As iField
  Dim lngIndex As Long

  lngLinkIndex = -1
  For lngIndex = 0 To pDonorFClass.Fields.FieldCount - 1
    Set pField = pDonorFClass.Fields.Field(lngIndex)
    If pField.Type <> esriFieldTypeGeometry And pField.Type <> esriFieldTypeOID Then
      If pFClass.FindField(pField.Name) > -1 Then
        lngLinkIndex = lngLinkIndex + 1
        ReDim Preserve lngLinks(1, lngLinkIndex)
        lngLinks(0, lngLinkIndex) = lngIndex  ' DONOR INDEX
        lngLinks(1, lngLinkIndex) = pFClass.FindField(pField.Name) ' RECIPIENT INDEX
      End If
    End If
  Next lngIndex

  Dim varFeatures() As Variant
  Dim lngArrayIndex As Long

  lngArrayIndex = -1
  Set pFCursor = pDonorFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    lngArrayIndex = lngArrayIndex + 1
    ReDim Preserve varFeatures(lngArrayIndex)
    Set varFeatures(lngArrayIndex) = pFeature
    Set pFeature = pFCursor.NextFeature
  Loop

  Dim lngSrcSpeciesIndex As Long
  Dim lngDestSpeciesIndex As Long
  Dim strSrcSpecies As String
  Dim strDestSpecies As String
  Dim pSrcEnv As IEnvelope
  Dim pDestEnv As IEnvelope
  lngSrcSpeciesIndex = pDonorFClass.FindField("species")
  lngDestSpeciesIndex = pFClass.FindField("species")

  Dim varDestData() As Variant
  Dim lngDestIndex As Long
  lngDestIndex = -1
  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pDestEnv = pFeature.ShapeCopy.Envelope
    strDestSpecies = pFeature.Value(lngDestSpeciesIndex)

    lngDestIndex = lngDestIndex + 1
    ReDim Preserve varDestData(1, lngDestIndex)
    varDestData(0, lngDestIndex) = strDestSpecies
    Set varDestData(1, lngDestIndex) = pDestEnv

    Set pFeature = pFCursor.NextFeature
  Loop

  Dim pFBuffer As IFeatureBuffer
  Dim booFeatureExists As Boolean
  Dim lngIndex2 As Long

  If lngArrayIndex > -1 Then

    Set pFBuffer = pFClass.CreateFeatureBuffer
    Set pFCursor = pFClass.Insert(True)
    For lngIndex = 0 To lngArrayIndex
      Set pFeature = varFeatures(lngIndex)
      Set pSrcEnv = pFeature.ShapeCopy.Envelope
      strSrcSpecies = pFeature.Value(lngSrcSpeciesIndex)

      booFeatureExists = False
      If lngDestIndex > -1 Then
        For lngIndex2 = 0 To lngDestIndex
          strDestSpecies = varDestData(0, lngIndex2)
          If StrComp(Trim(strDestSpecies), Trim(strSrcSpecies), vbTextCompare) = 0 Then
            Set pDestEnv = varDestData(1, lngIndex2)
            If pDestEnv.XMin = pSrcEnv.XMin And pDestEnv.XMax = pSrcEnv.XMax And _
                    pDestEnv.YMin = pSrcEnv.YMin And pDestEnv.YMax = pSrcEnv.YMax Then
              booFeatureExists = True
              Exit For
            End If
          End If
        Next lngIndex2
      End If

      If Not booFeatureExists Then
        Set pFBuffer.Shape = pFeature.ShapeCopy
        For lngIndex2 = 0 To UBound(lngLinks, 2)
          pFBuffer.Value(lngLinks(1, lngIndex2)) = pFeature.Value(lngLinks(0, lngIndex2))
        Next lngIndex2
        pFCursor.InsertFeature pFBuffer
        pFCursor.Flush

        booMadeEdits = True
        strOID = CStr(pFBuffer.Value(pFBuffer.Fields.FindField(pFClass.OIDFieldName)))
        strOID = String(4 - Len(strOID), " ") & strOID

        strName = pFBuffer.Value(lngNameIndex)
        strOrigName = strName

        strName = Replace(strName, vbCrLf, "")
        strName = Replace(strName, vbNewLine, "")
        strName = Trim(strName)
        Debug.Print "  --> " & CStr(pFeature.OID) & "] Copying new " & strName & " from " & strSource & "..."

        strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Species '" & _
            strName & "': Feature copied from " & strSource & " on Where Clause = " & _
            strQueryString & vbCrLf
        strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(strOID) & """" & vbTab & _
              """" & strName & """" & vbTab & """Feature copied from " & strSource & """" & vbCrLf

      End If
    Next lngIndex
  End If

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strQuerySplit
  Set pQueryFilt = Nothing
  Set pWS = Nothing
  Set pDataset = Nothing
  Set pDonorFClass = Nothing
  Erase lngLinks
  Set pField = Nothing
  Erase varFeatures
  Set pSrcEnv = Nothing
  Set pDestEnv = Nothing
  Erase varDestData
  Set pFBuffer = Nothing

End Sub

Public Sub DeleteFeaturesInFClassBasedOnQueryFilter(pFClass As IFeatureClass, _
    pQueryFiltOrNothing As IQueryFilter, strEditReport As String, strExcelReport As String, _
    booMadeEdits As Boolean, lngNameIndex As Long, strBase As String)

  Dim pTable As ITable
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strOID As String

  Set pTable = pFClass
  If pTable.RowCount(pQueryFiltOrNothing) > 0 Then

    Set pFCursor = pFClass.Search(pQueryFiltOrNothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing

      booMadeEdits = True
      strOID = CStr(pFeature.OID)
      strOID = String(4 - Len(strOID), " ") & strOID

      strName = pFeature.Value(lngNameIndex)
      strOrigName = strName

      strName = Replace(strName, vbCrLf, "")
      strName = Replace(strName, vbNewLine, "")
      strName = Trim(strName)
      Debug.Print "  --> " & CStr(pFeature.OID) & "] Deleting " & strName & "..."

      strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Species '" & _
          strName & "': Feature Deleted using 'DeleteSearchedRows' method on Where Clause = " & _
          pQueryFiltOrNothing.WhereClause & vbCrLf
      strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
            """" & strName & """" & vbTab & """Feature Deleted""" & vbCrLf

      Set pFeature = pFCursor.NextFeature
    Loop

    pTable.DeleteSearchedRows pQueryFiltOrNothing
  End If

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pTable = Nothing

End Sub

Public Sub UpdateGeometryInFClassBasedOnQueryFilter(pFClass As IFeatureClass, pQueryFiltOrNothing As IQueryFilter, _
    varSpecialConversions() As Variant, strQuadrat As String, strYear As String, strEditReport As String, _
    strExcelReport As String, booMadeEdits As Boolean, lngNameIndex As Long, pCheckCollection As Collection, _
    strBase As String, strSourceSpecies As String, strDestSpecies As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strCorrect As String
  Dim strHexify As String
  Dim strTrimName As String
  Dim strOID As String
  Dim strNoteOnChanges As String

  Dim pOrigPolygon As IPolygon
  Dim pNewPolygon As IPolygon
  Dim pCentroid As IPoint
  Dim pArea As IArea

  Set pFCursor = pFClass.Update(pQueryFiltOrNothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    Set pOrigPolygon = pFeature.ShapeCopy
    Set pArea = pOrigPolygon
    Set pCentroid = pArea.Centroid
    Set pNewPolygon = MyGeometricOperations.CreateCircleAroundPoint(pCentroid, 0.001, 36)
    Set pFeature.Shape = pNewPolygon
    pFCursor.UpdateFeature pFeature

    booMadeEdits = True
    strOID = CStr(pFeature.OID)
    strOID = String(4 - Len(strOID), " ") & strOID
    strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Changed Polygon Geometry to Centroid'" & vbCrLf
    strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
          """" & strName & """" & vbTab & """Converted to Polygon Centroid""" & vbCrLf

    Set pFeature = pFCursor.NextFeature
  Loop

  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Sub

Public Sub UpdateSpeciesInFClassBasedOnQueryFilter(pFClass As IFeatureClass, pQueryFiltOrNothing As IQueryFilter, _
    varSpecialConversions() As Variant, strQuadrat As String, strYear As String, strEditReport As String, _
    strExcelReport As String, booMadeEdits As Boolean, lngNameIndex As Long, pCheckCollection As Collection, _
    strBase As String, strSourceSpecies As String, strDestSpecies As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strCorrect As String
  Dim strHexify As String
  Dim strTrimName As String
  Dim strOID As String
  Dim strNoteOnChanges As String

  Set pFCursor = pFClass.Update(pQueryFiltOrNothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    strOrigName = strName

    strName = Replace(strName, vbCrLf, "")
    strName = Replace(strName, vbNewLine, "")
    strName = Trim(strName)

    If InStr(1, strName, "tricholepis", vbTextCompare) > 0 Then
      DoEvents
    End If

    strCorrect = CheckSpeciesAgainstSpecialConversions(varSpecialConversions, strQuadrat, CLng(strYear), _
                strName, strNoteOnChanges)

    strHexify = HexifyName(strName)
    If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
      strCorrect = pCheckCollection.Item(strHexify)
    End If

    strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
    strCorrect = Replace(strCorrect, "Pachera ", "Packera ")
    If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
      MsgBox "Found carriage return!"
      DoEvents
    End If

    If InStr(1, strCorrect, " Asclepias sp.", vbTextCompare) > 0 Then
      DoEvents
    End If

    If InStr(1, strCorrect, "formossisimus", vbTextCompare) > 0 Then
      DoEvents
      strCorrect = Replace(strCorrect, "Erigeron formossisimus", "Erigeron formosissimus")
    End If

    If Left(strCorrect, 1) = " " Then
      strTrimName = Trim(strCorrect)
      If strCorrect <> " " Then
        DoEvents
      End If
      strCorrect = strTrimName
      strHexify = HexifyName(strTrimName)
      If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
        strCorrect = pCheckCollection.Item(strHexify)
      End If
      strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
      If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
        MsgBox "Found carriage return!"
        DoEvents
      End If
    End If

    If Left(strCorrect, 1) = " " Or Left(strName, 1) = " " Or InStr(1, strCorrect, vbTab) > 0 Or InStr(1, strName, vbTab) > 0 Then
      If strName <> " " Then
        DoEvents
      End If
    End If

    If StrComp(Trim(strOrigName), Trim(strSourceSpecies), vbTextCompare) = 0 Or _
       StrComp(Trim(strName), Trim(strSourceSpecies), vbTextCompare) = 0 Or _
       StrComp(Trim(strCorrect), Trim(strSourceSpecies), vbTextCompare) = 0 Then

      booMadeEdits = True
      strOID = CStr(pFeature.OID)
      strOID = String(4 - Len(strOID), " ") & strOID
      strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Changed '" & _
          strName & "' to '" & strDestSpecies & "'" & vbCrLf
      strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
            """" & strName & """" & vbTab & """" & strDestSpecies & """" & vbCrLf
      pFeature.Value(lngNameIndex) = strDestSpecies
      pFCursor.UpdateFeature pFeature
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Sub

Public Sub UpdateSpeciesInFClassBasedConversionArray(pFClass As IFeatureClass, _
    varSpecialConversions() As Variant, strQuadrat As String, strYear As String, strEditReport As String, _
    strExcelReport As String, booMadeEdits As Boolean, lngNameIndex As Long, pCheckCollection As Collection, _
    strBase As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strCorrect As String
  Dim strHexify As String
  Dim strTrimName As String
  Dim strOID As String
  Dim strNoteOnChanges As String

  Set pFCursor = pFClass.Update(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    strOrigName = strName

    strName = Replace(strName, vbCrLf, "")
    strName = Replace(strName, Chr(9), "")  ' PROBLEM INTRODUCED 2020 WITH Q30 COVER Potentilla crinita
    strName = Replace(strName, vbNewLine, "")
    strName = Trim(strName)

    If InStr(1, strName, "tricholepis", vbTextCompare) > 0 Then
      DoEvents
    End If

    strCorrect = CheckSpeciesAgainstSpecialConversions(varSpecialConversions, strQuadrat, CLng(strYear), _
                strName, strNoteOnChanges)

    strHexify = HexifyName(strName)
    If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
      strCorrect = pCheckCollection.Item(strHexify)
    End If

    strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
    strCorrect = Replace(strCorrect, "Pachera ", "Packera ")
    If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
      MsgBox "Found carriage return!"
      DoEvents
    End If

    If InStr(1, strCorrect, " Asclepias sp.", vbTextCompare) > 0 Then
      DoEvents
    End If

    If InStr(1, strCorrect, "formossisimus", vbTextCompare) > 0 Then
      DoEvents
      strCorrect = Replace(strCorrect, "Erigeron formossisimus", "Erigeron formosissimus")
    End If

    If Left(strCorrect, 1) = " " Then
      strTrimName = Trim(strCorrect)
      If strCorrect <> " " Then
        DoEvents
      End If
      strCorrect = strTrimName
      strHexify = HexifyName(strTrimName)
      If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
        strCorrect = pCheckCollection.Item(strHexify)
      End If
      strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
      If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
        MsgBox "Found carriage return!"
        DoEvents
      End If
    End If

    If Left(strCorrect, 1) = " " Or Left(strName, 1) = " " Or InStr(1, strCorrect, vbTab) > 0 Or InStr(1, strName, vbTab) > 0 Then
      If strName <> " " Then
        DoEvents
      End If
    End If

    If strOrigName <> strCorrect Then
      booMadeEdits = True
      strOID = CStr(pFeature.OID)
      strOID = String(4 - Len(strOID), " ") & strOID
      strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Changed '" & _
          strName & "' to '" & strCorrect & "'" & vbCrLf
      strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
            """" & strName & """" & vbTab & """" & strCorrect & """" & vbCrLf
      pFeature.Value(lngNameIndex) = strCorrect
      pFCursor.UpdateFeature pFeature
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Sub

Public Sub ReplaceNamesInShapefile(pDataset As IDataset, pCheckCollection As Collection, booMadeEdits As Boolean, _
    strEditReport As String, strBase As String, strExcelReport As String, varSpecialConversions() As Variant, _
    varQueryConversions() As Variant)

  On Error GoTo ErrHandler

  Dim pFClass As IFeatureClass
  Set pFClass = pDataset
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngNameIndex As Long
  Dim lngIndex As Long
  Dim strName As String
  Dim strReturn() As String
  Dim strHexify As String
  Dim strCorrect As String
  Dim strTrimName As String
  booMadeEdits = False
  Dim strOID As String
  Dim strOrigName As String
  Dim booIsCover As Boolean
  booIsCover = StrComp(Right(pDataset.BrowseName, 2), "_C", vbTextCompare) = 0

  lngIndex = -1

  Dim pDoneColl As New Collection

  strEditReport = "Edits to '" & pDataset.BrowseName & "':" & vbCrLf
  strExcelReport = ""

  lngNameIndex = pFClass.FindField("Species")
  If lngNameIndex = -1 Then lngNameIndex = pFClass.FindField("Cover")
  If lngNameIndex = -1 Then
    MsgBox "Unexpected Event!"
    DoEvents
    GoTo ClearMemory
  End If

  Dim strSplit() As String
  Dim strQuadrat As String
  Dim strYear As String
  Dim strNoteOnChanges As String
  strSplit = Split(pDataset.BrowseName, "_")
  strQuadrat = strSplit(0)
  strYear = strSplit(1)

  If strYear = "2014" And strQuadrat = "Q50" Then
    DoEvents
  End If

  Dim pQueryFilt As IQueryFilter
  Dim strQueryString As String
  Dim varStrings() As Variant
  Dim strInstructions As String
  Dim varFeaturesToDelete() As Variant
  Dim lngDeleteIndex As Long
  Dim booYearQuadrat As Boolean
  Dim pTable As ITable

  Dim lngSpecialIndex As Long

  Set pQueryFilt = New QueryFilter

  For lngSpecialIndex = 0 To UBound(varQueryConversions, 2)

    strQueryString = ReturnQueryStringFromSpecialConversions(strQuadrat, CLng(strYear), booIsCover, _
        varQueryConversions, strInstructions, lngSpecialIndex, booYearQuadrat)

    If booYearQuadrat Then

      pQueryFilt.WhereClause = strQueryString

      If strInstructions = "Delete" Then
        DeleteFeaturesInFClassBasedOnQueryFilter pFClass, pQueryFilt, strEditReport, strExcelReport, _
              booMadeEdits, lngNameIndex, strBase

      ElseIf strInstructions = "Change Species" Then   ' THEN CHANGE SPECIES NAMES FOR SELECTED ROWS

        UpdateSpeciesInFClassBasedOnQueryFilter pFClass, pQueryFilt, varSpecialConversions, strQuadrat, _
              strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
              strBase, CStr(varQueryConversions(2, lngSpecialIndex)), CStr(varQueryConversions(3, lngSpecialIndex))

      ElseIf strInstructions = "Copy Features" Then
        CopyFeaturesInFClassBasedOnQueryFilter pFClass, strQueryString, strEditReport, strExcelReport, _
              booMadeEdits, lngNameIndex, strBase

      ElseIf strInstructions = "Return Centroid" Then
        UpdateGeometryInFClassBasedOnQueryFilter pFClass, pQueryFilt, varSpecialConversions, strQuadrat, _
              strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
              strBase, CStr(varQueryConversions(2, lngSpecialIndex)), CStr(varQueryConversions(3, lngSpecialIndex))

      End If
    End If
  Next lngSpecialIndex

  UpdateSpeciesInFClassBasedConversionArray pFClass, varSpecialConversions, strQuadrat, _
        strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
        strBase

  If SpecialConversionExistsForYearQuadrat(varSpecialConversions, strQuadrat, CLng(strYear)) Then
    UpdateSpeciesInFClassBasedConversionArray pFClass, varSpecialConversions, strQuadrat, _
          strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
          strBase
  End If

  GoTo ClearMemory
  Exit Sub

ErrHandler:
  DoEvents

ClearMemory:
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strReturn
  Set pDoneColl = Nothing
  Erase strSplit
  Set pQueryFilt = Nothing
  Erase varStrings
  Erase varFeaturesToDelete
  Set pTable = Nothing

End Sub

Public Function CreateVarSpecialConversions(varQueryConversions() As Variant) As Variant()

  Dim varSpecialConversions() As Variant
  ReDim varSpecialConversions(6, 9)
  Dim lngMaxIndex As Long

  varSpecialConversions(0, 0) = "Q90"
  varSpecialConversions(1, 0) = -999
  varSpecialConversions(2, 0) = "Antennaria parvifolia"
  varSpecialConversions(3, 0) = "Antennaria rosulata"
  varSpecialConversions(4, 0) = Array("", "")
  varSpecialConversions(5, 0) = "Email Margaret Dec. 21, 2017"
  varSpecialConversions(6, 0) = Array("", "")

  varSpecialConversions(0, 3) = "Q80"
  varSpecialConversions(1, 3) = -999
  varSpecialConversions(2, 3) = "Muhlenbergia tricholepis"
  varSpecialConversions(3, 3) = "Bouteloua gracilis"
  varSpecialConversions(4, 3) = Array("", "")
  varSpecialConversions(5, 3) = "Email Margaret Dec. 21, 2017"
  varSpecialConversions(6, 3) = Array("", "")

  varSpecialConversions(0, 4) = "Q80"
  varSpecialConversions(1, 4) = -999
  varSpecialConversions(2, 4) = "Muhlenbergia rigens"
  varSpecialConversions(3, 4) = "Muhlenbergia wrightii"
  varSpecialConversions(4, 4) = Array("", "")
  varSpecialConversions(5, 4) = "Email Margaret Dec. 21, 2017"
  varSpecialConversions(6, 4) = Array("", "")

  varSpecialConversions(0, 5) = "Q88"
  varSpecialConversions(1, 5) = -999
  varSpecialConversions(2, 5) = "Unknown forb"
  varSpecialConversions(3, 5) = "Coreopsis tinctoria"
  varSpecialConversions(4, 5) = Array("", "")
  varSpecialConversions(5, 5) = "Email Margaret Dec. 21, 2017"
  varSpecialConversions(6, 5) = Array("", "")

  varSpecialConversions(0, 6) = "Q46"
  varSpecialConversions(1, 6) = -999
  varSpecialConversions(2, 6) = "Muhlenbergia wrightii"
  varSpecialConversions(3, 6) = "Muhlenbergia rigens"
  varSpecialConversions(4, 6) = Array("", "")
  varSpecialConversions(5, 6) = "Email Margaret March 6, 2018"
  varSpecialConversions(6, 6) = Array("", "")

  varSpecialConversions(0, 7) = "Q46"
  varSpecialConversions(1, 7) = -999
  varSpecialConversions(2, 7) = "Blepharoneuron tricholepis"
  varSpecialConversions(3, 7) = "Muhlenbergia rigens"
  varSpecialConversions(4, 7) = Array("", "")
  varSpecialConversions(5, 7) = "Email Margaret March 6, 2018"
  varSpecialConversions(6, 7) = Array("", "")

  varSpecialConversions(0, 8) = "Q46"
  varSpecialConversions(1, 8) = -999
  varSpecialConversions(2, 8) = "Muhlenbergia tricholepis"
  varSpecialConversions(3, 8) = "Muhlenbergia rigens"
  varSpecialConversions(4, 8) = Array("", "")
  varSpecialConversions(5, 8) = "Email Margaret March 6, 2018"
  varSpecialConversions(6, 8) = Array("", "")

  varSpecialConversions(0, 9) = "Q90"
  varSpecialConversions(1, 9) = -999
  varSpecialConversions(2, 9) = "Antennaria parvifolia"
  varSpecialConversions(3, 9) = "Antennaria rosulata"
  varSpecialConversions(4, 9) = Array("", "")
  varSpecialConversions(5, 9) = "Email Margaret March 6, 2018"
  varSpecialConversions(6, 9) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q64"
  varSpecialConversions(1, lngMaxIndex) = -999
  varSpecialConversions(2, lngMaxIndex) = "Elymus elymoides"
  varSpecialConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q64"
  varSpecialConversions(1, lngMaxIndex) = 2012
  varSpecialConversions(2, lngMaxIndex) = "Muhlenbergia montana"
  varSpecialConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q4"
  varSpecialConversions(1, lngMaxIndex) = -999
  varSpecialConversions(2, lngMaxIndex) = "Pascopyrum smithii"
  varSpecialConversions(3, lngMaxIndex) = "Elymus trachycaulus"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2012
  varSpecialConversions(2, lngMaxIndex) = "Penstemon virgatus"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q1"
  varSpecialConversions(1, lngMaxIndex) = 2019
  varSpecialConversions(2, lngMaxIndex) = "Chenopodium fremontii"
  varSpecialConversions(3, lngMaxIndex) = "Dysphania graveolens"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q26"
  varSpecialConversions(1, lngMaxIndex) = 2019
  varSpecialConversions(2, lngMaxIndex) = "Eremogone eastwoodiae"
  varSpecialConversions(3, lngMaxIndex) = "Eremogone fendleri"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q27"
  varSpecialConversions(1, lngMaxIndex) = 2019
  varSpecialConversions(2, lngMaxIndex) = "Eremogone eastwoodiae"
  varSpecialConversions(3, lngMaxIndex) = "Eremogone fendleri"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q46"
  varSpecialConversions(1, lngMaxIndex) = 2019
  varSpecialConversions(2, lngMaxIndex) = "Eremogone eastwoodiae"
  varSpecialConversions(3, lngMaxIndex) = "Eremogone fendleri"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q79"
  varSpecialConversions(1, lngMaxIndex) = 2019
  varSpecialConversions(2, lngMaxIndex) = "Eremogone eastwoodiae"
  varSpecialConversions(3, lngMaxIndex) = "Eremogone fendleri"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q88"
  varSpecialConversions(1, lngMaxIndex) = 2019
  varSpecialConversions(2, lngMaxIndex) = "Eremogone eastwoodiae"
  varSpecialConversions(3, lngMaxIndex) = "Eremogone fendleri"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe May 29, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2004
  varSpecialConversions(2, lngMaxIndex) = "Phlox longifolia"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe June 9, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2005
  varSpecialConversions(2, lngMaxIndex) = "Pennellia longifolia"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe June 9, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2006
  varSpecialConversions(2, lngMaxIndex) = "Pennellia longifolia"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe June 9, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2007
  varSpecialConversions(2, lngMaxIndex) = "Pennellia longifolia"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe June 9, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2009
  varSpecialConversions(2, lngMaxIndex) = "Pennellia longifolia"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe June 9, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q7"
  varSpecialConversions(1, lngMaxIndex) = 2010
  varSpecialConversions(2, lngMaxIndex) = "Pennellia longifolia"
  varSpecialConversions(3, lngMaxIndex) = "Penstemon linarioides"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Gabe June 9, 2020"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q98"  ' Black Springs, 30750
  varSpecialConversions(1, lngMaxIndex) = 2002
  varSpecialConversions(2, lngMaxIndex) = " "
  varSpecialConversions(3, lngMaxIndex) = "Cirsium wheeleri"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q50"  ' Wild Bill, Dispersed, 29016
  varSpecialConversions(1, lngMaxIndex) = 2020
  varSpecialConversions(2, lngMaxIndex) = " "
  varSpecialConversions(3, lngMaxIndex) = "Polygonum douglasii"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q50"  ' Wild Bill, Dispersed, 29016
  varSpecialConversions(1, lngMaxIndex) = -999
  varSpecialConversions(2, lngMaxIndex) = "Poa compressa"
  varSpecialConversions(3, lngMaxIndex) = "Poa pratensis"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q51"  ' Wild Bill, Dispersed, 29017
  varSpecialConversions(1, lngMaxIndex) = -999
  varSpecialConversions(2, lngMaxIndex) = "Poa compressa"
  varSpecialConversions(3, lngMaxIndex) = "Poa pratensis"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q50"  ' Wild Bill, Dispersed, 29017
  varSpecialConversions(1, lngMaxIndex) = 2014
  varSpecialConversions(2, lngMaxIndex) = "Muhlenbergia montana"
  varSpecialConversions(3, lngMaxIndex) = "Muhlenbergia minutissima"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q50"  ' Wild Bill, Dispersed, 29017
  varSpecialConversions(1, lngMaxIndex) = 2014
  varSpecialConversions(2, lngMaxIndex) = "Muhlenbergia montana"
  varSpecialConversions(3, lngMaxIndex) = "Muhlenbergia minutissima"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q67"  '
  varSpecialConversions(1, lngMaxIndex) = 2020
  varSpecialConversions(2, lngMaxIndex) = "Bouteloua arizonica"
  varSpecialConversions(3, lngMaxIndex) = "Bouteloua gracilis"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = UBound(varSpecialConversions, 2) + 1
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q3"  '
  varSpecialConversions(1, lngMaxIndex) = 2020
  varSpecialConversions(2, lngMaxIndex) = "Lathyrus graminifolius"
  varSpecialConversions(3, lngMaxIndex) = "Vicia pulchella"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = -1

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q6"
  varQueryConversions(1, lngMaxIndex) = 2020
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 25", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Discussion with MMM; June 11, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q59"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 16", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Discussion with MMM; June 11, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q95"
  varQueryConversions(1, lngMaxIndex) = 2016
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Packera multilobata"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q82"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Phlox longifolia"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q67"
  varQueryConversions(1, lngMaxIndex) = 2004
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = ""
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Delete")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q36"
  varQueryConversions(1, lngMaxIndex) = 2007
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = ""
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = ' ' OR ""species"" = ''", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, June 16, 2020"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Delete")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q54"
  varQueryConversions(1, lngMaxIndex) = 2002
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Geranium caespitosum"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q49"
  varQueryConversions(1, lngMaxIndex) = 2017
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia minutissima"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q45"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = ""
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Delete")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q41"
  varQueryConversions(1, lngMaxIndex) = 2005
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Trifolium longipes"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' ' OR ""species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q97"
  varQueryConversions(1, lngMaxIndex) = 2016
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Sporobolus interruptus"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = ' ' OR ""species"" = ''", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q86"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = ' ' OR ""species"" = ''", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q28"
  varQueryConversions(1, lngMaxIndex) = 2016
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = ' ' OR ""species"" = ''", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Gabe Traver, August 2, 2019"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2009
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" >= 2 AND ""FID"" <= 12", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" >= 3 AND ""FID"" <= 7", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" <= 7", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" >= 2 AND ""FID"" <= 5", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("(""FID"" <= 4) OR (""FID"" = 6) OR (""FID"" = 8) OR " & _
          "(""FID"" = 9) OR (""FID"" = 11)", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2014
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" <= 8", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q61"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("(""FID"" = 13) OR (""FID"" = 14) OR (""FID"" >= 16)", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q62"
  varQueryConversions(1, lngMaxIndex) = 2004
  varQueryConversions(2, lngMaxIndex) = "Unknown graminoid"
  varQueryConversions(3, lngMaxIndex) = "Unknown graminoid"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Unknown graminoid'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q62"
  varQueryConversions(1, lngMaxIndex) = 2004
  varQueryConversions(2, lngMaxIndex) = "Unknown graminoid"
  varQueryConversions(3, lngMaxIndex) = "Unknown graminoid"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Unknown'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2002
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2003
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2004
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2005
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2007
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2008
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2009
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2014
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2016
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" > -1", """FID"" > -1")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2009
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""x"" <= 0.75 OR ""y"" >= 0.2", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2009
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""x"" <= 0.75 OR ""y"" >= 0.2", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""x"" <= 0.7 OR ""y"" >= 0.2", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""x"" <= 0.7 OR ""y"" >= 0.2", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""x"" <= 0.7 OR ""y"" >= 0.2", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q64"
  varQueryConversions(1, lngMaxIndex) = 2002
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("Q64_2003_C|""species"" = 'Festuca arizonica'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Copy Features", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q65"
  varQueryConversions(1, lngMaxIndex) = 2002
  varQueryConversions(2, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(3, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pinus ponderosa'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q65"
  varQueryConversions(1, lngMaxIndex) = 2003
  varQueryConversions(2, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(3, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pinus ponderosa'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q65"
  varQueryConversions(1, lngMaxIndex) = 2004
  varQueryConversions(2, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(3, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pinus ponderosa'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q65"
  varQueryConversions(1, lngMaxIndex) = 2017
  varQueryConversions(2, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(3, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pinus ponderosa'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "From Jeff, Sept. 10, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q65"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Bole"
  varQueryConversions(3, lngMaxIndex) = "Bole"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Bole'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q65"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Pine bole"
  varQueryConversions(3, lngMaxIndex) = "Pine bole"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pine bole'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q42"
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Carex geophila"
  varQueryConversions(3, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 52 OR ""FID"" = 123", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q42"
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Carex geophila"
  varQueryConversions(3, lngMaxIndex) = "Sporobolus interruptus"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" <> 127 AND ""FID"" <> 52 AND ""FID"" <> 123", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q43"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Elymus elymoides' AND ""y"" > 0.75", _
        """species"" = 'Elymus elymoides' AND ""coords_x2"" > 0.75")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q43"
  varQueryConversions(1, lngMaxIndex) = 2016
  varQueryConversions(2, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 163", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q44"
  varQueryConversions(1, lngMaxIndex) = 2003
  varQueryConversions(2, lngMaxIndex) = "Carex geophila"
  varQueryConversions(3, lngMaxIndex) = "Sporobolus interruptus"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Carex geophila' AND NOT (""x"" > 0.8 AND ""y"" < 0.2)" & _
        "AND NOT ( ""x"" > 0.6 AND ""y"" > 0.6 AND ""y"" < 0.8)", _
        """species"" = 'Carex geophila' AND NOT  (""coords_x1"" > 0.8 AND ""coords_x2"" < 0.2) AND NOT " & _
        "( ""coords_x1"" > 0.6 AND ""coords_x2"" > 0.6 AND  ""coords_x2"" < 0.8) ")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q45"
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Sporobolus interruptus"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 4", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q45"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(3, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria rosulata'", _
        """species"" = 'Antennaria rosulata' AND ""coords_x1"" > 0.6 AND ""coords_x2""> 0.2")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q90"
  varQueryConversions(1, lngMaxIndex) = 2009
  varQueryConversions(2, lngMaxIndex) = "Sporobolus interruptus"
  varQueryConversions(3, lngMaxIndex) = "Poa fendleriana"
        "(""FID"" = 9) OR (""FID"" = 10) OR (""FID"" = 11) OR (""FID"" = 14) OR (""FID"" = 15) OR  " & _
        "(""FID"" = 19) OR (""FID"" = 31) OR (""FID"" = 36) OR (""FID"" = 38) OR (""FID"" = 40) OR  " & _
        "(""FID"" = 43) OR (""FID"" = 44) OR (""FID"" = 51) OR (""FID"" = 61) OR (""FID"" = 68) OR  " & _
        "(""FID"" = 82) OR (""FID"" = 86) OR (""FID"" = 108) OR (""FID"" = 109) OR (""FID"" = 110)  " & _
        "OR (""FID"" = 111) OR (""FID"" = 112) OR (""FID"" = 113) OR (""FID"" = 116) OR  " & _
        "(""FID"" = 117) OR (""FID"" = 118) OR (""FID"" = 119) OR (""FID"" = 120) OR (""FID"" = 121)  " & _
        "OR (""FID"" = 123) OR (""FID"" = 124) OR (""FID"" = 125) OR (""FID"" = 126) OR (""FID"" = 127)  " & _
        "OR (""FID"" = 128) OR (""FID"" = 129) OR (""FID"" = 131) OR (""FID"" = 132) OR  " & _
        "(""FID"" = 133) OR (""FID"" = 134) OR (""FID"" = 135) OR (""FID"" = 136)", """FID"" = -999")
  varQueryConversions(4, lngMaxIndex) = Array("(""FID"" = 5) OR (""FID"" = 6) OR (""FID"" = 8) OR  " & _
        "(""FID"" = 9) OR (""FID"" = 11) OR (""FID"" = 15) OR  " & _
        "(""FID"" = 19) OR(""FID"" = 36) OR (""FID"" = 38) OR (""FID"" = 40) OR  " & _
        "(""FID"" = 44) OR (""FID"" = 51) OR (""FID"" = 61) OR (""FID"" = 68) OR  " & _
        "(""FID"" = 82) OR (""FID"" = 86) OR (""FID"" = 108) OR (""FID"" = 109) OR (""FID"" = 110)  " & _
        "OR (""FID"" = 111) OR (""FID"" = 112) OR (""FID"" = 113) OR  " & _
        "(""FID"" = 117) OR (""FID"" = 118) OR (""FID"" = 119) OR (""FID"" = 120) OR (""FID"" = 121)  " & _
        "OR (""FID"" = 123) OR (""FID"" = 124) OR (""FID"" = 125) OR (""FID"" = 126) OR (""FID"" = 127)  " & _
        "OR (""FID"" = 128) OR (""FID"" = 129) OR (""FID"" = 131) OR (""FID"" = 132) OR  " & _
        "(""FID"" = 133) OR (""FID"" = 134) OR (""FID"" = 135) OR (""FID"" = 136)", """FID"" = 28")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q32"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Festuca arizonica'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q32"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Elymus elymoides' AND ""x"" > 0.2", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q80"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Muhlenbergia minutissima"
  varQueryConversions(3, lngMaxIndex) = "Sporobolus interruptus"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 50 OR ""FID"" = 49 OR ""FID"" = 19", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q80"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Muhlenbergia minutissima"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua gracilis"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 26", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q80"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Muhlenbergia minutissima"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia wrightii"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" <> 50 AND ""FID"" <> 49 AND ""FID"" <> 19 AND ""FID"" <> 26 AND " & _
        """species"" = 'Muhlenbergia minutissima'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q80"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia wrightii"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Muhlenbergia montana'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q85"
  varQueryConversions(1, lngMaxIndex) = 2007
  varQueryConversions(2, lngMaxIndex) = "Elymus trachycaulum"
  varQueryConversions(3, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Elymus trachycaulum'", """species"" = 'Elymus trachycaulum'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q85"
  varQueryConversions(1, lngMaxIndex) = 2007
  varQueryConversions(2, lngMaxIndex) = "Koeleria macrantha"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Koeleria macrantha'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q86"
  varQueryConversions(1, lngMaxIndex) = 2002
  varQueryConversions(2, lngMaxIndex) = "Antennaria parviflora"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria parviflora'", """species"" = 'Antennaria parviflora'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q86"
  varQueryConversions(1, lngMaxIndex) = 2002
  varQueryConversions(2, lngMaxIndex) = "Erigeron formosissimus"
  varQueryConversions(3, lngMaxIndex) = "Symphyotrichum ascendens"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = 'Erigeron formosissimus'")
  varQueryConversions(5, lngMaxIndex) = "From Jeff: Sept. 10, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q87"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia wrightii"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Muhlenbergia montana'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q87"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Muhlenbergia rigens"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia wrightii"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Muhlenbergia rigens'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q88"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Unknown Forb"
  varQueryConversions(3, lngMaxIndex) = "Coreopsis tinctoria"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" LIKE '%Unknown%'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q41"
  varQueryConversions(1, lngMaxIndex) = 2004
  varQueryConversions(2, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria parvifolia'", """species"" = 'Antennaria parvifolia'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q41"
  varQueryConversions(1, lngMaxIndex) = 2005
  varQueryConversions(2, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria parvifolia'", """species"" = 'Antennaria parvifolia'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q41"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria parvifolia'", """species"" = 'Antennaria parvifolia'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q41"
  varQueryConversions(1, lngMaxIndex) = 2007
  varQueryConversions(2, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria parvifolia'", """species"" = 'Antennaria parvifolia'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q47"
  varQueryConversions(1, lngMaxIndex) = 2007
  varQueryConversions(2, lngMaxIndex) = "Symphyotrichum ascendens"
  varQueryConversions(3, lngMaxIndex) = "Erigeron formosissimus"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = 'Symphyotrichum ascendens'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q69"
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" >= 49 AND ""FID"" <= 52", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q69"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Poa fendleriana'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q71"
  varQueryConversions(1, lngMaxIndex) = 2013
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Elymus elymoides' AND " & _
        "NOT( ""FID"" = 3 OR ""FID"" = 41 OR ""FID"" = 44 OR ""FID"" = 45 OR ""FID"" = 46)", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q73"
  varQueryConversions(1, lngMaxIndex) = 2005
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Festuca arizonica'", """species"" = 'Festuca arizonica'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q73"
  varQueryConversions(1, lngMaxIndex) = 2006
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = 'Festuca arizonica'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q77"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Epilobium brachycarpum"
  varQueryConversions(3, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Epilobium brachycarpum'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q53"
  varQueryConversions(1, lngMaxIndex) = 2012
  varQueryConversions(2, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Poa fendleriana'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q59"
  varQueryConversions(1, lngMaxIndex) = 2005
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Festuca arizonica' AND ""x"" >= 0.2 AND " & _
      """x"" <= 0.6 AND ""y"" >= 0.6", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q13"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Pine bole"
  varQueryConversions(3, lngMaxIndex) = "Pine bole"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pine bole'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q24"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Solidago velutina"
  varQueryConversions(3, lngMaxIndex) = "Erigeron formosissimus"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = 'Solidago velutina'")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q48"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(3, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = 0", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q49"
  varQueryConversions(1, lngMaxIndex) = 2010
  varQueryConversions(2, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("(""species"" = 'Antennaria parvifolia' AND (( ""x"" <= 0.6 and " & _
      """y"" >= 0.6) OR (""x"" <= 0.4 AND ""y"" <= 0.4)) AND ""FID"" <> 58 AND ""FID"" <> 51) OR ""FID"" = 73 OR ""FID"" = 74", _
      """species"" = 'Antennaria parvifolia' AND (( ""coords_x1"" <= 0.6 and ""coords_x2"" >= 0.6) OR " & _
      "(""coords_x1"" <= 0.4 AND ""coords_x2"" <= 0.4)) ")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q52"
  varQueryConversions(1, lngMaxIndex) = 2015
  varQueryConversions(2, lngMaxIndex) = "Pine bole"
  varQueryConversions(3, lngMaxIndex) = "Pine bole"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pine bole'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = UBound(varQueryConversions, 2) + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q50"  ' Wild Bill, Dispersed, 29016
  varQueryConversions(1, lngMaxIndex) = 2011
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varQueryConversions(4, lngMaxIndex) = Array("species = 'Elymus elymoides' AND x > 0.81 AND x < 0.82 AND y > 0.77 AND y < 0.78", "")
  varQueryConversions(5, lngMaxIndex) = "Email MMM, April 10, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = UBound(varQueryConversions, 2) + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q54"  ' Wild Bill, Dispersed, 29016
  varQueryConversions(1, lngMaxIndex) = 2021
  varQueryConversions(2, lngMaxIndex) = ""
  varQueryConversions(3, lngMaxIndex) = "Muhlenbergia montana"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = ''", "")
  varQueryConversions(5, lngMaxIndex) = "Email Oliver Pickering, Feb. 21, 2022"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = UBound(varQueryConversions, 2) + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q67"  ' Big Fill, 30719
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Gutierrezia sarothrae"
  varQueryConversions(3, lngMaxIndex) = "Gutierrezia sarothrae"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Gutierrezia sarothrae'", "")
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Return Centroid", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q37"
  varQueryConversions(1, lngMaxIndex) = 2022
  varQueryConversions(2, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(3, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria rosulata'", "")
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q87"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Elymus elymoides"
  varQueryConversions(3, lngMaxIndex) = "Eremogone fendleri"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Elymus elymoides'", """species"" = 'Elymus elymoides'")
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q15"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Pinus ponderosa"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Pinus ponderosa'", "") ' There's also a PIPO seedling in the
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q21"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Poa fendleriana"
  varQueryConversions(3, lngMaxIndex) = "Festuca arizonica"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Poa fendleriana'", "")
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q27"
  varQueryConversions(1, lngMaxIndex) = 2020
  varQueryConversions(2, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(3, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria parvifolia'", "")
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q48"
  varQueryConversions(1, lngMaxIndex) = 2022
  varQueryConversions(2, lngMaxIndex) = "Antennaria rosulata"
  varQueryConversions(3, lngMaxIndex) = "Antennaria parvifolia"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria rosulata'", "")
  varQueryConversions(5, lngMaxIndex) = "Email Wade Gibson , Jan. 4, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q42"
  varQueryConversions(1, lngMaxIndex) = -999
  varQueryConversions(2, lngMaxIndex) = "Carex geophila"
  varQueryConversions(3, lngMaxIndex) = "Carex rossii"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = 'Antennaria rosulata'", "")
  varQueryConversions(5, lngMaxIndex) = "Email Margaret Moore , Feb. 17, 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  CreateVarSpecialConversions = varSpecialConversions
  Erase varSpecialConversions

End Function

Public Sub ReviseShapefiles()

  Dim strQuadrats() As String
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim lngFeatCount As Long
  Dim pQuadData As Collection
  Dim varSites() As Variant
  Dim varSitesSpecifics() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev(strQuadrats, pPlotToQuadratConversion, pQuadratToPlotConversion, _
      varSites, varSitesSpecifics)

  Dim strItems() As String
  Dim strNote As String

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim lngMaxIndex As Long

  Dim varSpecialConversions() As Variant
  Dim varQueryConversions() As Variant
  ReDim varSpecialConversions(6, 9)
  Dim strNoteOnChanges As String

  varSpecialConversions = CreateVarSpecialConversions(varQueryConversions)

  Dim pCoverCollection As New Collection
  Dim pDensityCollection As New Collection

  Dim pCoverToDensity As Collection
  Dim pDensityToCover As Collection
  Dim strCoverToDensityQuery As String
  Dim strDensityToCoverQuery As String
  Dim pCoverShouldChangeColl As Collection
  Dim pDensityShouldChangeColl As Collection

  Debug.Print "---------------------"
  Call FillCollections(pCoverCollection, pDensityCollection, pCoverToDensity, pDensityToCover, _
    strCoverToDensityQuery, strDensityToCoverQuery, pCoverShouldChangeColl, pDensityShouldChangeColl)

  Dim pQueryFilt As IQueryFilter
  Dim strQueryString As String
  Dim varStrings() As Variant

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String
  Call DeclareWorkspaces(strRoot, , , , , strContainingFolder)

  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long

  For lngIndex = 0 To pFolders.Count - 1
    Debug.Print CStr(lngIndex) & pFolders.Element(lngIndex)
  Next lngIndex

  Dim pDataset As IDataset
  Dim booFoundShapefiles As Boolean
  Dim varDatasets() As Variant

  Dim strNames() As String
  Dim strName As String
  Dim lngDatasetIndex As Long
  Dim lngNameIndex As Long
  Dim lngNameCount As Long
  Dim booFoundNames As Boolean
  Dim lngRecCount As Long

  Dim strFullNames() As String
  Dim lngFullNameCounter As Long

  Dim lngShapefileCount As Long
  Dim lngAcceptSFCount As Long
  lngShapefileCount = 0
  lngRecCount = 0
  lngAcceptSFCount = 0

  lngFullNameCounter = -1
  Dim pNameColl As New Collection
  Dim strHexify As String
  Dim strCorrect As String
  Dim pCheckCollection As Collection
  Dim strReport As String
  Dim booMadeChanges As Boolean
  Dim strEditReport As String
  Dim strExcelReport As String
  Dim strExcelFullReport As String
  Dim pFClass As IFeatureClass
  Dim strBase As String
  Dim strSplit() As String

  strExcelFullReport = """Shapefile""" & vbTab & """Quadrat""" & vbTab & """Year""" & vbTab & _
      """Type""" & vbTab & """Feature_ID""" & vbTab & """Original""" & vbTab & """Changed_To""" & vbCrLf

  For lngIndex = 0 To pFolders.Count - 1
    DoEvents
    strFolder = pFolders.Element(lngIndex)
    varDatasets = ReturnFeatureClassesOrNothing(strFolder, booFoundShapefiles)

    Debug.Print CStr(lngIndex + 1) & " of " & CStr(pFolders.Count) & "] " & strFolder
    If booFoundShapefiles Then
      Debug.Print "  --> Found Shapefiles = " & CStr(booFoundShapefiles) & " [n = " & CStr(UBound(varDatasets) + 1) & "]"

      lngShapefileCount = lngShapefileCount + UBound(varDatasets) + 1

      For lngDatasetIndex = 0 To UBound(varDatasets)
        Set pDataset = varDatasets(lngDatasetIndex)

        If Right(pDataset.BrowseName, 2) = "_D" Then
          Set pCheckCollection = pDensityCollection
        ElseIf Right(pDataset.BrowseName, 2) = "_C" Then
          Set pCheckCollection = pCoverCollection
        Else
          MsgBox "Unexpected Dataset Name!"
          DoEvents
        End If
        strSplit = Split(pDataset.BrowseName, "_")

        strBase = """" & pDataset.BrowseName & """" & vbTab & """" & strSplit(0) & """" & vbTab & _
            """" & strSplit(1) & """" & vbTab & """" & IIf(strSplit(2) = "C", "Cover", "Density") & """"

        Set pFClass = pDataset
        If pFClass.FindField("Cover") > -1 Or pFClass.FindField("Species") > -1 Then

          Call ReplaceNamesInShapefile(pDataset, pCheckCollection, booMadeChanges, strEditReport, strBase, _
              strExcelReport, varSpecialConversions, varQueryConversions)

          If booMadeChanges Then
            strReport = strReport & strEditReport
            strExcelFullReport = strExcelFullReport & strExcelReport
          Else
            strReport = strReport & "No changes to '" & pDataset.BrowseName & "'..." & vbCrLf
            strExcelFullReport = strExcelFullReport & strBase & vbTab & """<- No Changes ->""" & vbTab & vbTab & vbCrLf
          End If

        End If
      Next lngDatasetIndex

    End If

  Next lngIndex

  strReport = strReport & vbCrLf & "Done..." & vbCrLf & _
    MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

  Dim pDataObj As New MSForms.DataObject

  strExcelFullReport = Replace(strExcelFullReport, vbTab, ",")
  MyGeneralOperations.WriteTextFile strContainingFolder & "\Log_of_Changes_" & MyGeneralOperations.ReturnTimeStamp & ".csv", strExcelFullReport

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Erase strQuadrats
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Set pQuadData = Nothing
  Erase strItems
  Erase varSpecialConversions
  Erase varQueryConversions
  Set pCoverCollection = Nothing
  Set pDensityCollection = Nothing
  Set pCoverToDensity = Nothing
  Set pDensityToCover = Nothing
  Set pCoverShouldChangeColl = Nothing
  Set pDensityShouldChangeColl = Nothing
  Set pQueryFilt = Nothing
  Erase varStrings
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pCheckCollection = Nothing
  Set pFClass = Nothing
  Erase strSplit
  Set pDataObj = Nothing

End Sub

Public Sub AddVerbatimFields(pFClass As IFeatureClass, pQuadData As Collection)

  Dim strName As String
  Dim pFCursor As IFeatureCursor
  Dim lngSrcSpeciesNameIndex As Long
  Dim lngVerbSpeciesNameIndex As Long
  Dim lngRotationNameIndex As Long
  Dim lngVerbTypeIndex As Long
  Dim lngSiteIndex As Long
  Dim lngPlotIndex As Long

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pFeature As IFeature
  Dim pDataset As IDataset

  Dim strQuad As String
  Dim strItems() As String
  Dim strSite As String
  Dim strPlot As String
  Dim strFileHeader As String
  Set pDataset = pFClass

  strQuad = aml_func_mod.ReturnFilename2(pDataset.Workspace.PathName)
  strQuad = Replace(strQuad, "Q", "", , , vbTextCompare)
  strItems = pQuadData.Item(strQuad)
  strSite = Trim(strItems(1))
  If strSite = "" Then
    strSite = Trim(strItems(0))
  End If
  strPlot = Trim(strItems(2))
  strFileHeader = Trim(strItems(5))

  lngSrcSpeciesNameIndex = pFClass.FindField("Species")
  lngVerbSpeciesNameIndex = pFClass.FindField("Verb_Spcs")
  lngRotationNameIndex = pFClass.FindField("Revise_Rtn")
  lngSiteIndex = pFClass.FindField("Site")
  lngPlotIndex = pFClass.FindField("Plot")

  If lngSiteIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Site"
      .Type = esriFieldTypeString
      .length = 75
    End With
    pFClass.AddField pField
    lngSiteIndex = pFClass.FindField("Site")
  End If
  If lngPlotIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Plot"
      .Type = esriFieldTypeString
      .length = 25
    End With
    pFClass.AddField pField
    lngPlotIndex = pFClass.FindField("Plot")
  End If
  If lngVerbSpeciesNameIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Verb_Spcs"
      .Type = esriFieldTypeString
      .length = 50
    End With
    pFClass.AddField pField
    lngVerbSpeciesNameIndex = pFClass.FindField("Verb_Spcs")
  End If
  lngVerbTypeIndex = pFClass.FindField("Verb_Type")
  If lngVerbTypeIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Verb_Type"
      .Type = esriFieldTypeString
      .length = 50
    End With
    pFClass.AddField pField
    lngVerbTypeIndex = pFClass.FindField("Verb_Type")
  End If

  lngRotationNameIndex = pFClass.FindField("Revise_Rtn")
  If lngRotationNameIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Revise_Rtn"
      .Type = esriFieldTypeDouble
      .Precision = 12
      .Scale = 6
    End With
    pFClass.AddField pField
    lngRotationNameIndex = pFClass.FindField("Revise_Rtn")
  End If

  Set pFCursor = pFClass.Update(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pFeature.Value(lngVerbSpeciesNameIndex) = pFeature.Value(lngSrcSpeciesNameIndex)
    If StrComp(Right(pDataset.BrowseName, 2), "_C", vbTextCompare) = 0 Then
      pFeature.Value(lngVerbTypeIndex) = "Cover"
    Else
      pFeature.Value(lngVerbTypeIndex) = "Density"
    End If
    pFeature.Value(lngSiteIndex) = strSite
    pFeature.Value(lngPlotIndex) = strPlot
    pFCursor.UpdateFeature pFeature
    Set pFeature = pFCursor.NextFeature
  Loop
  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pFeature = Nothing
  Set pDataset = Nothing

End Sub

Public Sub FillCollections(pCoverCollection As Collection, pDensityCollection As Collection, _
    Optional pCoverToDensity As Collection, Optional pDensityToCover As Collection, _
    Optional strCoverToDensityQuery As String, Optional strDensityToCoverQuery As String, _
    Optional pCoverShouldChangeColl As Collection, Optional pDensityShouldChangeColl As Collection)

  Dim pRedigitizeColl As Collection
  Set pRedigitizeColl = ReturnReplacementColl

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pTable As ITable
  Set pWSFact = New ExcelWorkspaceFactory

  Dim pTestWS As IFeatureWorkspace
  Dim pTestWSFact As IWorkspaceFactory
  Set pTestWSFact = New FileGDBWorkspaceFactory
  Set pTestWS = pTestWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data\Combined_by_Quadrat.gdb", 0)
  Dim pTestFC As IFeatureClass
  Set pTestFC = pTestWS.OpenFeatureClass("Box")
  Dim strPrefix As String
  Dim strSuffix As String
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pTestFC, strPrefix, strSuffix)

  Dim lngCorrectIndex As Long
  Dim lngIncorrectIndex As Long
  Dim strCorrect As String
  Dim strIncorrect As String
  Dim strHexCorrect As String
  Dim strHexIncorrect As String
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngShouldChangeIndex As Long
  Dim strShouldChange As String
  Dim booShouldChange As Boolean

  Dim varPaths As Variant
  Dim varColls As Variant
  Dim varVals As Variant
  Dim varShouldChange As Variant
  Dim strCover() As String
  Dim strDensity() As String

  varPaths = Array("D:\arcGIS_stuff\consultation\Margaret_Moore\species_list_Cover_changes_Dec_2_2017.xlsx", _
                   "D:\arcGIS_stuff\consultation\Margaret_Moore\Species_list_Density_changes_Dec_2_2017.xlsx")
  varColls = Array(pCoverCollection, pDensityCollection)
  varVals = Array(strCover, strDensity)
  Set pCoverShouldChangeColl = New Collection
  Set pDensityShouldChangeColl = New Collection
  varShouldChange = Array(pCoverShouldChangeColl, pDensityShouldChangeColl)

  Dim lngIndex As Long
  Dim strPath As String
  Dim pColl As Collection
  Dim strVals() As String
  Dim varVal As Variant
  Dim strIncorrectVariant As String

  Dim pFromCoverColl As New Collection
  Dim strFromCoverVals() As String
  Dim pFromDensColl As New Collection
  Dim pShouldChangeColl As New Collection
  Dim strFromDensVals() As String
  Dim lngFromCoverCounter As Long
  Dim lngFromDensCounter As Long
  Dim booShouldChangeFromCover As Boolean
  Dim booShouldChangeFromDensity As Boolean

  Set pCoverToDensity = New Collection
  Set pDensityToCover = New Collection
  strCoverToDensityQuery = ""
  strDensityToCoverQuery = ""
  lngFromCoverCounter = -1
  lngFromDensCounter = -1

  Dim lngIndex2 As Long

  For lngIndex = 0 To UBound(varPaths)
    strPath = varPaths(lngIndex)
    Set pColl = varColls(lngIndex)
    strVals = varVals(lngIndex)
    Set pShouldChangeColl = varShouldChange(lngIndex)
    lngIndex2 = -1

    Set pWS = pWSFact.OpenFromFile(strPath, 0)
    Set pTable = pWS.OpenTable("For_ArcGIS_Dec_2017$")
    lngCorrectIndex = pTable.FindField("Correct")
    lngIncorrectIndex = pTable.FindField("Incorrect")
    lngShouldChangeIndex = pTable.FindField("Comment")
    If lngShouldChangeIndex = -1 Then lngShouldChangeIndex = pTable.FindField("Comments1")
    Debug.Print CStr(lngIndex) & "] " & IIf(lngIndex = 0, "Cover", "Density") & " Record Count = " & CStr(pTable.RowCount(Nothing))
    Set pCursor = pTable.Search(Nothing, False)
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      strCorrect = ""
      strIncorrect = ""
      strShouldChange = ""
      booShouldChangeFromCover = False
      booShouldChangeFromDensity = False
      booShouldChange = False

      varVal = pRow.Value(lngCorrectIndex)
      If Not IsNull(varVal) Then strCorrect = CStr(varVal)
      varVal = pRow.Value(lngIncorrectIndex)
      If Not IsNull(varVal) Then strIncorrect = CStr(varVal)
      varVal = pRow.Value(lngShouldChangeIndex)
      If Not IsNull(varVal) Then strShouldChange = Trim(CStr(varVal))

      If strIncorrect = "Drymaria leptophyllum" Then
        DoEvents
      End If

      If InStr(1, strShouldChange, "change to", vbTextCompare) > 0 Then booShouldChange = True

      Set pRow = pCursor.NextRow

      If strShouldChange <> "" Then

        If strShouldChange = "change to point shapefile" Then
          booShouldChangeFromCover = True
        ElseIf strShouldChange = "change to polygon shapefile" Then
          booShouldChangeFromDensity = True
        End If

        If lngIndex = 0 Then  ' cover
          If Not MyGeneralOperations.CheckCollectionForKey(pFromCoverColl, strShouldChange) Then
            lngFromCoverCounter = lngFromCoverCounter + 1
            ReDim Preserve strFromCoverVals(lngFromCoverCounter)
            strFromCoverVals(lngFromCoverCounter) = strShouldChange
            pFromCoverColl.Add True, strShouldChange
          End If
        ElseIf lngIndex = 1 Then  ' density
          If Not MyGeneralOperations.CheckCollectionForKey(pFromDensColl, strShouldChange) Then
            lngFromDensCounter = lngFromDensCounter + 1
            ReDim Preserve strFromDensVals(lngFromDensCounter)
            strFromDensVals(lngFromDensCounter) = strShouldChange
            pFromDensColl.Add True, strShouldChange
          End If
        End If
      End If

      If InStr(1, strIncorrect, "Erigeron formosissimus", vbTextCompare) > 0 Then
        DoEvents
      End If

      strCorrect = Replace(strCorrect, Chr(Asc(vbCrLf)), "")
      strCorrect = Replace(strCorrect, Chr(Asc(vbNewLine)), "")
      strCorrect = Trim(strCorrect)

      If InStr(1, strCorrect, "Muhlenbergia tricholepis", vbTextCompare) > 0 Then
        DoEvents
      End If

      If InStr(1, strCorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbTab)), vbTextCompare) > 0 Then
        Debug.Print "...strCorrect = " & strCorrect
      End If
      If InStr(1, strIncorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbTab)), vbTextCompare) > 0 Then
        Debug.Print "...strIncorrect = " & strIncorrect
      End If

      If strCorrect = "" Then
        strHexIncorrect = HexifyName(strIncorrect)
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexIncorrect) Then
           pShouldChangeColl.Add booShouldChange, strHexIncorrect  ' strHexIncorrect is the correct name in this case
        End If

      ElseIf strCorrect <> "" Then
        strHexCorrect = HexifyName(strCorrect)
        strHexIncorrect = HexifyName(strIncorrect)
        If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
          pColl.Add strCorrect, strHexIncorrect
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
          pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
        End If

        lngIndex2 = lngIndex2 + 1
        ReDim Preserve strVals(lngIndex2)
        strVals(lngIndex2) = strIncorrect

        If lngIndex = 0 And booShouldChangeFromCover Then
          If Not MyGeneralOperations.CheckCollectionForKey(pCoverToDensity, strHexIncorrect) Then
            pCoverToDensity.Add strCorrect, strHexIncorrect
          End If

          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          If Not MyGeneralOperations.CheckCollectionForKey(pDensityToCover, strHexIncorrect) Then
            pDensityToCover.Add strCorrect, strHexIncorrect
          End If
          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

        End If

        If InStr(1, strCorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Then
          strIncorrectVariant = strIncorrect

          strCorrect = Replace(strCorrect, Chr(Asc(vbCrLf)), "")
          strCorrect = Replace(strCorrect, Chr(Asc(vbNewLine)), "")
          strIncorrectVariant = Replace(strIncorrectVariant, Chr(Asc(vbCrLf)), "")
          strIncorrectVariant = Replace(strIncorrectVariant, Chr(Asc(vbNewLine)), "")

          strHexCorrect = HexifyName(strCorrect)
          strHexIncorrect = HexifyName(strIncorrectVariant)
          If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
            pColl.Add strCorrect, strHexIncorrect

            lngIndex2 = lngIndex2 + 1
            ReDim Preserve strVals(lngIndex2)
            strVals(lngIndex2) = strIncorrectVariant

            If lngIndex = 0 And booShouldChangeFromCover Then
              pCoverToDensity.Add strCorrect, strHexIncorrect

              strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

            ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
              pDensityToCover.Add strCorrect, strHexIncorrect

              strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
            End If

          End If
          If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
            pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
          End If
        End If

        If Trim(strIncorrect) <> strIncorrect Then
          strIncorrectVariant = Trim(strIncorrect)

          strHexIncorrect = HexifyName(strIncorrectVariant)
          If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
            pColl.Add strCorrect, strHexIncorrect

            lngIndex2 = lngIndex2 + 1
            ReDim Preserve strVals(lngIndex2)
            strVals(lngIndex2) = strIncorrectVariant

            If lngIndex = 0 And booShouldChangeFromCover Then
              pCoverToDensity.Add strCorrect, strHexIncorrect

              strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

            ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
              pDensityToCover.Add strCorrect, strHexIncorrect

              strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
            End If
          End If
          If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
            pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
          End If
        End If

      Else

        If lngIndex = 0 And booShouldChangeFromCover Then
          pCoverToDensity.Add strIncorrect, strIncorrect

          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strIncorrect & "'"

        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          pDensityToCover.Add strIncorrect, strIncorrect

          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strIncorrect & "'"

        End If
      End If
    Loop
    varVals(lngIndex) = strVals
  Next lngIndex

  Debug.Print "Checking From Cover to Density Values:"
  For lngIndex = 0 To lngFromCoverCounter
    Debug.Print "  " & CStr(lngIndex + 1) & "] " & strFromCoverVals(lngIndex)
  Next lngIndex
  Debug.Print "Checking From Density to Cover Values:"
  For lngIndex = 0 To lngFromDensCounter
    Debug.Print "  " & CStr(lngIndex + 1) & "] " & strFromDensVals(lngIndex)
  Next lngIndex

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  varPaths = Null
  varColls = Null
  Set pColl = Nothing

End Sub

Public Function ReturnReplacementColl() As Collection

  Dim pReturn As New Collection
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory

  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Newly_Georeferenced_Aug_2018\New_Feature_Classes.gdb", 0)

  Dim pDataset As IDataset
  Dim pEnumDataset As IEnumDataset

  Dim strDatasetName As String

  Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
  Set pDataset = pEnumDataset.Next
  Do Until pDataset Is Nothing
    strDatasetName = pDataset.BrowseName
    Select Case strDatasetName
      Case "BS_2004_46_C"
        pReturn.Add pDataset, "Q9_2015_C"
      Case "BS_2004_46_D"
        pReturn.Add pDataset, "Q9_2015_D"
      Case Else
        pReturn.Add pDataset, strDatasetName
    End Select
    Set pDataset = pEnumDataset.Next
  Loop

  Set ReturnReplacementColl = pReturn

  Set pReturn = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDataset = Nothing
  Set pEnumDataset = Nothing

End Function

Public Function HexifyName(strName As String) As String
  Dim lngIndex As Long
  HexifyName = Space$(Len(strName) * 4)
  For lngIndex = 0 To Len(strName) - 1
      Mid$(HexifyName, lngIndex * 4 + 1, 4) = Right$("0000" & Hex$(AscW(Mid$(strName, lngIndex + 1, 1))), 4)
  Next lngIndex
End Function

Public Function ReturnFeatureClassesOrNothing(strFolder As String, booWorked As Boolean, _
    Optional booFoundPolygonFClass As Boolean, Optional booFoundPointFClass As Boolean, _
    Optional pRepPointFClass As IFeatureClass, Optional pRepPolyFClass As IFeatureClass) As Variant()

  On Error GoTo ErrHandler

  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFolder, 0)
  Dim pEnumDataset As IEnumDataset
  Dim lngIndex As Long
  Dim varReturn() As Variant

  Dim pDataset As IDataset
  Dim pFClass As IFeatureClass

  booFoundPolygonFClass = False
  booFoundPointFClass = False

  lngIndex = -1

  Dim lngMaxPolygonFCount As Long
  Dim lngMaxPointFCount As Long
  lngMaxPolygonFCount = -1
  lngMaxPointFCount = -1

  Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
  pEnumDataset.Reset
  Set pDataset = pEnumDataset.Next
  Do Until pDataset Is Nothing
    lngIndex = lngIndex + 1
    ReDim Preserve varReturn(lngIndex)
    Set pFClass = pDataset
    If pFClass.ShapeType = esriGeometryPoint Then
      booFoundPointFClass = True
      If pFClass.Fields.FieldCount > lngMaxPointFCount Then
        Set pRepPointFClass = pFClass
        lngMaxPointFCount = pFClass.Fields.FieldCount
      End If
    ElseIf pFClass.ShapeType = esriGeometryPolygon Then
      booFoundPolygonFClass = True
      If pFClass.Fields.FieldCount > lngMaxPolygonFCount Then
        Set pRepPolyFClass = pFClass
        lngMaxPolygonFCount = pFClass.Fields.FieldCount
      End If
    End If
    Set varReturn(lngIndex) = pDataset
    Set pDataset = pEnumDataset.Next
  Loop

  ReturnFeatureClassesOrNothing = varReturn

  booWorked = lngIndex > -1
  GoTo ClearMemory
  Exit Function

ErrHandler:
  booWorked = False

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pEnumDataset = Nothing
  Erase varReturn
  Set pDataset = Nothing
  Set pFClass = Nothing

End Function

Public Function CopyFeatureClassToShapefile(pFClass As IFeatureClass, strPath As String) As IFeatureClass

  Dim strTemp As String
  strTemp = aml_func_mod.SetExtension(strPath, "shp")
  If aml_func_mod.ExistFileDir(strTemp) Then
    Debug.Print "...Feature Class '" & strPath & "' already exists.  Did not export..."
    GoTo ClearMemory
  End If

  Dim strDir As String
  Dim strFClassName As String
  strDir = aml_func_mod.ReturnDir3(strPath, False)
  strFClassName = aml_func_mod.ReturnFilename2(strPath)
  strFClassName = aml_func_mod.ClipExtension2(strFClassName)

  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Dim pFieldArray As esriSystem.IVariantArray
  Dim pField As iField
  Dim lngIndex As Long
  Dim pClone As IClone
  Dim pFieldEdit As IFieldEdit
  Dim pNewField As iField
  Dim pSourceNameColl As New Collection
  Dim strNewName As String
  Dim strSourceName As String

  Set pFieldArray = New esriSystem.varArray

  For lngIndex = 0 To pFClass.Fields.FieldCount - 1
    Set pField = pFClass.Fields.Field(lngIndex)
    If pField.Editable And pField.Type <> esriFieldTypeGeometry Then
      Set pClone = pField
      strNewName = MyGeneralOperations.ReturnAcceptableFieldName2(pField.Name, pFieldArray, True, False, False, False)
      pSourceNameColl.Add pField.Name, strNewName
      Set pNewField = pClone.Clone
      Set pFieldEdit = pNewField
      With pFieldEdit
        .Name = strNewName
      End With
      pFieldArray.Add pNewField
    End If
  Next lngIndex

  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference

  Dim pNewFClass As IFeatureClass
  Set pNewFClass = MyGeneralOperations.CreateShapefileFeatureClass2(strDir, strFClassName, pSpRef, _
      pFClass.ShapeType, pFieldArray, False)

  Dim lngLinks() As Long
  Dim lngSrcIndex As Long
  Dim lngArrayIndex As Long
  lngArrayIndex = -1
  For lngIndex = 0 To pNewFClass.Fields.FieldCount - 1
    Set pNewField = pNewFClass.Fields.Field(lngIndex)
    If pNewField.Type <> esriFieldTypeGeometry Then
      strNewName = pNewField.Name
      If MyGeneralOperations.CheckCollectionForKey(pSourceNameColl, strNewName) Then
        strSourceName = pSourceNameColl.Item(strNewName)
        lngArrayIndex = lngArrayIndex + 1
        ReDim Preserve lngLinks(1, lngArrayIndex)
        lngLinks(0, lngArrayIndex) = pFClass.FindField(strSourceName)
        lngLinks(1, lngArrayIndex) = pNewFClass.FindField(strNewName)
      End If
    End If
  Next lngIndex

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pNewFCursor As IFeatureCursor
  Dim pNewFBuffer As IFeatureBuffer
  Dim lngCounter As Long

  Set pFCursor = pFClass.Search(Nothing, False)
  Set pNewFCursor = pNewFClass.Insert(True)
  Set pNewFBuffer = pNewFClass.CreateFeatureBuffer

  Dim varVal As Variant

  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pNewFBuffer.Shape = pFeature.ShapeCopy
    For lngIndex = 0 To UBound(lngLinks, 2)
      varVal = pFeature.Value(lngLinks(0, lngIndex))
      If IsNull(varVal) Then
        If pNewFBuffer.Fields.Field(lngLinks(1, lngIndex)).Type = esriFieldTypeString Then
          pNewFBuffer.Value(lngLinks(1, lngIndex)) = ""
        Else
          pNewFBuffer.Value(lngLinks(1, lngIndex)) = -999
        End If
      Else
        pNewFBuffer.Value(lngLinks(1, lngIndex)) = varVal
      End If
    Next lngIndex
    pNewFCursor.InsertFeature pNewFBuffer

    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      pNewFCursor.Flush
    End If

    Set pFeature = pFCursor.NextFeature
  Loop
  pNewFCursor.Flush

  Set CopyFeatureClassToShapefile = pNewFClass

ClearMemory:
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pFieldArray = Nothing
  Set pField = Nothing
  Set pClone = Nothing
  Set pFieldEdit = Nothing
  Set pNewField = Nothing
  Set pSourceNameColl = Nothing
  Set pNewFClass = Nothing
  Erase lngLinks
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pNewFCursor = Nothing
  Set pNewFBuffer = Nothing

End Function

Public Sub ExportFinalDataset()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray

  Dim strNewSource As String
  strNewSource = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Source_Files_March_2018\HillPlotQC_Laughlin.xlsx"

  Dim strOrigRoot As String
  Dim strModRoot As String
  Dim strShiftRoot As String
  Dim strFinalFolder As String
  Call DeclareWorkspaces(strOrigRoot, , strShiftRoot, , strModRoot, , , strFinalFolder)

  Dim strFolder As String
  Dim lngIndex As Long

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Dim pNewWSFact As IWorkspaceFactory
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pSrcWS As IFeatureWorkspace
  Dim pNewWS As IFeatureWorkspace
  Dim pSrcCoverFClass As IFeatureClass
  Dim pSrcDensFClass As IFeatureClass
  Dim pTopoOp As ITopologicalOperator4
  Dim lngQuadIndex As Long

  Dim strQuadrat As String
  Dim strDestFolder As String
  Dim strItem() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strFileHeader As String
  Dim dblCentroidX As Double
  Dim dblCentroidY As Double

  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace

  Dim strFClassName As String
  Dim strNameSplit() As String

  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pSrcWS = pNewWSFact.OpenFromFile(strShiftRoot & "\Combined_by_Site.gdb", 0)
  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strFinalFolder & "\Combined_by_Site")

  Set pWS = pSrcWS
  Set pDatasetEnum = pWS.Datasets(esriDTFeatureClass)
  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    Debug.Print strFClassName

    ExportFinalFClass pNewWS, pDataset, pMxDoc, False
    Set pDataset = pDatasetEnum.Next
  Loop

  If Not aml_func_mod.ExistFileDir(strFinalFolder & "\Shapefiles") Then
    MyGeneralOperations.CreateNestedFoldersByPath (strFinalFolder & "\Shapefiles")
  End If
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Set pNewWS = pNewWSFact.OpenFromFile(strFinalFolder & "\Shapefiles", 0)

  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    Debug.Print strFClassName

    ExportFinalFClass pNewWS, pDataset, pMxDoc, True
    Set pDataset = pDatasetEnum.Next
  Loop

  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pNewWSFact = Nothing
  Set pSrcWS = Nothing
  Set pNewWS = Nothing
  Set pSrcCoverFClass = Nothing
  Set pSrcDensFClass = Nothing
  Set pTopoOp = Nothing
  Erase strItem
  Set pDatasetEnum = Nothing
  Set pWS = Nothing
  Erase strNameSplit

End Sub


