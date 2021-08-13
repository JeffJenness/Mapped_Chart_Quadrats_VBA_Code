Attribute VB_Name = "ThisDocument_for_VM"
Option Explicit

' FIRST RUN ReviseShapefiles TO SWITCH NAMES
' NEXT RUN ConvertPointShapefiles TO CREATE NEW DATASETS
' NEXT SHIFT LOCATIONS

Public Sub ShiftFinishedShapefilesToCoordinateSystem()
  
  ' This function will copy all data to new folder, set correct coordinates, and split shapefiles by year.
  ' AREA VALUES APPEAR TO BE GETTING CALCULATED SOMEWHERE, BUT I DON'T KNOW WHERE...
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"
    
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  
  Dim strNewSource As String
  strNewSource = "E:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_March_1_2018"
  
  Dim strNewDest As String
  strNewDest = "E:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_March_1_2018_CoordsSet"
  
  Dim strFolder As String
  Dim lngIndex As Long
  
  Dim strPlotLocNames() As String
  Dim pPlotLocColl As Collection

  Dim strPlotDataNames() As String
  Dim pPlotDataColl As Collection
  
  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  
  Call ReturnQuadratVegSoilData(pPlotDataColl, strPlotDataNames)
  Call ReturnQuadratCoordsAndNames(pPlotLocColl, strPlotLocNames)
  Set pQuadratColl = FillQuadratNameColl_Rev(strQuadratNames)
  
'  For lngIndex = 0 To pFolders.Count - 1
'    Debug.Print CStr(lngIndex) & "] " & pFolders.Element(lngIndex)
'  Next lngIndex
  
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
  
'  For lngQuadIndex = 0 To 600
'    DoEvents
'    strQuadrat = "Q" & Format(lngQuadIndex, "0")
'    Debug.Print "Working on " & strQuadrat & "..."
'    strDestFolder = strNewDest & "\Shapefiles\" & strQuadrat
'    If aml_func_mod.ExistFileDir(strNewSource & "\Shapefiles\" & strQuadrat) Then
'      Set pSrcWS = pNewWSFact.OpenFromFile(strNewSource & "\Shapefiles\" & strQuadrat, 0)
'
'      If MyGeneralOperations.CheckIfFeatureClassExists(pSrcWS, strQuadrat & "_Cover") Or _
'             MyGeneralOperations.CheckIfFeatureClassExists(pSrcWS, strQuadrat & "_Density") Then
'
'        strItem = pQuadratColl.Item(Replace(strQuadrat, "Q", ""))
''        strSite = strItem(0)
''        strSiteSpecific = strItem(1)
'        strPlot = strItem(2)
''        strQuadrat = strItem(3)
''        strFileHeader = strItem(5)
'        FillQuadratCenter strPlot, pPlotLocColl, dblCentroidX, dblCentroidY
'
'        If MyGeneralOperations.CheckIfFeatureClassExists(pSrcWS, strQuadrat & "_Cover") Then
'
'          Set pSrcCoverFClass = pSrcWS.OpenFeatureClass(strQuadrat & "_Cover")
'          SplitShapefiles strDestFolder, pSrcCoverFClass, pNewWSFact, "C", strQuadrat, pMxDoc, _
'              dblCentroidX, dblCentroidY
'        End If
'        If MyGeneralOperations.CheckIfFeatureClassExists(pSrcWS, strQuadrat & "_Density") Then
'          Set pSrcDensFClass = pSrcWS.OpenFeatureClass(strQuadrat & "_Density")
'          SplitShapefiles strDestFolder, pSrcDensFClass, pNewWSFact, "D", strQuadrat, pMxDoc, _
'              dblCentroidX, dblCentroidY
'        End If
'      End If
'    End If
'  Next lngQuadIndex
  
  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pSrcWS = pNewWSFact.OpenFromFile("E:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_March_1_2018\Combined_by_Quadrat.gdb", 0)
  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase( _
      "E:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_March_1_2018_CoordsSet\Combined_by_Quadrat")
  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace
  Set pWS = pSrcWS
  Dim pCoverAll As IFeatureClass
  Dim pDensityAll As IFeatureClass
  Dim varCoverIndexes() As Variant
  Dim varDensityIndexes() As Variant
  
  Dim strFClassName As String
  Dim strNameSplit() As String
  
  Set pDatasetEnum = pWS.Datasets(esriDTFeatureClass)
  pDatasetEnum.Reset
  
  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    If Left(strFClassName, 1) = "Q" Then
      strNameSplit = Split(strFClassName, "_", , vbTextCompare)
      strQuadrat = strNameSplit(0)
      Debug.Print strFClassName
      
      strItem = pQuadratColl.Item(Replace(strQuadrat, "Q", ""))
      strPlot = strItem(2)
      FillQuadratCenter strPlot, pPlotLocColl, dblCentroidX, dblCentroidY
      ExportFGDBFClass pNewWS, pDataset, pMxDoc, dblCentroidX, dblCentroidY, pCoverAll, pDensityAll, _
          varCoverIndexes, varDensityIndexes
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
  
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, CStr(varDensityIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Orig_FID")
  
  
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

Public Sub ExportFGDBFClass(pDestWS As IFeatureWorkspace, pSrcFClass As IFeatureClass, _
    pMxDoc As IMxDocument, dblCentroidX As Double, dblCentroidY As Double, pCoverAll As IFeatureClass, _
    pDensityAll As IFeatureClass, varCoverIndexes() As Variant, varDensityIndexes() As Variant)

  
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
  strBaseString = strBaseString & "Margaret M. Moore[1] (margaret.moore@nau.edu) *" & vbNewLine
  strBaseString = strBaseString & "Helen E. Dowling[2] (ldowling@pheasantsforever.org)   " & vbNewLine
  strBaseString = strBaseString & "Robert T. Strahan[3] (strahanr@sou.edu)" & vbNewLine
  strBaseString = strBaseString & "Daniel C. Laughlin[4] (daniel.laughlin@uwyo.edu)" & vbNewLine
  strBaseString = strBaseString & "Jonathan D. Bakker[5] (jdbakker@uw.edu)" & vbNewLine
  strBaseString = strBaseString & "Judith D. Springer[6] (judy.springer@nau.edu)" & vbNewLine
  strBaseString = strBaseString & "Jeffrey S. Jenness[7] (jeffj@jennessent.com)" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine
  strBaseString = strBaseString & "1 School of Forestry, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "2 Pheasants Forever, Waterville, WA 98858 USA" & vbNewLine
  strBaseString = strBaseString & "3 Southern Oregon University, Ashland, OR 97520 USA" & vbNewLine
  strBaseString = strBaseString & "4 Department of Botany, University of Wyoming, Laramie, WY 82072 USA" & vbNewLine
  strBaseString = strBaseString & "5 School of Environmental and Forest Sciences, University of Washington, Seattle, WA  98195 USA" & vbNewLine
  strBaseString = strBaseString & "6 Ecological Restoration Institute, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "7 Jenness Enterprises, GIS Analysis and Application Design, Flagstaff, AZ  86004 USA" & vbNewLine
  strBaseString = strBaseString & "*Corresponding author: Margaret M. Moore, E-mail: Margaret.Moore@nau.edu" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine

  strAbstract = "This data set consists of 98 permanent 1-m2 quadrats located on ponderosa pine  " & _
      "–bunchgrass ecosystems in near Flagstaff, Arizona, USA.  Individual plants in " & _
      "these quadrats were identified and mapped annually from 2002-2016.  The temporal and spatial data provide " & _
      "unique opportunities to examine the effects of climate and  " & _
      "land-use variables on plant demography, population and community processes.  The original chart quadrats were " & _
      "established between 1912 and 1927 to determine the effects of livestock grazing on herbaceous plants and pine " & _
      "seedlings.  We provide the following data and data formats: (1) the digitized maps in shapefile format; (2) a " & _
      "tabular representation of centroid or point location (x, y coordinates) for species mapped as points; (3) a tabular " & _
      "representation of basal cover for species mapped as polygons; (4) a species list including synonymy of names and plant " & _
      "growth forms; (5) an inventory of the years each quadrat was sampled; and (6) tree density and basal area records " & _
      "for overstory plots that surround each quadrat." & vbCrLf & vbCrLf & strBaseString

  Dim strPurpose As String
  strPurpose = "An analysis of cover and density of southwestern ponderosa pine-bunchgrass plants mapped for fifteen years (2002-2016) in permanent quadrats."
    
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
    
    Set pPolygon = pFeature.ShapeCopy
    Call Margaret_Functions.ShiftPolygon(pPolygon, dblCentroidX, dblCentroidY)
    Set pClone = pPolygon
    
    Set pInsertFBuffer.Shape = pClone.Clone
    For lngIndex = 0 To UBound(varIndexArray, 2)
      pInsertFBuffer.Value(varIndexArray(3, lngIndex)) = pFeature.Value(varIndexArray(1, lngIndex))
    Next lngIndex
    pInsertFCursor.InsertFeature pInsertFBuffer
    
    If booDoDensity Then
      Set pDensityFBuffer.Shape = pClone.Clone
      For lngIndex = 0 To UBound(varDensityIndexes, 2)
        If varDensityIndexes(1, lngIndex) = 2 Then ' if SPCODE field, which should be integer
          If IsNull(pFeature.Value(varDensityIndexes(1, lngIndex))) Then
            pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = Null
          Else
            If Trim(CStr(pFeature.Value(varDensityIndexes(1, lngIndex)))) = "" Then
              pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = Null
            Else
              pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = pFeature.Value(varDensityIndexes(1, lngIndex))
            End If
          End If
        Else
          pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = pFeature.Value(varDensityIndexes(1, lngIndex))
        End If
'        pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = pFeature.Value(varDensityIndexes(1, lngIndex))
      Next lngIndex
      pDensityFCursor.InsertFeature pDensityFBuffer
    End If
    
    If booDoCover Then
      Set pCoverFBuffer.Shape = pClone.Clone
      For lngIndex = 0 To UBound(varCoverIndexes, 2)
        If varCoverIndexes(1, lngIndex) = 2 Then ' if SPCODE field, which should be integer
          If IsNull(pFeature.Value(varCoverIndexes(1, lngIndex))) Then
            pCoverFBuffer.Value(varCoverIndexes(3, lngIndex)) = Null
          Else
            If Trim(CStr(pFeature.Value(varCoverIndexes(1, lngIndex)))) = "" Then
              pCoverFBuffer.Value(varCoverIndexes(3, lngIndex)) = Null
            Else
              pCoverFBuffer.Value(varCoverIndexes(3, lngIndex)) = pFeature.Value(varCoverIndexes(1, lngIndex))
            End If
          End If
        Else
          pCoverFBuffer.Value(varCoverIndexes(3, lngIndex)) = pFeature.Value(varCoverIndexes(1, lngIndex))
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
  strBaseString = strBaseString & "Margaret M. Moore[1] (margaret.moore@nau.edu) *" & vbNewLine
  strBaseString = strBaseString & "Helen E. Dowling[2] (ldowling@pheasantsforever.org)   " & vbNewLine
  strBaseString = strBaseString & "Robert T. Strahan[3] (strahanr@sou.edu)" & vbNewLine
  strBaseString = strBaseString & "Daniel C. Laughlin[4] (daniel.laughlin@uwyo.edu)" & vbNewLine
  strBaseString = strBaseString & "Jonathan D. Bakker[5] (jdbakker@uw.edu)" & vbNewLine
  strBaseString = strBaseString & "Judith D. Springer[6] (judy.springer@nau.edu)" & vbNewLine
  strBaseString = strBaseString & "Jeffrey S. Jenness[7] (jeffj@jennessent.com)" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine
  strBaseString = strBaseString & "1 School of Forestry, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "2 Pheasants Forever, Waterville, WA 98858 USA" & vbNewLine
  strBaseString = strBaseString & "3 Southern Oregon University, Ashland, OR 97520 USA" & vbNewLine
  strBaseString = strBaseString & "4 Department of Botany, University of Wyoming, Laramie, WY 82072 USA" & vbNewLine
  strBaseString = strBaseString & "5 School of Environmental and Forest Sciences, University of Washington, Seattle, WA  98195 USA" & vbNewLine
  strBaseString = strBaseString & "6 Ecological Restoration Institute, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "7 Jenness Enterprises, GIS Analysis and Application Design, Flagstaff, AZ  86004 USA" & vbNewLine
  strBaseString = strBaseString & "*Corresponding author: Margaret M. Moore, E-mail: Margaret.Moore@nau.edu" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine

  strAbstract = "This data set consists of 98 permanent 1-m2 quadrats located on ponderosa pine  " & _
      "–bunchgrass ecosystems in near Flagstaff, Arizona, USA.  Individual plants in " & _
      "these quadrats were identified and mapped annually from 2002-2016.  The temporal and spatial data provide " & _
      "unique opportunities to examine the effects of climate and  " & _
      "land-use variables on plant demography, population and community processes.  The original chart quadrats were " & _
      "established between 1912 and 1927 to determine the effects of livestock grazing on herbaceous plants and pine " & _
      "seedlings.  We provide the following data and data formats: (1) the digitized maps in shapefile format; (2) a " & _
      "tabular representation of centroid or point location (x, y coordinates) for species mapped as points; (3) a tabular " & _
      "representation of basal cover for species mapped as polygons; (4) a species list including synonymy of names and plant " & _
      "growth forms; (5) an inventory of the years each quadrat was sampled; and (6) tree density and basal area records " & _
      "for overstory plots that surround each quadrat." & vbCrLf & vbCrLf & strBaseString

  Dim strPurpose As String
  strPurpose = "An analysis of cover and density of southwestern ponderosa pine-bunchgrass plants mapped for fifteen years (2002-2016) in permanent quadrats."
  
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

Public Sub ConvertPointShapefiles()
  
  ' This function will take all shapefiles in a set of nested folders, in which each shapefile represents a different year
  ' in a different quadrat, and combines all shapefiles by quadrat and saves in both shapefile and File Geodatabase format.
  
  ' AREA VALUES APPEAR TO BE GETTING CALCULATED SOMEWHERE, BUT I DON'T KNOW WHERE...
  
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
  
  Debug.Print "---------------------"
  Call FillCollections(pCoverCollection, pDensityCollection, pCoverToDensity, pDensityToCover, _
    strCoverToDensityQuery, strDensityToCoverQuery, pCoverShouldChangeColl, pDensityShouldChangeColl)

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  strRoot = "E:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data - March_1_2018b"
  
  Dim strNewRoot As String
  strNewRoot = "E:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_March_1_2018"
  
  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long
  
'  For lngIndex = 0 To pFolders.Count - 1
'    Debug.Print CStr(lngIndex) & "] " & pFolders.Element(lngIndex)
'  Next lngIndex
  
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
  
  ' ADDED APRIL 21 TO CREATE NEW FEATURE CLASSES
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
  Dim pField As IField
  Dim pNewFields As esriSystem.IVariantArray
  Dim lngIndex2 As Long
  
  Dim pNewDensityFClass As IFeatureClass
  Dim varDensityFieldIndexArray() As Variant
  Dim strNewDensityFClassName As String
  Dim booDensityHasFields As Boolean
  Dim lngDensityFClassIndex As Long
  Dim lngDensityQuadratIndex As Long
  Dim lngDensityYearIndex As Long
  Dim lngDensityTypeIndex As Long
  Dim lngDensityOrigFIDIndex As Long
  
  Dim pNewGDBDensityFClass As IFeatureClass
  Dim varGDBDensityFieldIndexArray() As Variant
  Dim strGDBNewDensityFClassName As String
  Dim booGDBDensityHasFields As Boolean
  Dim lngGDBDensityFClassIndex As Long
  Dim lngGDBDensityQuadratIndex As Long
  Dim lngGDBDensityYearIndex As Long
  Dim lngGDBDensityTypeIndex As Long
  Dim lngGDBDensityOrigFIDIndex As Long
      
  Dim pNewCoverFClass As IFeatureClass
  Dim varCoverFieldIndexArray() As Variant
  Dim strNewCoverFClassName As String
  Dim booCoverHasFields As Boolean
  Dim lngCoverFClassIndex As Long
  Dim lngCoverQuadratIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverTypeIndex As Long
  Dim lngCoverOrigFIDIndex As Long
      
  Dim pNewGDBCoverFClass As IFeatureClass
  Dim varGDBCoverFieldIndexArray() As Variant
  Dim strGDBNewCoverFClassName As String
  Dim booGDBCoverHasFields As Boolean
  Dim lngGDBCoverFClassIndex As Long
  Dim lngGDBCoverQuadratIndex As Long
  Dim lngGDBCoverYearIndex As Long
  Dim lngGDBCoverTypeIndex As Long
  Dim lngGDBCoverOrigFIDIndex As Long
  
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
  Dim lngGDBFClassIndex As Long
  Dim lngGDBQuadratIndex As Long
  Dim lngGDBYearIndex As Long
  Dim lngGDBTypeIndex As Long
  Dim lngGDBOrigFIDIndex As Long
  Dim lngGDBIsEmptyIndex As Long
  
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
        
  Dim pNewCombinedCoverFClass As IFeatureClass
  Dim varCombinedCoverFieldIndexArray() As Variant
  Dim strNewCombinedCoverFClassName As String
  Dim booCombinedCoverHasFields As Boolean
  Dim lngCombinedCoverFClassIndex As Long
  Dim lngCombinedCoverQuadratIndex As Long
  Dim lngCombinedCoverYearIndex As Long
  Dim lngCombinedCoverTypeIndex As Long
  Dim lngCombinedCoverOrigFIDIndex As Long
  
  Dim pCombinedDestFClass As IFeatureClass
  Dim varCombinedIndexArray() As Variant
  Dim lngCombinedFClassIndex As Long
  Dim lngCombinedQuadratIndex As Long
  Dim lngCombinedYearIndex As Long
  Dim lngCombinedTypeIndex As Long
  Dim lngCombinedOrigFIDIndex As Long
  Dim lngCombinedIsEmptyIndex As Long
        
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
  pSRTol.XYTolerance = 0.0001
  
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pTopoOp As ITopologicalOperator4

  ' MODIFIED DEC. 9 2017 TO DELETE EMPTY DATASETS
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
  
  Dim pAltDestGDBFClass As IFeatureClass
  Dim varAltGDBIndexArray() As Variant
  Dim lngAltGDBFClassIndex As Long
  Dim lngAltGDBQuadratIndex As Long
  Dim lngAltGDBYearIndex As Long
  Dim lngAltGDBTypeIndex As Long
  Dim lngAltGDBIsEmptyIndex As Long
  
  Dim pAltCombinedDestFClass As IFeatureClass
  Dim varAltCombinedIndexArray() As Variant
  Dim lngAltCombinedFClassIndex As Long
  Dim lngAltCombinedQuadratIndex As Long
  Dim lngAltCombinedYearIndex As Long
  Dim lngAltCombinedTypeIndex As Long
  Dim lngAltCombinedIsEmptyIndex As Long
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
  strBaseString = strBaseString & "Margaret M. Moore[1] (margaret.moore@nau.edu) *" & vbNewLine
  strBaseString = strBaseString & "Helen E. Dowling[2] (ldowling@pheasantsforever.org)   " & vbNewLine
  strBaseString = strBaseString & "Robert T. Strahan[3] (strahanr@sou.edu)" & vbNewLine
  strBaseString = strBaseString & "Daniel C. Laughlin[4] (daniel.laughlin@uwyo.edu)" & vbNewLine
  strBaseString = strBaseString & "Jonathan D. Bakker[5] (jdbakker@uw.edu)" & vbNewLine
  strBaseString = strBaseString & "Judith D. Springer[6] (judy.springer@nau.edu)" & vbNewLine
  strBaseString = strBaseString & "Jeffrey S. Jenness[7] (jeffj@jennessent.com)" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine
  strBaseString = strBaseString & "1 School of Forestry, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "2 Pheasants Forever, Waterville, WA 98858 USA" & vbNewLine
  strBaseString = strBaseString & "3 Southern Oregon University, Ashland, OR 97520 USA" & vbNewLine
  strBaseString = strBaseString & "4 Department of Botany, University of Wyoming, Laramie, WY 82072 USA" & vbNewLine
  strBaseString = strBaseString & "5 School of Environmental and Forest Sciences, University of Washington, Seattle, WA  98195 USA" & vbNewLine
  strBaseString = strBaseString & "6 Ecological Restoration Institute, Northern Arizona University, Flagstaff, AZ 86011 USA" & vbNewLine
  strBaseString = strBaseString & "7 Jenness Enterprises, GIS Analysis and Application Design, Flagstaff, AZ  86004 USA" & vbNewLine
  strBaseString = strBaseString & "*Corresponding author: Margaret M. Moore, E-mail: Margaret.Moore@nau.edu" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine

  strAbstract = "This data set consists of 98 permanent 1-m2 quadrats located on ponderosa pine  " & _
      "–bunchgrass ecosystems in near Flagstaff, Arizona, USA.  Individual plants in " & _
      "these quadrats were identified and mapped annually from 2002-2016.  The temporal and spatial data provide " & _
      "unique opportunities to examine the effects of climate and  " & _
      "land-use variables on plant demography, population and community processes.  The original chart quadrats were " & _
      "established between 1912 and 1927 to determine the effects of livestock grazing on herbaceous plants and pine " & _
      "seedlings.  We provide the following data and data formats: (1) the digitized maps in shapefile format; (2) a " & _
      "tabular representation of centroid or point location (x, y coordinates) for species mapped as points; (3) a tabular " & _
      "representation of basal cover for species mapped as polygons; (4) a species list including synonymy of names and plant " & _
      "growth forms; (5) an inventory of the years each quadrat was sampled; and (6) tree density and basal area records " & _
      "for overstory plots that surround each quadrat." & vbCrLf & vbCrLf & strBaseString

  Dim strPurpose As String
  strPurpose = "An analysis of cover and density of southwestern ponderosa pine-bunchgrass plants mapped for fifteen years (2002-2016) in permanent quadrats."
  
  ' REMEMBER TO REMOVE INITIAL SPACES
  ' REMEMBER TO CHANGE GRAMMINOID TO GRAMINOID
  ' REMEMBER TO REMOVE LINE RETURNS
  
  Set pNewFGDBWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strNewRoot & "\Combined_by_Quadrat")
  Set pNewFeatFGDBWS = pNewFGDBWS
  
  For lngIndex = 0 To pFolders.Count - 1
    DoEvents
    strFolder = pFolders.Element(lngIndex)
    
'    strFolder = "E:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data\Q80\"
    
    varDatasets = ReturnFeatureClassesOrNothing(strFolder, booFoundShapefiles, booFoundPolys, booFoundPoints, _
        pRepPointFClass, pRepPolyFClass)
    
    ' MODIFIED DEC. 9 2017 TO ALLOW FOR FEATURES TO BE CONVERTED TO OTHER TYPE.
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
        
        ' FOR SHAPEFILE
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
        
        ' FOR FILE GEODATABASE
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
        
        ' FOR COMBINED
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
        
        ' FOR SHAPEFILE
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
        
        ' FOR FILE GEODATABASE
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
        
        
        ' FOR COMBINED
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
      
      ' NEW METHOD:  AT THIS POINT WE HAVE BOTH COVER AND DENSITY FEATURE CLASSES FOR THIS QUADRAT.
      ' THE NEW FEATURE CLASS WILL BE EITHER COVER OR DENSITY.
      ' WE NEED TO DECIDE WHICH QUADRAT FEATURE CLASS TO WRITE EACH FEATURE TO, BASED
      ' ON WHETHER THIS NEW FEATURE CLASS IS COVER OR DENSITY, AND WHETHER WE ARE SUPPOSED TO CONVERT IT.
      
      For lngDatasetIndex = 0 To UBound(varDatasets)
        DoEvents
        
        Set pDataset = varDatasets(lngDatasetIndex)
        Set pFClass = pDataset
        lngSpeciesIndex = pFClass.FindField("Species")
        If lngSpeciesIndex = -1 Then
          DoEvents
        End If
        strSplit = Split(pDataset.BrowseName, "_")
        
        Debug.Print "  --> Adding Dataset '" & pDataset.BrowseName & "'"
'        strBase = """" & pDataset.BrowseName & """" & vbTab & """" & strSplit(0) & """" & vbTab & _
'            """" & strSplit(1) & """" & vbTab & """" & IIf(strSplit(2) = "C", "Cover", "Density") & """"
        
        strQuadrat = strSplit(0)
        strYear = strSplit(1)
        strFClassName = pDataset.BrowseName
        
        If strFClassName = "Q10_2006_D" Then
          DoEvents
        End If
        
        ' pCoverShouldChangeColl , pDensityShouldChangeColl
        
        If strSplit(2) = "C" Then
          strType = "Cover"
          Set pDestFClass = pNewCoverFClass
          varIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestFClass)
          lngFClassIndex = lngCoverFClassIndex
          lngQuadratIndex = lngCoverQuadratIndex
          lngYearIndex = lngCoverYearIndex
          lngTypeIndex = lngCoverTypeIndex
          lngIsEmptyIndex = pDestFClass.FindField("IsEmpty")
          
          Set pDestGDBFClass = pNewGDBCoverFClass
          varGDBIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestGDBFClass)
          lngGDBFClassIndex = lngGDBCoverFClassIndex
          lngGDBQuadratIndex = lngGDBCoverQuadratIndex
          lngGDBYearIndex = lngGDBCoverYearIndex
          lngGDBTypeIndex = lngGDBCoverTypeIndex
          lngGDBIsEmptyIndex = pDestGDBFClass.FindField("IsEmpty")
          
          Set pCombinedDestFClass = pNewCombinedCoverFClass
          varCombinedIndexArray = ReturnArrayOfFieldLinks(pFClass, pCombinedDestFClass)
          lngCombinedFClassIndex = lngCombinedCoverFClassIndex
          lngCombinedQuadratIndex = lngCombinedCoverQuadratIndex
          lngCombinedYearIndex = lngCombinedCoverYearIndex
          lngCombinedTypeIndex = lngCombinedCoverTypeIndex
          lngCombinedIsEmptyIndex = pCombinedDestFClass.FindField("IsEmpty")
          Set pCombinedFCursor = pCombinedCoverFCursor
          Set pCombinedFBuffer = pCombinedCoverFBuffer
          
          ' ALTERNATE, IF SUPPOSED TO SWITCH
          strAltType = "Density"
          Set pAltDestFClass = pNewDensityFClass
          ReDim varAltIndexArray(3, 4)
          
          varAltIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestFClass)
          lngAltFClassIndex = lngDensityFClassIndex
          lngAltQuadratIndex = lngDensityQuadratIndex
          lngAltYearIndex = lngDensityYearIndex
          lngAltTypeIndex = lngDensityTypeIndex
          lngAltIsEmptyIndex = pAltDestFClass.FindField("IsEmpty")
          
          Set pAltDestGDBFClass = pNewGDBDensityFClass
          varAltGDBIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestGDBFClass)
          lngAltGDBFClassIndex = lngGDBDensityFClassIndex
          lngAltGDBQuadratIndex = lngGDBDensityQuadratIndex
          lngAltGDBYearIndex = lngGDBDensityYearIndex
          lngAltGDBTypeIndex = lngGDBDensityTypeIndex
          lngAltGDBIsEmptyIndex = pAltDestGDBFClass.FindField("IsEmpty")
          
          Set pAltCombinedDestFClass = pNewCombinedDensityFClass
          varAltCombinedIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltCombinedDestFClass)
          lngAltCombinedFClassIndex = lngCombinedDensityFClassIndex
          lngAltCombinedQuadratIndex = lngCombinedDensityQuadratIndex
          lngAltCombinedYearIndex = lngCombinedDensityYearIndex
          lngAltCombinedTypeIndex = lngCombinedDensityTypeIndex
          lngAltCombinedIsEmptyIndex = pAltCombinedDestFClass.FindField("IsEmpty")
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
          
          Set pDestGDBFClass = pNewGDBDensityFClass
          varGDBIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestGDBFClass)
          lngGDBFClassIndex = lngGDBDensityFClassIndex
          lngGDBQuadratIndex = lngGDBDensityQuadratIndex
          lngGDBYearIndex = lngGDBDensityYearIndex
          lngGDBTypeIndex = lngGDBDensityTypeIndex
          lngGDBIsEmptyIndex = pDestGDBFClass.FindField("IsEmpty")
          
          Set pCombinedDestFClass = pNewCombinedDensityFClass
          varCombinedIndexArray = ReturnArrayOfFieldLinks(pFClass, pCombinedDestFClass)
          lngCombinedFClassIndex = lngCombinedDensityFClassIndex
          lngCombinedQuadratIndex = lngCombinedDensityQuadratIndex
          lngCombinedYearIndex = lngCombinedDensityYearIndex
          lngCombinedTypeIndex = lngCombinedDensityTypeIndex
          lngCombinedIsEmptyIndex = pCombinedDestFClass.FindField("IsEmpty")
          Set pCombinedFCursor = pCombinedDensityFCursor
          Set pCombinedFBuffer = pCombinedDensityFBuffer
          
          ' ALTERNATE, IF SUPPOSED TO SWITCH
          strAltType = "Cover"
          Set pAltDestFClass = pNewCoverFClass
          varAltIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestFClass)
          lngAltFClassIndex = lngCoverFClassIndex
          lngAltQuadratIndex = lngCoverQuadratIndex
          lngAltYearIndex = lngCoverYearIndex
          lngAltTypeIndex = lngCoverTypeIndex
          lngAltIsEmptyIndex = pAltDestFClass.FindField("IsEmpty")
          
          Set pAltDestGDBFClass = pNewGDBCoverFClass
          varAltGDBIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestGDBFClass)
          lngAltGDBFClassIndex = lngGDBCoverFClassIndex
          lngAltGDBQuadratIndex = lngGDBCoverQuadratIndex
          lngAltGDBYearIndex = lngGDBCoverYearIndex
          lngAltGDBTypeIndex = lngGDBCoverTypeIndex
          lngAltGDBIsEmptyIndex = pAltDestGDBFClass.FindField("IsEmpty")
          
          Set pAltCombinedDestFClass = pNewCombinedCoverFClass
          varAltCombinedIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltCombinedDestFClass)
          lngAltCombinedFClassIndex = lngCombinedCoverFClassIndex
          lngAltCombinedQuadratIndex = lngCombinedCoverQuadratIndex
          lngAltCombinedYearIndex = lngCombinedCoverYearIndex
          lngAltCombinedTypeIndex = lngCombinedCoverTypeIndex
          lngAltCombinedIsEmptyIndex = pAltCombinedDestFClass.FindField("IsEmpty")
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
            
'            strSpecies = CheckSpeciesAgainstSpecialConversions(varSpecialConversions, strQuadrat, CLng(strYear), _
                strSpecies, strNoteOnChanges)
            
'            If strSpecies = "Muhlenbergia tricholepis" Then
'              DoEvents
'            End If
            If Trim(strSpecies) = "" And strType = "Density" Then strSpecies = "No Point Species"
            If Trim(strSpecies) = "" And strType = "Cover" Then strSpecies = "No Polygon Species"

            strHexSpecies = HexifyName(strSpecies)
            If strType = "Density" Then
              booShouldChange = pDensityShouldChangeColl.Item(strHexSpecies)
            Else
              booShouldChange = pCoverShouldChangeColl.Item(strHexSpecies)
            End If
            
'            If booShouldChange Then
'              Debug.Print "Changing for " & strType & " species '" & strSpecies & "'..."
'              DoEvents
'            Else
'              Debug.Print "Not changing for " & strType & " species '" & strSpecies & "'..."
'            End If
            
            If strType = "Density" Then
              Set pPoint = pSrcFeature.ShapeCopy
              If pPoint.IsEmpty Then
                Set pPolygon = New Polygon
              Else
'                Set pPolygon = MyGeometricOperations.CreateCircleAroundPoint(pPoint, 0.001, 30)
                Set pPolygon = ReturnCircleClippedToQuadrat(pPoint, 0.001, 30, pQuadrat)
              End If
            Else
              Set pPolygon = pSrcFeature.ShapeCopy
            End If
            
            strIsEmpty = CBool(pPolygon.IsEmpty)
            Set pPolygon.SpatialReference = pSpRef
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
'                    Set pTestPolygon = MyGeometricOperations.CreateCircleAroundPoint(pPoint, 0.001, 30)
                    Set pTestPolygon = ReturnCircleClippedToQuadrat(pTestPoint, 0.001, 30, pQuadrat, pPolygon)
                    Set pTestPolygon.SpatialReference = pSpRef
                    Set varPoints(lngConvertIndex) = pTestPolygon
                    If Not pTestPolygon.IsEmpty Then
'                      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pTestPolygon, "Delete_Me"
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
                  pAltDestFBuffer.Value(lngAltTypeIndex) = strType
                  pAltDestFBuffer.Value(lngAltIsEmptyIndex) = strIsEmpty
                  pAltDestFCursor.InsertFeature pAltDestFBuffer
                  
                  Set pAltDestGDBFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltGDBIndexArray, 2)
                    pAltDestGDBFBuffer.Value(varAltGDBIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltGDBIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltDestGDBFBuffer.Value(lngAltGDBFClassIndex) = strFClassName
                  pAltDestGDBFBuffer.Value(lngAltGDBQuadratIndex) = strQuadrat
                  pAltDestGDBFBuffer.Value(lngAltGDBYearIndex) = strYear
                  pAltDestGDBFBuffer.Value(lngAltGDBTypeIndex) = strType
                  pAltDestGDBFBuffer.Value(lngAltGDBIsEmptyIndex) = strIsEmpty
                  pAltDestGDBFCursor.InsertFeature pAltDestGDBFBuffer
                  
                  Set pAltCombinedFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltCombinedIndexArray, 2)
                    pAltCombinedFBuffer.Value(varAltCombinedIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltCombinedIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltCombinedFBuffer.Value(lngAltCombinedFClassIndex) = strFClassName
                  pAltCombinedFBuffer.Value(lngAltCombinedQuadratIndex) = strQuadrat
                  pAltCombinedFBuffer.Value(lngAltCombinedYearIndex) = strYear
                  pAltCombinedFBuffer.Value(lngAltCombinedTypeIndex) = strType
                  pAltCombinedFBuffer.Value(lngAltCombinedIsEmptyIndex) = strIsEmpty
                  pAltCombinedFCursor.InsertFeature pAltCombinedFBuffer
                End If
              Next lngConvertIndex

            Else
              Set pDestFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varIndexArray, 2)
  '              Debug.Print "Copying '" & CStr(pSrcFeature.Value(varIndexArray(1, lngIndex2))) & _
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                  
                pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = pSrcFeature.Value(varIndexArray(1, lngIndex2))
              Next lngIndex2
              pDestFBuffer.Value(lngFClassIndex) = strFClassName
              pDestFBuffer.Value(lngQuadratIndex) = strQuadrat
              pDestFBuffer.Value(lngYearIndex) = strYear
              pDestFBuffer.Value(lngTypeIndex) = strType
              pDestFBuffer.Value(lngIsEmptyIndex) = strIsEmpty
              pDestFCursor.InsertFeature pDestFBuffer
              
              Set pDestGDBFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varGDBIndexArray, 2)
  '              Debug.Print "Copying '" & CStr(pSrcFeature.Value(varIndexArray(1, lngIndex2))) & _
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                pDestGDBFBuffer.Value(varGDBIndexArray(3, lngIndex2)) = pSrcFeature.Value(varGDBIndexArray(1, lngIndex2))
              Next lngIndex2
              pDestGDBFBuffer.Value(lngGDBFClassIndex) = strFClassName
              pDestGDBFBuffer.Value(lngGDBQuadratIndex) = strQuadrat
              pDestGDBFBuffer.Value(lngGDBYearIndex) = strYear
              pDestGDBFBuffer.Value(lngGDBTypeIndex) = strType
              pDestGDBFBuffer.Value(lngGDBIsEmptyIndex) = strIsEmpty
              pDestGDBFCursor.InsertFeature pDestGDBFBuffer
              
              
              Set pCombinedFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varCombinedIndexArray, 2)
  '              Debug.Print "Copying '" & CStr(pSrcFeature.Value(varIndexArray(1, lngIndex2))) & _
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                  
                pCombinedFBuffer.Value(varCombinedIndexArray(3, lngIndex2)) = pSrcFeature.Value(varCombinedIndexArray(1, lngIndex2))
              Next lngIndex2
              pCombinedFBuffer.Value(lngCombinedFClassIndex) = strFClassName
              pCombinedFBuffer.Value(lngCombinedQuadratIndex) = strQuadrat
              pCombinedFBuffer.Value(lngCombinedYearIndex) = strYear
              pCombinedFBuffer.Value(lngCombinedTypeIndex) = strType
              pCombinedFBuffer.Value(lngCombinedIsEmptyIndex) = strIsEmpty
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
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Orig_FID")
        
      ' MODIFIED DEC. 9 2017 TO DELETE EMPTY DATASETS
      Set pTempFClass = pNewDensityFClass
      If pTempFClass.FeatureCount(Nothing) = 0 Then
        Set pTempDataset = pNewDensityFClass
        Set pNewDensityFClass = Nothing
        pTempDataset.DELETE
      Else ' MAKE METADATA
      End If
      Set pTempFClass = pNewGDBDensityFClass
      If pTempFClass.FeatureCount(Nothing) = 0 Then
        Set pTempDataset = pNewGDBDensityFClass
        Set pNewGDBDensityFClass = Nothing
        pTempDataset.DELETE
      Else ' MAKE METADATA
      End If
      Set pTempFClass = pNewCoverFClass
      If pTempFClass.FeatureCount(Nothing) = 0 Then
        Set pTempDataset = pNewCoverFClass
        Set pNewCoverFClass = Nothing
        pTempDataset.DELETE
      Else ' MAKE METADATA
      End If
      Set pTempFClass = pNewGDBCoverFClass
      If pTempFClass.FeatureCount(Nothing) = 0 Then
        Set pTempDataset = pNewGDBCoverFClass
        Set pNewGDBCoverFClass = Nothing
        pTempDataset.DELETE
      Else ' MAKE METADATA
      End If
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
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Orig_FID")
  
  ' MAKE METADATA
  
  ' varFieldIndexArray WILL HAVE 4 COLUMNS AND ANY NUMBER OR ROWS.
  ' COLUMN 0 = SOURCE FIELD NAME
  ' COLUMN 1 = SOURCE FIELD INDEX
  ' COLUMN 2 = NEW FIELD NAME
  ' COLUMN 3 = NEW FIELD INDEX
  
  strReport = strReport & vbCrLf & "Done..." & vbCrLf & _
    MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
       
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.Clear
'  pDataObj.SetText strReport & vbCrLf & "-----------------------------------" & vbCrLf & strExcelFullReport
'  pDataObj.PutInClipboard
  
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
ClearMemory:
  Set pCoverCollection = Nothing
  Set pDensityCollection = Nothing
  Set pCoverToDensity = Nothing
  Set pDensityToCover = Nothing
  Set pCoverShouldChangeColl = Nothing
  Set pDensityShouldChangeColl = Nothing
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




End Sub



Public Sub CheckExtentYearly()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim pFeatDef As IFeatureLayerDefinition2
  Dim lngYearIndex As Long
  Dim lngQuadIndex As Long
  Dim strPrefix As String
  Dim strSuffix As String
  Dim pQueryFilt As IQueryFilter
  Dim pYearColl As New Collection
  Dim pQuadColl As New Collection
  Dim varVal As Variant
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  
  Dim pLabelArray As esriSystem.IArray
  Set pLabelArray = MyGeneralOperations.ReturnGraphicsByName(pMxDoc, "QuadName", True)
  Dim pLabel As ITextElement
  Set pLabel = pLabelArray.Element(0)
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Cover_All", pMxDoc.FocusMap)
  Set pFClass = pFLayer.FeatureClass
  Set pFeatDef = pFLayer
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix)
  lngYearIndex = pFClass.FindField("Year")
  lngQuadIndex = pFClass.FindField("Quadrat")
  
  Dim varQuads As Variant
  varQuads = Array("Q1", "Q10", "Q106", "Q11", "Q12", "Q13", "Q14", "Q15", "Q16", "Q17", "Q18", "Q19", "Q2", "Q20", "Q21", "Q22", "Q23", "Q24", "Q25", "Q26", "Q27", "Q28", "Q29", "Q3", "Q30", "Q31", "Q32", "Q33", "Q34", "Q35", "Q36", "Q37", "Q38", "Q39", "Q4", "Q40", "Q41", "Q42", "Q43", "Q44", "Q45", "Q46", "Q47", "Q48", "Q49", "Q494", "Q498", "Q5", "Q50", "Q51", "Q52", "Q53", "Q54", "Q55", "Q56", "Q57", "Q58", "Q59", "Q6", "Q60", "Q61", "Q62", "Q63", "Q64", "Q65", "Q66", "Q67", "Q68", "Q69", "Q7", "Q70", "Q71", "Q72", "Q73", "Q74", "Q75", "Q76", "Q77", "Q78", "Q79", "Q8", "Q80", "Q81", "Q82", "Q83", "Q84", "Q85", "Q86", "Q87", "Q88", "Q89", "Q9", "Q90", "Q91", "Q92", "Q93", "Q94", "Q95", "Q96", "Q97", "Q98")
  Dim varYears As Variant
  varYears = Array("2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016")
  
  Dim pEnv As IEnvelope
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim strYears() As String
  Dim strQuads() As String
  Dim lngYearCounter As Long
  Dim lngQuadCounter As Long
  Dim strYear As String
  Dim strQuad As String
  Set pQueryFilt = New QueryFilter
  Dim strQueryString As String
  Dim strCheck() As String
  Dim lngCheckCounter As Long
  lngCheckCounter = -1
  Dim pGeoDataset As IGeoDataset
  
  Set pGeoDataset = pFClass
  Set pEnv = New Envelope
  Set pEnv.SpatialReference = pGeoDataset.SpatialReference
    
'  For lngIndex1 = 0 To UBound(varQuads)
'    strQuad = CStr(varQuads(lngIndex1))
'    Debug.Print "Quad = " & strQuad
'
'    For lngIndex2 = 0 To UBound(varYears)
'      strYear = CStr(varYears(lngIndex2))
'
'      Set pEnv = New Envelope
'
'      strQueryString = strPrefix & "Quadrat" & strSuffix & " = '" & strQuad & "' AND " & _
'          strPrefix & "Year" & strSuffix & " = '" & strYear & "'"
'      pFeatDef.DefinitionExpression = strQueryString
'      pQueryFilt.WhereClause = strQueryString
'
'      If pFClass.FeatureCount(pQueryFilt) > 0 Then
'
'        Set pFCursor = pFClass.Search(pQueryFilt, False)
'        Set pFeature = pFCursor.NextFeature
'        pEnv.PutCoords 0, 0, 1, 1
'
'        Do Until pFeature Is Nothing
'          If Not pFeature.ShapeCopy.IsEmpty Then
'            pEnv.Union pFeature.ShapeCopy.Envelope
'          End If
'
'          Set pFeature = pFCursor.NextFeature
'        Loop
'
'        If pEnv.XMin < -0.1 Or pEnv.YMin < -0.1 Or pEnv.XMax > 1.1 Or pEnv.YMax > 1.1 Then
'          lngCheckCounter = lngCheckCounter + 1
'          ReDim Preserve strCheck(1, lngCheckCounter)
'          strCheck(0, lngCheckCounter) = strQuad
'          strCheck(1, lngCheckCounter) = strYear
'          Debug.Print CStr(lngCheckCounter) & "] " & "Quad " & strQuad & ", Year " & strYear
'        End If
'
''        pLabel.Text = "Quad " & strQuad & ", Year " & strYear
''        pMxDoc.ActiveView.Refresh
''
''        Debug.Print pLabel.Text
'
'        DoEvents
'
'        DoEvents
'      End If
'    Next lngIndex2
'  Next lngIndex1
  
  ReDim strCheck(1, 22)
  strCheck(0, 0) = "Q1"
  strCheck(1, 0) = "2009"
  strCheck(0, 1) = "Q1"
  strCheck(1, 1) = "2012"
  strCheck(0, 2) = "Q1"
  strCheck(1, 2) = "2014"
  strCheck(0, 3) = "Q1"
  strCheck(1, 3) = "2015"
  strCheck(0, 4) = "Q1"
  strCheck(1, 4) = "2016"
  strCheck(0, 5) = "Q5"
  strCheck(1, 5) = "2014"
  strCheck(0, 6) = "Q5"
  strCheck(1, 6) = "2015"
  strCheck(0, 7) = "Q11"
  strCheck(1, 7) = "2007"
  strCheck(0, 8) = "Q11"
  strCheck(1, 8) = "2009"
  strCheck(0, 9) = "Q11"
  strCheck(1, 9) = "2010"
  strCheck(0, 10) = "Q11"
  strCheck(1, 10) = "2011"
  strCheck(0, 11) = "Q14"
  strCheck(1, 11) = "2007"
  strCheck(0, 12) = "Q16"
  strCheck(1, 12) = "2015"
  strCheck(0, 13) = "Q19"
  strCheck(1, 13) = "2012"
  strCheck(0, 14) = "Q19"
  strCheck(1, 14) = "2015"
  strCheck(0, 15) = "Q19"
  strCheck(1, 15) = "2016"
  strCheck(0, 16) = "Q29"
  strCheck(1, 16) = "2002"
  strCheck(0, 17) = "Q36"
  strCheck(1, 17) = "2012"
  strCheck(0, 18) = "Q36"
  strCheck(1, 18) = "2014"
  strCheck(0, 19) = "Q36"
  strCheck(1, 19) = "2015"
  strCheck(0, 20) = "Q57"
  strCheck(1, 20) = "2012"
  strCheck(0, 21) = "Q57"
  strCheck(1, 21) = "2015"
  strCheck(0, 22) = "Q84"
  strCheck(1, 22) = "2012"
  
  Dim pActiveView As IActiveView
  Set pActiveView = pMxDoc.FocusMap
  Dim pScreenDisp As IScreenDisplay
  Set pScreenDisp = pActiveView.ScreenDisplay
  Dim strName As String
  
  For lngIndex1 = 0 To UBound(strCheck, 2)
    strQuad = strCheck(0, lngIndex1)
    strYear = strCheck(1, lngIndex1)
    Debug.Print "Quad = " & strQuad
    
    Set pEnv = New Envelope

    strQueryString = strPrefix & "Quadrat" & strSuffix & " = '" & strQuad & "' AND " & _
        strPrefix & "Year" & strSuffix & " = '" & strYear & "'"
    pFeatDef.DefinitionExpression = strQueryString
    pQueryFilt.WhereClause = strQueryString

    Set pFCursor = pFClass.Search(pQueryFilt, False)
    Set pFeature = pFCursor.NextFeature
    pEnv.PutCoords 0, 0, 1, 1

    Do Until pFeature Is Nothing
      If Not pFeature.ShapeCopy.IsEmpty Then
        pEnv.Union pFeature.ShapeCopy.Envelope
      End If

      Set pFeature = pFCursor.NextFeature
    Loop
    
    pEnv.Expand 1.1, 1.1, True
    pActiveView.Extent = pEnv
    pLabel.Text = "Quad " & strQuad & ", Year " & strYear
    pMxDoc.ActiveView.Refresh
    
    strName = MyGeneralOperations.MakeUniqueShapeFilename( _
        "E:\arcGIS_stuff\consultation\Margaret_Moore\Odd_Data\" & strQuad & "_" & strYear & ".png")
    Map_Module.ExportActiveView strName
    
    Debug.Print pLabel.Text

    DoEvents

    DoEvents
  Next lngIndex1
  
  Debug.Print "Done..."
      
ClearMemory:
  Set pMxDoc = Nothing
  Set pFLayer = Nothing
  Set pFClass = Nothing
  Set pFeatDef = Nothing
  Set pQueryFilt = Nothing
  Set pYearColl = Nothing
  Set pQuadColl = Nothing
  varVal = Null
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pLabelArray = Nothing
  Set pLabel = Nothing
  varQuads = Null
  varYears = Null
  Erase strYears
  Erase strQuads


End Sub


Public Sub LookYearly()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim pFeatDef As IFeatureLayerDefinition2
  Dim lngYearIndex As Long
  Dim lngQuadIndex As Long
  Dim strPrefix As String
  Dim strSuffix As String
  Dim pQueryFilt As IQueryFilter
  Dim pYearColl As New Collection
  Dim pQuadColl As New Collection
  Dim varVal As Variant
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  
  Dim pLabelArray As esriSystem.IArray
  Set pLabelArray = MyGeneralOperations.ReturnGraphicsByName(pMxDoc, "QuadName", True)
  Dim pLabel As ITextElement
  Set pLabel = pLabelArray.Element(0)
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Cover_All", pMxDoc.FocusMap)
  Set pFClass = pFLayer.FeatureClass
  Set pFeatDef = pFLayer
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix)
  lngYearIndex = pFClass.FindField("Year")
  lngQuadIndex = pFClass.FindField("Quadrat")
  
  Dim varQuads As Variant
  varQuads = Array("Q1", "Q10", "Q106", "Q11", "Q12", "Q13", "Q14", "Q15", "Q16", "Q17", "Q18", "Q19", "Q2", "Q20", "Q21", "Q22", "Q23", "Q24", "Q25", "Q26", "Q27", "Q28", "Q29", "Q3", "Q30", "Q31", "Q32", "Q33", "Q34", "Q35", "Q36", "Q37", "Q38", "Q39", "Q4", "Q40", "Q41", "Q42", "Q43", "Q44", "Q45", "Q46", "Q47", "Q48", "Q49", "Q494", "Q498", "Q5", "Q50", "Q51", "Q52", "Q53", "Q54", "Q55", "Q56", "Q57", "Q58", "Q59", "Q6", "Q60", "Q61", "Q62", "Q63", "Q64", "Q65", "Q66", "Q67", "Q68", "Q69", "Q7", "Q70", "Q71", "Q72", "Q73", "Q74", "Q75", "Q76", "Q77", "Q78", "Q79", "Q8", "Q80", "Q81", "Q82", "Q83", "Q84", "Q85", "Q86", "Q87", "Q88", "Q89", "Q9", "Q90", "Q91", "Q92", "Q93", "Q94", "Q95", "Q96", "Q97", "Q98")
  Dim varYears As Variant
  varYears = Array("2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016")
  
  
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim strYears() As String
  Dim strQuads() As String
  Dim lngYearCounter As Long
  Dim lngQuadCounter As Long
  Dim strYear As String
  Dim strQuad As String
  Set pQueryFilt = New QueryFilter
  Dim strQueryString As String
  
  For lngIndex1 = 0 To UBound(varQuads)
    strQuad = CStr(varQuads(lngIndex1))
    
    For lngIndex2 = 0 To UBound(varYears)
      strYear = CStr(varYears(lngIndex2))
      
      strQueryString = strPrefix & "Quadrat" & strSuffix & " = '" & strQuad & "' AND " & _
          strPrefix & "Year" & strSuffix & " = '" & strYear & "'"
      pFeatDef.DefinitionExpression = strQueryString
      pQueryFilt.WhereClause = strQueryString
      
      If pFClass.FeatureCount(pQueryFilt) > 0 Then
        pLabel.Text = "Quad " & strQuad & ", Year " & strYear
        pMxDoc.ActiveView.Refresh
      
        Debug.Print pLabel.Text
        
        DoEvents
        
        DoEvents
      End If
    Next lngIndex2
  Next lngIndex1
  
      
ClearMemory:
  Set pMxDoc = Nothing
  Set pFLayer = Nothing
  Set pFClass = Nothing
  Set pFeatDef = Nothing
  Set pQueryFilt = Nothing
  Set pYearColl = Nothing
  Set pQuadColl = Nothing
  varVal = Null
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pLabelArray = Nothing
  Set pLabel = Nothing
  varQuads = Null
  varYears = Null
  Erase strYears
  Erase strQuads


End Sub

Public Sub LookYearly_Helper()

  Dim pMxDoc As IMxDocument
  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim pFeatDef As IFeatureLayerDefinition2
  Dim lngYearIndex As Long
  Dim lngQuadIndex As Long
  Dim strPrefix As String
  Dim strSuffix As String
  Dim pQueryFilt As IQueryFilter
  Dim pYearColl As New Collection
  Dim pQuadColl As New Collection
  Dim varVal As Variant
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  
  Set pMxDoc = ThisDocument
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Cover_All", pMxDoc.FocusMap)
  Set pFClass = pFLayer.FeatureClass
  Set pFeatDef = pFLayer
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix)
  lngYearIndex = pFClass.FindField("Year")
  lngQuadIndex = pFClass.FindField("Quadrat")
  
  Dim strYears() As String
  Dim strQuads() As String
  Dim lngYearCounter As Long
  Dim lngQuadCounter As Long
  Dim strYear As String
  Dim strQuad As String
  
  lngYearCounter = -1
  lngQuadCounter = -1
  
  
  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  
  Debug.Print "Sorting..."
  
  Do Until pFeature Is Nothing
    varVal = pFeature.Value(lngYearIndex)
    If Not IsNull(varVal) Then
      strYear = CStr(varVal)
      If Not MyGeneralOperations.CheckCollectionForKey(pYearColl, strYear) Then
        lngYearCounter = lngYearCounter + 1
        ReDim Preserve strYears(lngYearCounter)
        strYears(lngYearCounter) = strYear
        pYearColl.Add True, strYear
      End If
    End If
    
    varVal = pFeature.Value(lngQuadIndex)
    If Not IsNull(varVal) Then
      strQuad = CStr(varVal)
      If Not MyGeneralOperations.CheckCollectionForKey(pQuadColl, strQuad) Then
        lngQuadCounter = lngQuadCounter + 1
        ReDim Preserve strQuads(lngQuadCounter)
        strQuads(lngQuadCounter) = strQuad
        pQuadColl.Add True, strQuad
      End If
    End If
                      
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strYears, 0, UBound(strYears)
  QuickSort.StringsAscending strQuads, 0, UBound(strQuads)
  
  Dim lngIndex As Long
  Dim strReport As String
  strReport = "  dim varQuads as array" & vbCrLf & _
              "  varquads = array("
  For lngIndex = 0 To UBound(strQuads)
    strReport = strReport & """" & strQuads(lngIndex) & IIf(lngIndex < UBound(strQuads), """,", """")
  Next lngIndex
  strReport = strReport & ")" & vbCrLf & vbCrLf
  strReport = strReport & "  dim varYears as array" & vbCrLf & _
              "  varyears = array("
  For lngIndex = 0 To UBound(strYears)
    strReport = strReport & """" & strYears(lngIndex) & IIf(lngIndex < UBound(strYears), """,", """")
  Next lngIndex
  strReport = strReport & ")" & vbCrLf & vbCrLf
  
  Debug.Print strReport
  
  Debug.Print "Done..."
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFLayer = Nothing
  Set pFClass = Nothing
  Set pFeatDef = Nothing
  Set pQueryFilt = Nothing
  Set pYearColl = Nothing
  Set pQuadColl = Nothing
  varVal = Null
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strYears
  Erase strQuads


End Sub



Public Sub MakeBox()

  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  Set pEnv.SpatialReference = pSpRef
  pEnv.PutCoords 0, 0, 1, 1
  Dim pPoly As IPolygon
  Set pPoly = MyGeometricOperations.EnvelopeToPolygon(pEnv)
  Dim pArray As esriSystem.IArray
  Set pArray = New esriSystem.Array
  pArray.Add pPoly
  Dim pFClass As IFeatureClass
  Set pFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pArray)
  Dim pFLayer As IFeatureLayer
  Set pFLayer = New FeatureLayer
  Set pFLayer.FeatureClass = pFClass
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  pMxDoc.AddLayer pFLayer
  pMxDoc.UpdateContents

ClearMemory:
  Set pSpRef = Nothing
  Set pEnv = Nothing
  Set pPoly = Nothing
  Set pArray = Nothing
  Set pFClass = Nothing
  Set pFLayer = Nothing
  Set pMxDoc = Nothing



End Sub

Public Function ReturnArrayOfFieldLinks(pSrcFClass As IFeatureClass, pDestFClass As IFeatureClass) As Variant()

  Dim pSrcFields As IFields
  Dim pDestFields As IFields
  Set pSrcFields = pSrcFClass.Fields
  Set pDestFields = pDestFClass.Fields
  
  Dim pField As IField
  Dim lngIndex As Long
  Dim varReturn() As Variant
  Dim lngCounter As Long
  Dim strNewName As String
  
  ' varFieldIndexArray WILL HAVE 4 COLUMNS AND ANY NUMBER OR ROWS.
  ' COLUMN 0 = SOURCE FIELD NAME
  ' COLUMN 1 = SOURCE FIELD INDEX
  ' COLUMN 2 = NEW FIELD NAME
  ' COLUMN 3 = NEW FIELD INDEX
  
  lngCounter = -1
  For lngIndex = 0 To pSrcFields.FieldCount - 1
    Set pField = pSrcFields.Field(lngIndex)
'    Debug.Print pField.Name
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
      
'      Debug.Print varReturn(0, lngCounter) & " | " & varReturn(1, lngCounter) & _
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
  
  Dim pField As IField
  Dim lngIndex As Long
  Dim varReturn() As Variant
  Dim lngCounter As Long
  Dim strNewName As String
  
  ' varFieldIndexArray WILL HAVE 4 COLUMNS AND ANY NUMBER OR ROWS.
  ' COLUMN 0 = SOURCE FIELD NAME
  ' COLUMN 1 = SOURCE FIELD INDEX
  ' COLUMN 2 = NEW FIELD NAME
  ' COLUMN 3 = NEW FIELD INDEX
  
  lngCounter = -1
  For lngIndex = 0 To pSrcFields.FieldCount - 1
    Set pField = pSrcFields.Field(lngIndex)
'    Debug.Print pField.Name
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
      
'      Debug.Print varReturn(0, lngCounter) & " | " & varReturn(1, lngCounter) & _
          " | " & varReturn(2, lngCounter) & " | " & varReturn(3, lngCounter)
      
    End If
  Next lngIndex
  
'          ReDim varAltIndexArray(3, 4)
'          varAltIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestFClass)
          
  ReturnArrayOfFieldCrossLinks = varReturn
  
  Set pSrcFields = Nothing
  Set pDestFields = Nothing
  Erase varReturn

End Function
Public Sub CreateNewFields(pNewFClass As IFeatureClass, lngFClassNameIndex As Long, _
    lngQuadratIndex As Long, lngYearIndex As Long, lngTypeIndex As Long, lngOrigFIDIndex As Long)

  Dim pField As IField
  Dim pFieldEdit As IFieldEdit
  Dim lngIDIndex As Long
  Dim lngIsEmptyIndex As Long
  
  ' FORCE "SP_CODE" TO BE "SP_CODE"
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
    
  ' ARRAY BELOW ALLOWS FOR QUADRAT- AND YEAR-SPECIFIC CONVERSIONS
  ' GLOBAL CONVERSIONS ARE SPECIFIED IN FILES IN FillCollections FUNCTION
  ' varSpecialConversions: Rows for conversions, Columns as follows:
  ' 0) Quadrat
  ' 1) Year: -999 if All Years
  ' 2) Source Species
  ' 3) Converted Species
  ' 4) Note on Changes
  
  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long
  Dim strTestSpecies As String
  
  If InStr(1, strSpecies, "Muhlenbergia rigens", vbTextCompare) > 0 Then
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
          strNoteOnChanges = Trim(CStr(varSpecialConversions(4, lngIndex)))
          Exit Function
        End If
      End If
    End If
  Next lngIndex

End Function
Public Sub ReviseShapefiles()
    
  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"
  
  ' ARRAY BELOW ALLOWS FOR QUADRAT- AND YEAR-SPECIFIC CONVERSIONS
  ' GLOBAL CONVERSIONS ARE SPECIFIED IN FILES IN FillCollections FUNCTION
  ' varSpecialConversions: Rows for conversions, Columns as follows:
  ' 0) Quadrat
  ' 1) Year: -999 if All Years
  ' 2) Source Species
  ' 3) Converted Species
  ' 4) Note on Changes
  Dim varSpecialConversions() As Variant
  ReDim varSpecialConversions(4, 5)
  Dim strNoteOnChanges As String
  varSpecialConversions(0, 0) = "Q90"
  varSpecialConversions(1, 0) = -999
  varSpecialConversions(2, 0) = "Antennaria parvifolia"
  varSpecialConversions(3, 0) = "Antennaria rosulata"
  varSpecialConversions(4, 0) = "Email Margaret Dec. 21, 2018"
  
  varSpecialConversions(0, 1) = "Q93"
  varSpecialConversions(1, 1) = -999
  varSpecialConversions(2, 1) = "Elymus elymoides"
  varSpecialConversions(3, 1) = "Muhlenbergia montana"
  varSpecialConversions(4, 1) = "Email Margaret Dec. 21, 2018"
  
  varSpecialConversions(0, 2) = "Q93"
  varSpecialConversions(1, 2) = -999
  varSpecialConversions(2, 2) = "Poa fendleriana"
  varSpecialConversions(3, 2) = "Muhlenbergia montana"
  varSpecialConversions(4, 2) = "Email Margaret Dec. 21, 2018"
  
  varSpecialConversions(0, 3) = "Q80"
  varSpecialConversions(1, 3) = -999
  varSpecialConversions(2, 3) = "Muhlenbergia tricholepis"
  varSpecialConversions(3, 3) = "Bouteloua gracilis"
  varSpecialConversions(4, 3) = "Email Margaret Dec. 21, 2018"
  
  varSpecialConversions(0, 4) = "Q80"
  varSpecialConversions(1, 4) = -999
  varSpecialConversions(2, 4) = "Muhlenbergia rigens"
  varSpecialConversions(3, 4) = "Muhlenbergia wrightii"
  varSpecialConversions(4, 4) = "Email Margaret Dec. 21, 2018"
  
  varSpecialConversions(0, 5) = "Q88"
  varSpecialConversions(1, 5) = -999
  varSpecialConversions(2, 5) = "Unknown forb"
  varSpecialConversions(3, 5) = "Coreopsis tinctoria"
  varSpecialConversions(4, 5) = "Email Margaret Dec. 21, 2018"
  
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

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  strRoot = "E:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data - March_1_2018b"
  
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
    
  ' REMEMBER TO REMOVE INITIAL SPACES
  ' REMEMBER TO CHANGE GRAMMINOID TO GRAMINOID
  ' REMEMBER TO REMOVE LINE RETURNS
  
  For lngIndex = 0 To pFolders.Count - 1
    DoEvents
    strFolder = pFolders.Element(lngIndex)
'    strFolder = "E:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data - March_1_2018b\Q46"
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
              strExcelReport, varSpecialConversions)
            
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
  pDataObj.Clear
  pDataObj.SetText strReport & vbCrLf & "-----------------------------------" & vbCrLf & strExcelFullReport
  pDataObj.PutInClipboard
  
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pDataObj = Nothing




End Sub

Public Sub ReplaceNamesInShapefile(pDataset As IDataset, pCheckCollection As Collection, booMadeEdits As Boolean, _
    strEditReport As String, strBase As String, strExcelReport As String, varSpecialConversions() As Variant)
  
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
  
  lngIndex = -1
  
  Dim pDoneColl As New Collection
  
'  Debug.Print "Updating '" & pDataset.BrowseName & "'..."
  strEditReport = "Edits to '" & pDataset.BrowseName & "':" & vbCrLf
  strExcelReport = ""
  
  lngNameIndex = pFClass.FindField("Species")
  If lngNameIndex = -1 Then lngNameIndex = pFClass.FindField("Cover")
  If lngNameIndex = -1 Then
    'booFoundData = False
    MsgBox "Unexpected Event!"
    DoEvents
    GoTo ClearMemory
  End If
  
  Dim strSplit() As String
  Dim strQuadrat As String
  Dim strYear As Long
  Dim strNoteOnChanges As String
  strSplit = Split(pDataset.BrowseName, "_")
  strQuadrat = strSplit(0)
  strYear = strSplit(1)
  
  ' ARRAY BELOW ALLOWS FOR QUADRAT- AND YEAR-SPECIFIC CONVERSIONS
  ' GLOBAL CONVERSIONS ARE SPECIFIED IN FILES IN FillCollections FUNCTION
  ' varSpecialConversions: Rows for conversions, Columns as follows:
  ' 0) Quadrat
  ' 1) Year: -999 if All Years
  ' 2) Source Species
  ' 3) Converted Species
  ' 4) Note on Changes
  
  Set pFCursor = pFClass.Update(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    strOrigName = strName
    
    ' REMOVE CARRIAGE RETURNS AND TRIM
    strName = Replace(strName, vbCrLf, "")
    strName = Replace(strName, vbNewLine, "")
    strName = Trim(strName)
    
    ' BY DEFAULT, ASSUME NAME IS CORRECT.  ONLY CHANGE IT IF WE FIND A REPLACEMENT VALUE
'    strCorrect = strName
    strCorrect = CheckSpeciesAgainstSpecialConversions(varSpecialConversions, strQuadrat, CLng(strYear), _
                strName, strNoteOnChanges)
    
    strHexify = HexifyName(strName)
    If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
      strCorrect = pCheckCollection.Item(strHexify)
    End If
    
    ' SPECIAL CASES
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
'      Debug.Print "  --> " & CStr(pFeature.OID) & "] Changing '" & strName & "' to '" & strCorrect & "'..."
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

End Sub



Public Sub TestFillCollections()
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

Debug.Print "Should have changed from Cover to Density:" & vbCrLf & strCoverToDensityQuery
Debug.Print "Should have changed from Density to Cover:" & vbCrLf & strDensityToCoverQuery

Debug.Print "Done..."

ClearMemory:
  Set pCoverCollection = Nothing
  Set pDensityCollection = Nothing


End Sub

Public Sub FillCollections(pCoverCollection As Collection, pDensityCollection As Collection, _
    Optional pCoverToDensity As Collection, Optional pDensityToCover As Collection, _
    Optional strCoverToDensityQuery As String, Optional strDensityToCoverQuery As String, _
    Optional pCoverShouldChangeColl As Collection, Optional pDensityShouldChangeColl As Collection)

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pTable As ITable
  Set pWSFact = New ExcelWorkspaceFactory
  
  Dim pTestWS As IFeatureWorkspace
  Dim pTestWSFact As IWorkspaceFactory
  Set pTestWSFact = New FileGDBWorkspaceFactory
  Set pTestWS = pTestWSFact.OpenFromFile("E:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data\Combined_by_Quadrat.gdb", 0)
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
  
  ' PATHS BELOW LIST ALL SPECES BY ORIGINAL NAMES, ALL SPECIES NAMES THAT ORIGINALS SHOULD BE CHANGED TO, AND WHETHER
  ' COVER AND DENSITY SHOULD BE SWITCHED.
  
  varPaths = Array("E:\arcGIS_stuff\consultation\Margaret_Moore\species_list_Cover_changes_Dec_2_2017.xlsx", _
                   "E:\arcGIS_stuff\consultation\Margaret_Moore\Species_list_Density_changes_Dec_2_2017.xlsx")
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
  
  ' TO CONFIRM THAT ONLY ONE POSSIBLE VALUE IN "SHOULD CHANGE" COMMENT FIELD
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
      
      ' FOR DEBUG
      If strIncorrect = "Muhlenbergia rigens" Then
        DoEvents
      End If
      
      If InStr(1, strShouldChange, "change to", vbTextCompare) > 0 Then booShouldChange = True
      
      Set pRow = pCursor.NextRow
      
      ' index 0 = cover, index 1 = density
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
      
'      If InStr(1, strIncorrect, "Muhlenbergia tricholepis", vbTextCompare) > 0 Then
'        DoEvents
'      End If
      
      If InStr(1, strIncorrect, "Erigeron formosissimus", vbTextCompare) > 0 Then
        DoEvents
      End If
      
      ' FORCE "CORRECT" VERSIONS TO BE TRIMMED WITHOUT NEWLINES
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
'        Debug.Print "Adding Incorrect = " & strIncorrect & ", Correct = " & strCorrect
        pColl.Add strCorrect, strHexIncorrect
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
          pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
        End If
        
        lngIndex2 = lngIndex2 + 1
        ReDim Preserve strVals(lngIndex2)
        strVals(lngIndex2) = strIncorrect
                
        If lngIndex = 0 And booShouldChangeFromCover Then
          pCoverToDensity.Add strCorrect, strHexIncorrect
          
          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
  
        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          pDensityToCover.Add strCorrect, strHexIncorrect
          
          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
  
        End If
        
        ' UNLIKELY THAT REPLACEMENTS OR INCORRECT VERSIONS INTENTIONALLY HAVE CARRIAGE RETURNS, SO ALSO ADD
        ' VERSION WITHOUT
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
  '        Debug.Print "Adding Incorrect = " & strIncorrect & ", Correct = " & strCorrect
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
        
        ' IF INCORRECT VERSIONS START WTIH A SPACE, INCLUDE TRIMMED VERSION OF INCORRECT
        If Trim(strIncorrect) <> strIncorrect Then
          strIncorrectVariant = Trim(strIncorrect)
            
          strHexIncorrect = HexifyName(strIncorrectVariant)
  '        Debug.Print "Adding Incorrect = " & strIncorrect & ", Correct = " & strCorrect
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
  
'  Debug.Print "Checking Cover:"
'  strCover = varVals(0)
'  For lngIndex = 0 To UBound(strCover)
'    strIncorrect = strCover(lngIndex)
'    strHexIncorrect = HexifyName(strIncorrect)
'    If MyGeneralOperations.CheckCollectionForKey(pCoverCollection, strHexIncorrect) Then
'      Debug.Print "--> '" & strIncorrect & "' converts to '" & pCoverCollection.Item(strHexIncorrect) & "'..."
'
'      If MyGeneralOperations.CheckCollectionForKey(pDensityCollection, strHexIncorrect) Then
'        If StrComp(pCoverCollection.Item(strHexIncorrect), pDensityCollection.Item(strHexIncorrect), vbBinaryCompare) <> 0 Then
'          Debug.Print "*** " & strIncorrect & "' converts to Cover '" & pCoverCollection.Item(strHexIncorrect) & _
'                    "' and Density '" & pDensityCollection.Item(strHexIncorrect) & "'..."
'        End If
'      End If
'    End If
'  Next lngIndex
'
'  Debug.Print "Checking Density:"
'  strDensity = varVals(1)
'  For lngIndex = 0 To UBound(strDensity)
'    strIncorrect = strDensity(lngIndex)
'    strHexIncorrect = HexifyName(strIncorrect)
'    If MyGeneralOperations.CheckCollectionForKey(pDensityCollection, strHexIncorrect) Then
'      Debug.Print "--> '" & strIncorrect & "' converts to '" & pDensityCollection.Item(strHexIncorrect) & "'..."
'
'      If MyGeneralOperations.CheckCollectionForKey(pCoverCollection, strHexIncorrect) Then
'        If StrComp(pCoverCollection.Item(strHexIncorrect), pDensityCollection.Item(strHexIncorrect), vbBinaryCompare) <> 0 Then
'          Debug.Print "*** " & strIncorrect & "' converts to Cover '" & pCoverCollection.Item(strHexIncorrect) & _
'                    "' and Density '" & pDensityCollection.Item(strHexIncorrect) & "'..."
'        End If
'      End If
'    End If
'  Next lngIndex
  
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
Public Sub ReviewShapefiles_IncludeType()
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  strRoot = "E:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data"
  
  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long
  
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
  
'  pNameColl.Add True, "CASE"
'  pNameColl.Add True, "case"
  Dim strCoverOrDens As String
  
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
        strNames = ReturnListOfNames(pDataset, booFoundNames)
        
        If InStr(1, pDataset.BrowseName, "_C", vbTextCompare) > 0 Then
          strCoverOrDens = " [Cover]"
        ElseIf InStr(1, pDataset.BrowseName, "_D", vbTextCompare) > 0 Then
          strCoverOrDens = " [Density]"
        Else
          strCoverOrDens = " [Neither]"
        End If
        
        If booFoundNames Then
          lngAcceptSFCount = lngAcceptSFCount + 1
          lngRecCount = lngRecCount + UBound(strNames) + 1
          For lngNameIndex = 0 To UBound(strNames)
            strName = strNames(lngNameIndex) & strCoverOrDens
            strHexify = HexifyName(strName)
            
            If strName = "No point species" Then
              DoEvents
            End If
            If Left(strName, 1) = " " Then
              DoEvents
            End If
            If MyGeneralOperations.CheckCollectionForKey(pNameColl, strHexify) Then
              lngNameCount = pNameColl.Item(strHexify)
              pNameColl.Remove strHexify
            Else
              lngNameCount = 0
              lngFullNameCounter = lngFullNameCounter + 1
              ReDim Preserve strFullNames(lngFullNameCounter)
              strFullNames(lngFullNameCounter) = strHexify
            End If
            lngNameCount = lngNameCount + 1
'            Debug.Print " --> Adding " & strName
            pNameColl.Add lngNameCount, strHexify
          Next lngNameIndex
        End If
      Next lngDatasetIndex
      
    End If
    
  Next lngIndex
  
  DoEvents
  Dim strFullNamesWords() As String
  ReDim strFullNamesWords(UBound(strFullNames))
       
  For lngIndex = 0 To UBound(strFullNames)
    strHexify = strFullNames(lngIndex)
    strName = WordifyHex(strHexify)
    strFullNamesWords(lngIndex) = "Species Name = '" & strName & "'  [n = " & Format(pNameColl.Item(strHexify), "#,##0") & "]"
  Next lngIndex
  
  QuickSort.StringsAscending strFullNamesWords, 0, UBound(strFullNamesWords)
  
  Dim strReport As String
  strReport = "Root Folder: " & strRoot & vbCrLf & _
              "Shapefiles Examined: " & Format(lngShapefileCount, "#,##0") & vbCrLf & _
              "Shapefiles with Species Fields found: " & Format(lngAcceptSFCount, "#,##0") & vbCrLf & _
              "Species Values Examined: " & Format(lngRecCount, "#,##0") & vbCrLf & _
              "Unique Species Names: " & Format(UBound(strFullNames) + 1, "#,##0") & vbCrLf & _
              "-------------------------------------------" & vbCrLf
  
  Dim strNumber As String
  
  For lngIndex = 0 To UBound(strFullNamesWords)
    strName = strFullNamesWords(lngIndex)
    strNumber = CStr(lngIndex + 1)
    strNumber = MyGeneralOperations.SpacesInFrontOfText(strNumber, 3)
    strReport = strReport & strNumber & "] " & strName & vbCrLf
  Next lngIndex
       
  Dim pDataObj As New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReport
  pDataObj.PutInClipboard
  
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pDataObj = Nothing




End Sub
Public Sub ReviewShapefiles()
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  strRoot = "E:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data"
  
  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long
  
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
  
'  pNameColl.Add True, "CASE"
'  pNameColl.Add True, "case"
  
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
        strNames = ReturnListOfNames(pDataset, booFoundNames)
        
        If booFoundNames Then
          lngAcceptSFCount = lngAcceptSFCount + 1
          lngRecCount = lngRecCount + UBound(strNames) + 1
          For lngNameIndex = 0 To UBound(strNames)
            strName = strNames(lngNameIndex)
            strHexify = HexifyName(strName)
            
            If strName = "No point species" Then
              DoEvents
            End If
            If Left(strName, 1) = " " Then
              DoEvents
            End If
            If MyGeneralOperations.CheckCollectionForKey(pNameColl, strHexify) Then
              lngNameCount = pNameColl.Item(strHexify)
              pNameColl.Remove strHexify
            Else
              lngNameCount = 0
              lngFullNameCounter = lngFullNameCounter + 1
              ReDim Preserve strFullNames(lngFullNameCounter)
              strFullNames(lngFullNameCounter) = strHexify
            End If
            lngNameCount = lngNameCount + 1
'            Debug.Print " --> Adding " & strName
            pNameColl.Add lngNameCount, strHexify
          Next lngNameIndex
        End If
      Next lngDatasetIndex
      
    End If
    
  Next lngIndex
  
  DoEvents
  Dim strFullNamesWords() As String
  ReDim strFullNamesWords(UBound(strFullNames))
       
  For lngIndex = 0 To UBound(strFullNames)
    strHexify = strFullNames(lngIndex)
    strName = WordifyHex(strHexify)
    strFullNamesWords(lngIndex) = "Species Name = '" & strName & "'  [n = " & Format(pNameColl.Item(strHexify), "#,##0") & "]"
  Next lngIndex
  
  QuickSort.StringsAscending strFullNamesWords, 0, UBound(strFullNamesWords)
  
  Dim strReport As String
  strReport = "Root Folder: " & strRoot & vbCrLf & _
              "Shapefiles Examined: " & Format(lngShapefileCount, "#,##0") & vbCrLf & _
              "Shapefiles with Species Fields found: " & Format(lngAcceptSFCount, "#,##0") & vbCrLf & _
              "Species Values Examined: " & Format(lngRecCount, "#,##0") & vbCrLf & _
              "Unique Species Names: " & Format(UBound(strFullNames) + 1, "#,##0") & vbCrLf & _
              "-------------------------------------------" & vbCrLf
  
  Dim strNumber As String
  
  For lngIndex = 0 To UBound(strFullNamesWords)
    strName = strFullNamesWords(lngIndex)
    strNumber = CStr(lngIndex + 1)
    strNumber = MyGeneralOperations.SpacesInFrontOfText(strNumber, 3)
    strReport = strReport & strNumber & "] " & strName & vbCrLf
  Next lngIndex
       
  Dim pDataObj As New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReport
  pDataObj.PutInClipboard
  
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pDataObj = Nothing




End Sub

Public Function HexifyName(strName As String) As String
'  Dim strChar As String
'  Dim strReturn As String
'  For lngIndex = 1 To Len(strName)
'    strChar = CStr(Hex(Asc(Mid(strName, lngIndex, 1))))
'    strChar = String(3 - Len(strChar), "0") & strChar
'    strReturn = strReturn & strChar
'  Next lngIndex
'  HexifyName = strReturn
'
'    Dim i As Long
  Dim lngIndex As Long
  HexifyName = Space$(Len(strName) * 4)
  For lngIndex = 0 To Len(strName) - 1
      Mid$(HexifyName, lngIndex * 4 + 1, 4) = Right$("0000" & Hex$(AscW(Mid$(strName, lngIndex + 1, 1))), 4)
  Next lngIndex
End Function

Public Function WordifyHex(strHexed As String) As String

  Dim lngIndex As Long
'  lngIndex = 1
'  Dim strReturn As String
'
'  Do Until lngIndex > Len(strHexed)
'    strReturn = strReturn & Chr(CLng("&H" & Mid(strHexed, lngIndex, 3)))
'
'    lngIndex = lngIndex + 3
'  Loop
'
'  WordifyHex = strReturn
'
'
'  Dim i As Long
  WordifyHex = Space$(Len(strHexed) \ 4)
  For lngIndex = 0 To Len(strHexed) - 1 Step 4
      Mid$(WordifyHex, lngIndex \ 4 + 1, 1) = ChrW$(Val("&h" & Mid$(strHexed, lngIndex + 1, 4)))
  Next lngIndex
  
End Function



Public Function ReturnListOfNames(pDataset As IDataset, booFoundData As Boolean) As String()
  
  On Error GoTo ErrHandler
  
  Dim pFClass As IFeatureClass
  Set pFClass = pDataset
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngNameIndex As Long
  Dim lngIndex As Long
  Dim strName As String
  Dim strReturn() As String
  
  lngIndex = -1
  
  Dim pDoneColl As New Collection
  
  lngNameIndex = pFClass.FindField("Species")
  If lngNameIndex = -1 Then lngNameIndex = pFClass.FindField("Cover")
  If lngNameIndex = -1 Then
    booFoundData = False
    GoTo ClearMemory
  End If
  
  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    If Trim(strName) = "" Then strName = "<- No Value -> [from " & pDataset.BrowseName & "]"
    lngIndex = lngIndex + 1
    ReDim Preserve strReturn(lngIndex)
    strReturn(lngIndex) = strName
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  ReturnListOfNames = strReturn
  booFoundData = lngIndex > -1
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  booFoundData = False
  
  
ClearMemory:
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strReturn
  Set pDoneColl = Nothing

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

Public Sub TestExportCSV()
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Set pApp = Application
  
  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim strFilename As String
  Dim booSucceeded As Boolean
  Dim strResult As String
  Dim varFields As Variant
  
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_v2", pMxDoc.FocusMap)
'  varFields = Array("OBJECTID", "ADM2_CODE", "ADM2_NAME", _
'          "STATUS", "DISP_AREA", "STR_YEAR", "EXP_YEAR", _
'          "ADM0_CODE", "ADM0_NAME", "ADM1_CODE", "ADM1_NAME", _
'          "Density_UN", "Area (Sq. Km.)", "UN-Adjusted Population", "UN-Adjusted Population Density")
'
'  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\GAUL_All.csv"
'
'  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
'      False, False, booSucceeded, varFields, pApp)
'
'  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult
  
  
  
  ' Population_by_GAUL_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", "ADM0_CODE", "ADM0_NAME", "ADM1_CODE", "ADM1_NAME", _
          "ADM2_CODE", "ADM2_NAME", "STATUS", "DISP_AREA", "STR_YEAR", "EXP_YEAR", "Sph_Area", _
          "Density_UN", "Area (Sq. Km.)", "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\GAUL_All.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult

' Population_by_GAUL_Dissolve_Level_2_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_Dissolve_Level_2_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", "ADM0_CODE", "ADM0_NAME", "ADM1_CODE", "ADM1_NAME", _
          "ADM2_CODE", "ADM2_NAME", "STATUS", "DISP_AREA", "STR_YEAR", "EXP_YEAR", "SUM_Sph_Area", _
          "Area (Sq. Km.)", "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\GAUL_2.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult

' Population_by_GAUL_Dissolve_Level_1_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_Dissolve_Level_1_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", "ADM0_CODE", "ADM0_NAME", _
          "ADM1_CODE", "ADM1_NAME", "SUM_Sph_Area", "Area (Sq. Km.)", _
          "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\GAUL_1.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult

' Population_by_GAUL_Dissolve_Level_0_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_Dissolve_Level_0_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", "ADM0_CODE", "ADM0_NAME", _
          "SUM_Sph_Area", "Area (Sq. Km.)", "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\GAUL_0.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult

' hydrobasins_world_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("hydrobasins_world_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", _
          "MAJ_BAS", "MAJ_NAME", "MAJ_AREA", "SUB_BAS", "TO_BAS", "SUB_NAME", "SUB_AREA", _
          "LEGEND", "Sph_Area", "Area (Sq. Km.)", _
          "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\Hydrobasins_All.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult

' hydrobasins_Sub_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("hydrobasins_Sub_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", "MAJ_BAS", "MAJ_NAME", "SUB_BAS", _
          "SUB_NAME", "SUM_Sph_Area", "Area (Sq. Km.)", _
          "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\Hydrobasins_Sub_Major.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult

' hydrobasins_Major_v2
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("hydrobasins_Major_v2", pMxDoc.FocusMap)
  varFields = Array("OBJECTID", "MAJ_BAS", "MAJ_NAME", _
          "SUM_Sph_Area", "Area (Sq. Km.)", "UN-Adjusted Population", "UN-Adjusted Population Density")
  strFilename = "E:\arcGIS_stuff\consultation\UN_FAO\2016_Population_Density\CSVs\Hydrobasins_Major.csv"
  strResult = MyGeneralOperations.ExportToCSV(pFLayer, strFilename, True, _
      False, False, booSucceeded, varFields, pApp)
  Debug.Print "Succeeded = " & CStr(booSucceeded) & vbCrLf & "  --> Message: " & strResult


  
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pFLayer = Nothing
  Set pFClass = Nothing
  varFields = Null

End Sub

Public Sub CodeHelper()
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Set pApp = Application
  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim pFields As IFields
  Dim lngIndex As Long
  Dim strReport As String
  Dim pField As IField
  Dim varFLayers As Variant
  Dim lngLayerIndex As Long
  varFLayers = Array(MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_v2", pMxDoc.FocusMap), _
      MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_Dissolve_Level_2_v2", pMxDoc.FocusMap), _
      MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_Dissolve_Level_1_v2", pMxDoc.FocusMap), _
      MyGeneralOperations.ReturnLayerByName("Population_by_GAUL_Dissolve_Level_0_v2", pMxDoc.FocusMap), _
      MyGeneralOperations.ReturnLayerByName("hydrobasins_world_v2", pMxDoc.FocusMap), _
      MyGeneralOperations.ReturnLayerByName("hydrobasins_Sub_v2", pMxDoc.FocusMap), _
      MyGeneralOperations.ReturnLayerByName("hydrobasins_Major_v2", pMxDoc.FocusMap))
    
  For lngLayerIndex = 0 To UBound(varFLayers)
    Set pFLayer = varFLayers(lngLayerIndex)
    Set pFClass = pFLayer.FeatureClass
    Set pFields = pFClass.Fields
    
    strReport = strReport & "' " & pFLayer.Name & vbCrLf & _
        "  Set pFLayer = MyGeneralOperations.ReturnLayerByName(""" & pFLayer.Name & """, pMxDoc.FocusMap)" & vbCrLf & _
        "  varFields = array("
    For lngIndex = 0 To pFields.FieldCount - 1
      Set pField = pFields.Field(lngIndex)
      If lngIndex > 0 And lngIndex Mod 4 = 0 Then strReport = strReport & " _" & vbCrLf & "          "
      If pField.Type <> esriFieldTypeGeometry Then
        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & _
            IIf(lngIndex = pFields.FieldCount - 1, ")" & vbCrLf, ",")
      End If
    Next lngIndex
    strReport = strReport & vbCrLf
  Next lngLayerIndex
  
  Debug.Print strReport
  
  Dim pDataObj As New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReport
  pDataObj.PutInClipboard
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pFLayer = Nothing
  Set pFClass = Nothing
  Set pFields = Nothing
  Set pField = Nothing
  Set pDataObj = Nothing
  Erase varFLayers


  
End Sub

Public Sub FillDistanceAndDirection(varLocationArray() As Variant, dblX As Double, dblY As Double, _
      dblMinDist As Double, dblNearDir As Double, strNearID As String)
  
  Dim lngIndex As Long
  Dim dblTestDist As Double
  Dim dblTestX As Double
  Dim dblTestY As Double
  
  dblTestX = CDbl(varLocationArray(1, 0))
  dblTestY = CDbl(varLocationArray(2, 0))
  strNearID = CStr(varLocationArray(0, 0))
  
  dblMinDist = MyGeometricOperations.DistancePythagoreanNumbers(dblX, dblY, dblTestX, dblTestY)
  dblNearDir = MyGeometricOperations.CalcBearingNumbers(dblX, dblY, dblTestX, dblTestY)
  
  For lngIndex = 1 To UBound(varLocationArray, 2)
    dblTestX = CDbl(varLocationArray(1, lngIndex))
    dblTestY = CDbl(varLocationArray(2, lngIndex))
    dblTestDist = MyGeometricOperations.DistancePythagoreanNumbers(dblX, dblY, dblTestX, dblTestY)
    If dblTestDist < dblMinDist Then
      dblMinDist = dblTestDist
      strNearID = CStr(varLocationArray(0, lngIndex))
      dblNearDir = MyGeometricOperations.CalcBearingNumbers(dblX, dblY, dblTestX, dblTestY)
    End If
  Next lngIndex
  
  
End Sub

Public Function MakeLocationArray(pPointFClass As IFeatureClass, lngIDIndex As Long, _
    Optional strQueryString As String = "", Optional pSpRef As ISpatialReference = Nothing) As Variant()

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPoint As IPoint
  Dim lngIndex As Long
  Dim pQueryFilt As IQueryFilter
  Dim varID As Variant
  
  If strQueryString <> "" Then
    Set pQueryFilt = New QueryFilter
    pQueryFilt.WhereClause = strQueryString
  Else
    Set pQueryFilt = Nothing
  End If
  
  Dim varReturn() As Variant
  ReDim varReturn(2, pPointFClass.FeatureCount(Nothing) - 1)
  
  Set pFCursor = pPointFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  lngIndex = -1
  
  Do Until pFeature Is Nothing
    Set pPoint = pFeature.ShapeCopy
    If Not pSpRef Is Nothing Then pPoint.Project pSpRef
    
    varID = pFeature.Value(lngIDIndex)
    
    lngIndex = lngIndex + 1
    varReturn(0, lngIndex) = varID
    varReturn(1, lngIndex) = pPoint.X
    varReturn(2, lngIndex) = pPoint.Y
    
    Set pFeature = pFCursor.NextFeature
  Loop
  
  MakeLocationArray = varReturn
  
ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pPoint = Nothing
  Set pQueryFilt = Nothing
  varID = Null
  Erase varReturn


  
End Function


