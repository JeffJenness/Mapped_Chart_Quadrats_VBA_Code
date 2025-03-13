Attribute VB_Name = "Margaret"
Option Explicit

Public Sub DeclareWorkspaces(strOrigShapefiles As String, Optional strModifiedRoot As String, _
    Optional strShiftedRoot As String, Optional strExportBase As String, Optional strRecreatedModifiedRoot As String, _
    Optional strSetFolder As String, Optional strExtractShapefileFolder As String, Optional strFinalFolder As String)

  Dim booUseCurrentDate As Boolean
  booUseCurrentDate = False

  Dim strSpecifiedDate As String
  strSpecifiedDate = "2025_02_27"

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
  strExportBase = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Final_Datasets_" & strCurrentDate & "\Quadrat_Map_Image_Files"
  strSetFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate
  strExtractShapefileFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Shapefile_Extractions_" & strCurrentDate
  strFinalFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\" & strDate & "\Final_Datasets_" & strCurrentDate

End Sub

Public Sub ExportImages()

  Debug.Print "----------------"
  Dim lngStart As Long
  lngStart = GetTickCount
  Dim pEnv As IEnvelope

  Dim booOnlyDoSpecificQuadrats As Boolean
  Dim varSpecificQuadrats() As Variant
  Dim booExportThisOne As Boolean

  booOnlyDoSpecificQuadrats = False
  varSpecificQuadrats = Array("44", "43", "75")

  Dim strExportBase As String
  Dim strModifiedRoot As String
  Dim strOrig As String

  Call ClearAnyInitialData

  Call DeclareWorkspaces(strOrig, , , strExportBase, strModifiedRoot)
  If Right(strExportBase, 1) <> "\" Then strExportBase = strExportBase & "\"
  MyGeneralOperations.CreateNestedFoldersByPath strExportBase

  Dim pSourceWS As IFeatureWorkspace
  Dim pSourceWSFact As IWorkspaceFactory
  Set pSourceWSFact = New FileGDBWorkspaceFactory
  Set pSourceWS = pSourceWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)

  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Set pDensityFClass = pSourceWS.OpenFeatureClass("Density_All")
  Set pCoverFClass = pSourceWS.OpenFeatureClass("Cover_All")

  Dim pWS2 As IWorkspace2
  Set pWS2 = pSourceWS
  Dim pSymTable As ITable
  Dim pSymRow As IRow
  Dim lngBLOBIndex As Long
  Dim pSymbolColl As Collection
  Dim pSymCursor As ICursor
  Dim pField As IField
  Dim pFieldEdit As IFieldEdit
  Dim pFields As esriSystem.IVariantArray

  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strQuadrat As String
  Dim strFolder As String
  Dim strFileHeader As String
  Dim strItem() As String

  Dim pLocationsAndNotes As Collection
  Dim strPlotNames() As String
  Call ReturnQuadratCoordsAndNames(pLocationsAndNotes, strPlotNames)
  Dim varLocNotes() As Variant
  Dim strFinalQuadrats() As String
  Dim pQuadratNumColl As Collection
  Set pQuadratNumColl = FillQuadratNameColl_Rev(strFinalQuadrats)
  Dim strItems() As String
  Dim strNote As String

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngIndex3 As Long
  Dim strYears() As String
  Dim strQuads() As String
  Dim lngYearCounter As Long
  Dim lngQuadCounter As Long
  Dim strYear As String
  Dim strQuad As String
  Dim strQueryString As String
  Dim strCheck() As String
  Dim lngCheckCounter As Long
  Dim pGeoDataset As IGeoDataset
  Dim lngSpeciesIndex As Long
  Dim lngPointSpeciesIndex As Long

  Dim pLayersToDelete As esriSystem.IArray
  Set pLayersToDelete = New esriSystem.Array

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFLayer As IFeatureLayer

  Dim pFeatDef As IFeatureLayerDefinition2
  Dim lngYearIndex As Long
  Dim lngQuadIndex As Long
  Dim lngPointYearIndex As Long
  Dim lngPointQuadIndex As Long
  Dim strPrefix As String
  Dim strSuffix As String
  Dim strPointPrefix As String
  Dim strPointSuffix As String
  Dim pQueryFilt As IQueryFilter
  Dim pYearColl As New Collection
  Dim pQuadColl As New Collection
  Dim varVal As Variant
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPointFClass As IFeatureClass
  Dim pFClass As IFeatureClass

  Dim strExportFolder As String

  Dim lngYearIndex2 As Long
  Dim strExportFilename As String

  Set pFClass = pCoverFClass
  Set pPointFClass = pDensityFClass

  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pPointFClass, strPointPrefix, strPointSuffix)
  lngPointYearIndex = pPointFClass.FindField("Year")
  lngPointQuadIndex = pPointFClass.FindField("Quadrat")
  lngPointSpeciesIndex = pPointFClass.FindField("Species")

  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix)
  lngYearIndex = pFClass.FindField("Year")
  lngQuadIndex = pFClass.FindField("Quadrat")
  lngSpeciesIndex = pFClass.FindField("Species")

  Dim varQuads As Variant
  Dim varYears As Variant

  Dim lngIndex As Long
  ReDim varQuads(499)
  For lngIndex = 1 To 500
    varQuads(lngIndex - 1) = "Q" & Format(lngIndex, "0")
  Next lngIndex
  ReDim varYears(25)
  For lngIndex = 2000 To 2025
    varYears(lngIndex - 2000) = Format(lngIndex, "0")
  Next lngIndex

  Dim pPointSymbols As Collection
  Dim pFillSymbols As Collection

  Dim lngPointSymIndex As Long
  Dim lngFillSymIndex As Long
  Dim pSymBuffer As IRowBuffer

  Dim strSymbolText As String
  Dim strSymbolTextPath As String
  strSymbolTextPath = strExportBase & "Symbology_Instructions.txt"

  If aml_func_mod.FileExists(strSymbolTextPath) Then
    strSymbolText = MyGeneralOperations.ReadTextFile(strSymbolTextPath)
    Call CreateCollectionOfSymbolsFromTextFile(strSymbolText, pPointSymbols, pFillSymbols, pSBar, pProg)
  Else

    Call CreateCollectionOfSymbols(pPointFClass, lngPointSpeciesIndex, _
        pFClass, lngSpeciesIndex, pPointSymbols, pFillSymbols, pSBar, pProg, strSymbolText)
    MyGeneralOperations.WriteTextFile strSymbolTextPath, strSymbolText, True, False
  End If

  Set pQueryFilt = New QueryFilter
  lngCheckCounter = -1
  Dim pPolyColl As Collection
  Dim strSpecies() As String
  Dim booFoundSomething As Boolean
  Dim lngCount As Long
  Dim lngHighCount As Long
  Dim strHighQuad As String
  Dim lngPointCount As Long
  Dim lngPointHighCount As Long
  Dim strPointHighQuad As String
  strHighQuad = ""
  strPointHighQuad = ""
  lngHighCount = -999
  lngPointHighCount = -999
  Dim pPolygon As IPolygon
  Dim varPolys() As Variant
  Dim pArea As IArea
  Dim dblArea As Double
  Dim dblCumulative As Double
  Dim strVal As String
  Dim strObsCount As String

  Dim strReport1 As String
  Dim strReport2 As String
  Dim pDataObj As New MSForms.DataObject

  Dim pNewFClass As IFeatureClass
  Dim varWorkOrder() As Variant
  varWorkOrder = Array("Cover", "Density")
  Dim lngWorkIndex As Long
  Dim strWorkOption As String
  Dim pDataset As IDataset
  Dim pSubLayer As IFeatureLayer
  Dim lngDeleteIndex As Long
  Dim pSymbol As ISymbol

  Dim strSiteName As String
  Dim strPlotName As String
  Dim strCrewName As String
  Dim strPhoto As String
  Dim strDate As String
  Dim strUTME As String
  Dim strUTMN As String
  Dim strComment As String
  Dim strElev As String

  Dim lngMaxAllSpeciesCount As Long
  Dim strMaxAllSpeciesQuad As String

  Dim pAreaColl As Collection
  Dim dblCurrentArea As Double

  Dim pLegendColl As Collection
  Dim strLegendKeys() As String

  Dim lngQuadCount As Long
  lngQuadCount = UBound(varQuads) + 1
  Dim strQuadCount As String
  strQuadCount = CStr(lngQuadCount)

  pSBar.ShowProgressBar "Working on Quad", 0, lngQuadCount, 1, True
  pProg.position = 0

  Dim strSaveFolder As String

  For lngYearIndex2 = 0 To UBound(varYears)
    strYear = CStr(varYears(lngYearIndex2))
    strExportFolder = strExportBase & strYear & "\"

    For lngIndex = 0 To lngQuadCount - 1
      strQuad = varQuads(lngIndex)

      strQuadrat = Replace(strQuad, "Q", "", , , vbTextCompare)

      If MyGeneralOperations.CheckCollectionForKey(pQuadratNumColl, strQuadrat) Then
        strItem = pQuadratNumColl.Item(strQuadrat)
        strSite = strItem(0)
        strSiteSpecific = strItem(1)
        strPlot = strItem(2)
        strFolder = strItem(4)
        strFileHeader = strItem(5)
        strUTME = strItem(6)
        strUTMN = strItem(7)
        strComment = strItem(8)
        strElev = strItem(9)

        strSiteName = strSiteSpecific
        strPlotName = strPlot
        strCrewName = "HED, JDS"
        strPhoto = ""
        strDate = strYear

      Else
        strSiteName = "Site for Quad " & strQuad
        strPlotName = "Plot for Quad " & strQuad
        strCrewName = "<-- Need Crew Name -->"
        strPhoto = "<-- Need Photo Number -->"
        strDate = strYear & ", Need Month/Day"

        strUTME = ""
        strUTMN = ""
        strComment = ""
        strElev = ""

        strFolder = "NA"
        strFileHeader = "NA"
      End If

      pProg.Step
      pProg.Message = strYear & ": Working on Quad " & strQuad & "...[" & CStr(lngIndex + 1) & " of " & strQuadCount & "]"
      DoEvents

      Debug.Print "Checking '" & strQuad & "' [" & strYear & "]"
      DoEvents

      strExportFilename = MyGeneralOperations.MakeUniquedBASEName(strSaveFolder & strFileHeader & "-" & _
          strPlot & "_" & strYear & ".tif")
      strExportFilename = strSaveFolder & strFileHeader & "-" & strPlot & "_" & strYear & ".tif"
      strExportFilename = Replace(strExportFilename, " / ", "_")

      If booOnlyDoSpecificQuadrats Then
        booExportThisOne = CheckIfShouldExport(strQuadrat, varSpecificQuadrats) And Not aml_func_mod.ExistFileDir(strExportFilename)
      Else
        booExportThisOne = Not aml_func_mod.ExistFileDir(strExportFilename)
      End If

      If booExportThisOne Then

        If aml_func_mod.ExistFileDir(strOrig & "\" & strQuad & "\" & strQuad & "_" & strYear & "_C.shp") Or _
            aml_func_mod.ExistFileDir(strOrig & "\" & strQuad & "\" & strQuad & "_" & strYear & "_D.shp") Then

          strSaveFolder = strExportFolder & strFolder & "\"
          strSaveFolder = Replace(strSaveFolder, " / ", "_")

          strExportFilename = strSaveFolder & strFileHeader & "-" & strPlot & "_" & strYear & ".tif"
          strExportFilename = Replace(strExportFilename, " / ", "_")
          If Not aml_func_mod.FileExists(strExportFilename) Then

            Set pAreaColl = New Collection
            Set pLegendColl = New Collection
            Erase strLegendKeys
            pLayersToDelete.RemoveAll

            Set pFLayer = Margaret.MakeGridFLayer
            pMxDoc.FocusMap.AddLayer pFLayer
            pLayersToDelete.Add pFLayer

            For lngWorkIndex = 0 To UBound(varWorkOrder)

              Erase strSpecies
              strWorkOption = varWorkOrder(lngWorkIndex)
              AddToAppropriateReport strWorkOption, strReport1, strReport2, "Checking '" & strQuad & "'" & vbCrLf
              Set pPolyColl = ReturnAppropriatePolyCollection(strWorkOption, strPointPrefix, strPointSuffix, pPointFClass, _
                    lngPointSpeciesIndex, strPrefix, strSuffix, pFClass, lngSpeciesIndex, pQueryFilt, _
                    booFoundSomething, strSpecies, strQuad, strYear, pLegendColl, strLegendKeys, _
                    pPointSymbols, pFillSymbols)

              If booFoundSomething Then
                FillAppropriateCountStats strWorkOption, lngCount, lngHighCount, strHighQuad, _
                    lngPointCount, lngPointHighCount, strPointHighQuad, strSpecies, strQuad

                For lngIndex2 = 0 To UBound(strSpecies)
                  strVal = strSpecies(lngIndex2)
                  varPolys = pPolyColl.Item(strVal)
                  dblCumulative = 0
                  For lngIndex3 = 0 To UBound(varPolys)
                    Set pPolygon = varPolys(lngIndex3)
                    Set pArea = pPolygon
                    dblArea = pArea.Area
                    dblCumulative = dblCumulative + dblArea
                  Next lngIndex3
                  strObsCount = Format(UBound(varPolys) + 1, "#,##0") & _
                      IIf(UBound(varPolys) = 0, " polygon", " polygons") & ", "
                  Debug.Print "  --> [" & strVal & "]: " & strObsCount & _
                      "Area = " & Format(dblCumulative * 10000, "#,##0.000") & " sq. cm. (" & _
                      Format(dblCumulative, "0.00%") & ")"

                  If MyGeneralOperations.CheckCollectionForKey(pAreaColl, strVal) Then
                    dblCurrentArea = pAreaColl.Item(strVal)
                    pAreaColl.Remove strVal
                  Else
                    dblCurrentArea = 0
                  End If
                  pAreaColl.Add dblCurrentArea + dblCumulative, strVal

                  AddToAppropriateReport strWorkOption, strReport1, strReport2, _
                        "  --> [" & strVal & "]: " & strObsCount & _
                        "Area = " & Format(dblCumulative * 10000, "#,##0.000") & " sq. cm. (" & _
                        Format(dblCumulative / 100, "0.00%") & ")" & vbCrLf

                  Set pNewFClass = ReturnInMemoryFClassFromPolys(varPolys, strVal, strWorkOption)
                  Set pSubLayer = ReturnSymbolizedFLayer(pNewFClass, strWorkOption, pSymbol, strVal, lngIndex2, _
                      pPointSymbols, pFillSymbols)
                  pLayersToDelete.Add pSubLayer
                  pMxDoc.FocusMap.AddLayer pSubLayer

                Next lngIndex2
              Else
                Debug.Print "  --> Found no " & strWorkOption & " Species..."
                AddToAppropriateReport strWorkOption, strReport1, strReport2, _
                    "  --> Found no " & strWorkOption & " Species..." & vbCrLf
                FillAppropriateCountStatsIfZero strWorkOption, lngHighCount, strHighQuad, lngPointHighCount, _
                    strPointHighQuad, lngCount, lngPointCount, strQuad
              End If
              AddToAppropriateReport strWorkOption, strReport1, strReport2, vbCrLf

            Next lngWorkIndex

            Call CreateCornerTab
            Call InsertHeaderInfo(strSiteName, strPlotName, strCrewName, strPhoto, strDate, strUTME, strUTMN, _
                strComment, strElev, strQuad)
            Debug.Print "Found " & CStr(pLegendColl.Count) & " total species for Quadrat " & strQuad & "."

            If pLegendColl.Count > lngMaxAllSpeciesCount Then
              lngMaxAllSpeciesCount = pLegendColl.Count
              strMaxAllSpeciesQuad = strQuad
            End If

              Call ConstructLegend(pLegendColl, strLegendKeys, pAreaColl, pMxDoc)

            pMxDoc.UpdateContents
            pMxDoc.ActiveView.Refresh

            If aml_func_mod.ExistFileDir(strOrig & "\" & strQuad & "\" & strQuad & "_" & strYear & "_C.shp") Or _
                aml_func_mod.ExistFileDir(strOrig & "\" & strQuad & "\" & strQuad & "_" & strYear & "_D.shp") Then

              strSaveFolder = strExportFolder & strFolder & "\"
              strSaveFolder = Replace(strSaveFolder, " / ", "_")
              MyGeneralOperations.CreateNestedFoldersByPath (strSaveFolder)

              strExportFilename = MyGeneralOperations.MakeUniquedBASEName(strSaveFolder & strFileHeader & "-" & _
                  strPlot & "_" & strYear & ".tif")
              strExportFilename = strSaveFolder & strFileHeader & "-" & strPlot & "_" & strYear & ".tif"
              strExportFilename = Replace(strExportFilename, " / ", "_")
              If Not aml_func_mod.ExistFileDir(strExportFilename) Then
                Map_Module.ExportActiveView strExportFilename, True, False
              End If
            End If

            For lngDeleteIndex = 0 To pLayersToDelete.Count - 1
              Set pFLayer = pLayersToDelete.Element(lngDeleteIndex)
              pMxDoc.FocusMap.DeleteLayer pFLayer
              Set pDataset = pFLayer.FeatureClass
              pDataset.DELETE
              Set pFLayer = Nothing
              Set pDataset = Nothing
              pMxDoc.UpdateContents
              pMxDoc.ActiveView.Refresh
            Next lngDeleteIndex
          End If
        End If
      End If
    Next lngIndex
  Next lngYearIndex2

  Debug.Print vbCrLf & "Most Species on " & strHighQuad & " [n = " & CStr(lngHighCount) & "]"
  Debug.Print vbCrLf & "Most Point Species on " & strPointHighQuad & " [n = " & CStr(lngPointHighCount) & "]"
  Debug.Print vbCrLf & "Most Combined Species on " & strMaxAllSpeciesQuad & " [n = " & CStr(lngMaxAllSpeciesCount) & "]"
  strReport2 = strReport2 & vbCrLf & "-----------------------" & vbCrLf & _
      "Most Cover (Polygon) Species on " & strHighQuad & " [n = " & CStr(lngHighCount) & "]" & vbCrLf & _
      "Most Density (Point) Species on " & strPointHighQuad & " [n = " & CStr(lngPointHighCount) & "]" & vbCrLf & _
      "Most Combined Species on " & strMaxAllSpeciesQuad & " [n = " & CStr(lngMaxAllSpeciesCount) & "]"

  For lngIndex = 0 To pLayersToDelete.Count - 1
    Set pFLayer = pLayersToDelete.Element(lngIndex)
    pMxDoc.FocusMap.DeleteLayer pFLayer
    Set pDataset = pFLayer.FeatureClass
    pDataset.DELETE
    Set pFLayer = Nothing
    Set pDataset = Nothing
  Next lngIndex

  pMxDoc.UpdateContents
  pMxDoc.ActiveView.Refresh
  pSBar.HideProgressBar
  pProg.position = 0

  pDataObj.Clear
  pDataObj.SetText strReport1 & vbCrLf & "--------------------------" & vbCrLf & strReport2

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(Abs(GetTickCount) - Abs(lngStart))

ClearMemory:
  Set pEnv = Nothing
  Set pSourceWS = Nothing
  Set pSourceWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pWS2 = Nothing
  Set pSymTable = Nothing
  Set pSymRow = Nothing
  Set pSymbolColl = Nothing
  Set pSymCursor = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pFields = Nothing
  Erase strItem
  Set pLocationsAndNotes = Nothing
  Erase strPlotNames
  Erase varLocNotes
  Erase strFinalQuadrats
  Set pQuadratNumColl = Nothing
  Erase strItems
  Erase strYears
  Erase strQuads
  Erase strCheck
  Set pGeoDataset = Nothing
  Set pLayersToDelete = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pFLayer = Nothing
  Set pFeatDef = Nothing
  Set pQueryFilt = Nothing
  Set pYearColl = Nothing
  Set pQuadColl = Nothing
  varVal = Null
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pPointFClass = Nothing
  Set pFClass = Nothing
  varQuads = Null
  varYears = Null
  Set pPointSymbols = Nothing
  Set pFillSymbols = Nothing
  Set pSymBuffer = Nothing
  Set pPolyColl = Nothing
  Erase strSpecies
  Set pPolygon = Nothing
  Erase varPolys
  Set pArea = Nothing
  Set pDataObj = Nothing
  Set pNewFClass = Nothing
  Erase varWorkOrder
  Set pDataset = Nothing
  Set pSubLayer = Nothing
  Set pSymbol = Nothing
  Set pAreaColl = Nothing
  Set pLegendColl = Nothing
  Erase strLegendKeys

End Sub

Public Sub ClearAnyInitialData()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pLayer As ILayer
  Dim pLayers As esriSystem.IArray
  Set pLayers = New esriSystem.Array
  Dim lngIndex As Long
  For lngIndex = 0 To pMxDoc.FocusMap.LayerCount - 1
    Set pLayer = pMxDoc.FocusMap.Layer(lngIndex)
    If pLayer.Name <> "Box" And pLayer.Name <> "Sample Grid" Then
      pLayers.Add pLayer
    End If
  Next lngIndex

  If pLayers.Count > 0 Then
    For lngIndex = 0 To pLayers.Count - 1
      Set pLayer = pLayers.Element(lngIndex)
      pMxDoc.FocusMap.DeleteLayer pLayer
    Next lngIndex
  End If

  pMxDoc.UpdateContents

  Debug.Print "Done clearing layers..."

End Sub

Public Function CheckIfShouldExport(strQuadrat As String, varSpecific() As Variant) As Boolean

  CheckIfShouldExport = False
  Dim lngIndex As Long
  For lngIndex = 0 To UBound(varSpecific)
    If strQuadrat = CStr(varSpecific(lngIndex)) Then
      CheckIfShouldExport = True
      Exit For
    End If
  Next lngIndex

End Function

Public Sub MakePageNumbers(Optional pPageNumberByPlotDate As Collection)

  Dim pFiles As esriSystem.IStringArray

  Dim strFinalQuadrats() As String
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim lngFeatCount As Long
  Dim pQuadratNumColl As Collection
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadratNumColl = Margaret_Analysis_Functions.FillQuadratNameColl_Rev(strFinalQuadrats, pPlotToQuadratConversion, pQuadratToPlotConversion, _
      varSites, varSitesSpecific)

  Dim strItems() As String
  Dim strNote As String

  Dim strExportBase As String
  Dim strModifiedRoot As String
  Dim strOrig As String
  Call DeclareWorkspaces(strOrig, strModifiedRoot, , strExportBase)

  Set pFiles = ReturnFilesFromNestedFolders2(strExportBase, ".tif")

  Debug.Print pFiles.Count
  Dim strFileNames() As String
  ReDim strFileNames(pFiles.Count - 1)

  Dim strModName As String
  Dim strSplit() As String
  Dim strPlot As String
  Dim strQuad As String
  Dim lngPlot As Long
  Dim lngQuad As Long
  Dim strSiteSpecific As String
  Dim strSiteShort As String
  Dim lngYear As Long

  Dim varSortBySite() As Variant
  ReDim varSortBySite(3, pFiles.Count - 1)
  Dim varSortByQuad() As Variant
  ReDim varSortByQuad(2, pFiles.Count - 1)

  Dim lngIndex As Long
  Dim strPath As String
  Dim strJustName As String
  For lngIndex = 0 To pFiles.Count - 1
    strPath = pFiles.Element(lngIndex)
    strJustName = aml_func_mod.ReturnFilename2(strPath)
    strJustName = aml_func_mod.ClipExtension2(strJustName)

    strModName = Replace(strJustName, "-", "|", , , vbTextCompare)
    strModName = Replace(strModName, "_", "|", , , vbTextCompare)
    strSplit = Split(strModName, "|")
    strSiteShort = strSplit(0)
    lngYear = CLng(strSplit(UBound(strSplit)))

    If UBound(strSplit) > 2 Then
      Debug.Print MyGeneralOperations.SpacesInFrontOfText(Format(lngIndex, "#,##0"), 5) & "] " & strJustName
      If IsNumeric(strSplit(1)) Then
        strPlot = strSplit(1)
      Else
        strPlot = strSplit(2)
      End If
    Else
      strPlot = strSplit(1)
    End If

    Select Case strPlot
      Case "8"
        strQuad = CStr(pPlotToQuadratConversion.Item("30708"))
      Case "10"
        strQuad = CStr(pPlotToQuadratConversion.Item("30710"))
      Case "16"
        strQuad = CStr(pPlotToQuadratConversion.Item("30716"))
      Case "18"
        strQuad = CStr(pPlotToQuadratConversion.Item("30718"))
      Case Else
        strQuad = CStr(pPlotToQuadratConversion.Item(strPlot))
    End Select
    lngPlot = CLng(strPlot)
    lngQuad = CLng(Replace(strQuad, "Q", ""))
    strItems = pQuadratNumColl.Item(Format(lngQuad, "0"))
    strSiteSpecific = strItems(1)

    varSortBySite(0, lngIndex) = strSiteShort
    varSortBySite(1, lngIndex) = lngPlot
    varSortBySite(2, lngIndex) = lngYear
    varSortBySite(3, lngIndex) = strJustName

    varSortByQuad(0, lngIndex) = lngQuad
    varSortByQuad(1, lngIndex) = lngYear
    varSortByQuad(2, lngIndex) = strJustName

    strFileNames(lngIndex) = strJustName
  Next lngIndex

  Dim varTypes() As Variant
  ReDim varTypes(3)
  varTypes(0) = enum_TypeString
  varTypes(1) = enum_TypeLong
  varTypes(2) = enum_TypeLong
  varTypes(3) = enum_TypeString
  QuickSort.MultiSort varSortBySite, varTypes, vbTextCompare

  ReDim varTypes(2)
  varTypes(0) = enum_TypeLong
  varTypes(1) = enum_TypeLong
  varTypes(2) = enum_TypeString
  QuickSort.MultiSort varSortByQuad, varTypes, vbTextCompare

  Dim strReport As String
  Dim pPageByNameColl As New Collection
  Set pPageNumberByPlotDate = New Collection
  For lngIndex = 0 To UBound(strFileNames)
    strJustName = varSortBySite(3, lngIndex)
    pPageByNameColl.Add lngIndex + 1, strJustName
    strReport = strReport & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf

    strPlot = "Plot_" & Format(varSortBySite(1, lngIndex), "0")
    lngYear = varSortBySite(2, lngIndex)
    pPageNumberByPlotDate.Add lngIndex + 1, strPlot & "_" & Format(lngYear, "0")
  Next lngIndex

  Dim str2019_2024Report As String
  Dim p2018_2019PageByNameColl As New Collection
  For lngIndex = 0 To UBound(strFileNames)
    strJustName = varSortBySite(3, lngIndex)
    p2018_2019PageByNameColl.Add lngIndex + 1, strJustName

    If InStr(1, strJustName, "_2019", vbTextCompare) > 0 Then
      str2019_2024Report = str2019_2024Report & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf
    ElseIf InStr(1, strJustName, "_2020", vbTextCompare) > 0 Then
      str2019_2024Report = str2019_2024Report & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf
    ElseIf InStr(1, strJustName, "_2021", vbTextCompare) > 0 Then
      str2019_2024Report = str2019_2024Report & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf
    ElseIf InStr(1, strJustName, "_2022", vbTextCompare) > 0 Then
      str2019_2024Report = str2019_2024Report & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf & vbCrLf
    ElseIf InStr(1, strJustName, "_2023", vbTextCompare) > 0 Then
      str2019_2024Report = str2019_2024Report & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf & vbCrLf
    ElseIf InStr(1, strJustName, "_2024", vbTextCompare) > 0 Then
      str2019_2024Report = str2019_2024Report & strJustName & vbTab & "p. " & Format(lngIndex + 1 + 18, "#,##0") & vbCrLf & vbCrLf
    End If
  Next lngIndex

  Dim strSortByQuadReport As String
  Dim lngPageNum As Long
  For lngIndex = 0 To UBound(strFileNames)
    strJustName = varSortByQuad(2, lngIndex)
    lngPageNum = pPageByNameColl.Item(strJustName)

    strSortByQuadReport = strSortByQuadReport & strJustName & vbTab & "p. " & Format(lngPageNum, "#,##0") & vbCrLf
  Next lngIndex

  Dim pDataObj As New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReport
  pDataObj.PutInClipboard

  Debug.Print "Done..."

ClearMemory:
  Set pFiles = Nothing
  Erase strFinalQuadrats
  Set pQuadratNumColl = Nothing
  Erase strItems
  Erase strFileNames
  Erase strSplit
  Set pDataObj = Nothing

End Sub

Public Sub CreateCollectionOfSymbolsFromTextFile(strTextFile As String, pPointSymbols As Collection, _
      pFillSymbols As Collection, pSBar As IStatusBar, pProg As IStepProgressor)

  Dim strVal As String
  Dim lngCount As Long
  Dim lngCounter As Long
  Dim pDoneColl As Collection
  Dim strSpecies As String
  Dim pFillColor As IRgbColor
  Dim pOutlineColor As IRgbColor
  Dim pMarkerSymbol As IMarkerSymbol
  Dim pPolySymbol As IFillSymbol
  Dim strMarkerText As String
  Dim strMarkerLine As String
  Dim lngIndex As Long
  Dim strLine As String
  Dim strSymbologyText As String
  Dim strLines() As String
  Dim strLineSplit() As String

  Set pPointSymbols = New Collection
  Set pFillSymbols = New Collection

  strLines = Split(strTextFile, vbCrLf)
  lngCount = UBound(strLines)
  lngCounter = 0
  pSBar.ShowProgressBar "Making initial symbols...", 0, lngCount, 1, True
  pProg.position = 0

  For lngIndex = 0 To UBound(strLines)
    pProg.Step
    If lngIndex Mod 100 = 0 Then
      DoEvents
    End If
    strLine = Trim(strLines(lngIndex))
    If strLine <> "" Then
      strLineSplit = Split(strLine, "^%$#@")
      strSpecies = strLineSplit(1)
      strSymbologyText = strLineSplit(2)

      If strLineSplit(0) = "Marker" Then
        If Not MyGeneralOperations.CheckCollectionForKey(pPointSymbols, strSpecies) Then
          Set pMarkerSymbol = ReturnMarkerFromLine(strSymbologyText)
          pPointSymbols.Add pMarkerSymbol, strSpecies
        End If
      ElseIf strLineSplit(0) = "Fill" Then
        If Not MyGeneralOperations.CheckCollectionForKey(pFillSymbols, strSpecies) Then
          Set pPolySymbol = ReturnFillFromLine(strSymbologyText)
          pFillSymbols.Add pPolySymbol, strSpecies
        End If
      Else
        MsgBox "Unexpected Symbol Line!" & vbCrLf & "  --> " & strLine
      End If
    End If
  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

ClearMemory:
  Set pDoneColl = Nothing
  Set pFillColor = Nothing
  Set pOutlineColor = Nothing
  Set pMarkerSymbol = Nothing
  Set pPolySymbol = Nothing
  Erase strLines
  Erase strLineSplit

End Sub

Public Function ReturnSymbolizedFLayer(pFClass As IFeatureClass, strWorkOrderOption As String, _
    pSymbol As ISymbol, strSpeciesName As String, lngIndex As Long, pPointSymbols As Collection, _
    pFillSymbols As Collection) As IFeatureLayer

  Dim pLyr As IGeoFeatureLayer
  Dim pRender As ISimpleRenderer
  Dim pPolySymbol As IFillSymbol
  Dim pNewFlayer As IFeatureLayer
  Dim hx As IRendererPropertyPage
  Dim pLayerEffects As ILayerEffects
  Dim lngTransparency As Long

  lngTransparency = 20

  Dim pFillColor As IRgbColor
  Dim pOutlineColor As IRgbColor
  Set pFillColor = New RgbColor
  Set pOutlineColor = New RgbColor
  Dim pMarkerSymbol As IMarkerSymbol
  Randomize

  If strWorkOrderOption = "Density" Then
    Set pMarkerSymbol = pPointSymbols.Item(strSpeciesName)

  ElseIf strWorkOrderOption = "Cover" Then
    Set pPolySymbol = pFillSymbols.Item(strSpeciesName)
  End If

  Dim pClone As IClone

  Set pNewFlayer = New FeatureLayer
  Set pNewFlayer.FeatureClass = pFClass
  pNewFlayer.Name = strSpeciesName
  Set pLyr = pNewFlayer
  Set pRender = New SimpleRenderer
  If strWorkOrderOption = "Density" Then
    Set pRender.Symbol = pMarkerSymbol
    Set pClone = pMarkerSymbol
  Else
    Set pClone = pPolySymbol
    Set pRender.Symbol = pPolySymbol
  End If
  pRender.Label = strSpeciesName
  Set pLyr.Renderer = pRender
  Set hx = New SingleSymbolPropertyPage
  pLyr.RendererPropertyPageClassID = hx.ClassID
  Set pLayerEffects = pNewFlayer
  pLayerEffects.Transparency = lngTransparency

  Set pSymbol = pClone.Clone

  Set ReturnSymbolizedFLayer = pNewFlayer

ClearMemory:
  Set pClone = Nothing
  Set pLyr = Nothing
  Set pRender = Nothing
  Set pPolySymbol = Nothing
  Set pFillColor = Nothing
  Set pOutlineColor = Nothing
  Set hx = Nothing
  Set pLayerEffects = Nothing

End Function

Public Sub FillAppropriateCountStatsIfZero(strWorkOrderOption As String, lngHighCount As Long, strHighQuad As String, _
    lngPointHighCount As Long, strPointHighQuad As String, lngCount As Long, lngPointCount As Long, strQuad As String)

  If strWorkOrderOption = "Cover" Then
    lngCount = 0
    If lngCount > lngHighCount Then
      lngHighCount = lngCount
      strHighQuad = strQuad
    End If
  ElseIf strWorkOrderOption = "Density" Then
    lngPointCount = 0
    If lngPointCount > lngPointHighCount Then
      lngPointHighCount = lngPointCount
      strPointHighQuad = strQuad
    End If
  End If

End Sub

Public Sub FillAppropriateCountStats(strWorkOrderOption As String, lngCount As Long, lngHighCount As Long, _
    strHighQuad As String, lngPointCount As Long, lngPointHighCount As Long, strPointHighQuad As String, _
    strSpecies() As String, strQuad As String)

  If strWorkOrderOption = "Cover" Then
    lngCount = UBound(strSpecies) + 1
    If lngCount > lngHighCount Then
      lngHighCount = lngCount
      strHighQuad = strQuad
    End If
  ElseIf strWorkOrderOption = "Density" Then
    lngPointCount = UBound(strSpecies) + 1
    If lngPointCount > lngPointHighCount Then
      lngPointHighCount = lngPointCount
      strPointHighQuad = strQuad
    End If
  End If

End Sub

Public Function ReturnAppropriatePolyCollection(strWorkOrderOption As String, strPointPrefix As String, _
    strPointSuffix As String, pPointFClass As IFeatureClass, lngPointSpeciesIndex As Long, _
    strPrefix As String, strSuffix As String, pFClass As IFeatureClass, lngSpeciesIndex As Long, _
    pQueryFilt As IQueryFilter, booFoundSomething As Boolean, strSpecies() As String, _
    strQuad As String, strYear As String, pLegendColl As Collection, _
    strLegendKeys() As String, pPointSymbols As Collection, pFillSymbols As Collection) As Collection

  Dim pPolyColl As New Collection

  If strWorkOrderOption = "Cover" Then
    pQueryFilt.WhereClause = strPrefix & "Quadrat" & strSuffix & " = '" & strQuad & "' AND " & _
        strPrefix & "Year" & strSuffix & " = '" & strYear & "'"
    Set pPolyColl = ReturnUniqueSpecies(pFClass, pQueryFilt, lngSpeciesIndex, booFoundSomething, strSpecies)

  ElseIf strWorkOrderOption = "Density" Then
    pQueryFilt.WhereClause = strPointPrefix & "Quadrat" & strPointSuffix & " = '" & strQuad & "' AND " & _
        strPointPrefix & "Year" & strPointSuffix & " = '" & strYear & "'"
    Set pPolyColl = ReturnUniqueSpecies(pPointFClass, pQueryFilt, lngPointSpeciesIndex, booFoundSomething, strSpecies)
  End If

  Call FillLegendCollection(strWorkOrderOption, strSpecies, _
      pLegendColl, strLegendKeys, pPointSymbols, pFillSymbols)

  Set ReturnAppropriatePolyCollection = pPolyColl
  Set pPolyColl = Nothing

End Function

Public Sub FillLegendCollection(strWorkOrderOption As String, strSpecies() As String, _
    pLegendColl As Collection, strLegendKeys() As String, _
    pPointSymbols As Collection, pFillSymbols As Collection)

  Dim lngIndex As Long
  Dim varPair() As Variant
  Dim strKey As String

  If MyGeneralOperations.IsDimmed(strSpecies) Then
    For lngIndex = 0 To UBound(strSpecies)
      strKey = strSpecies(lngIndex)
      If Not MyGeneralOperations.CheckCollectionForKey(pLegendColl, strKey) Then
        ReDim varPair(1)
        varPair(0) = Null
        varPair(1) = Null
        If strWorkOrderOption = "Cover" Then
          Set varPair(0) = pFillSymbols.Item(strKey)
        ElseIf strWorkOrderOption = "Density" Then
          Set varPair(1) = pPointSymbols.Item(strKey)
        End If
        pLegendColl.Add varPair, strKey
        If MyGeneralOperations.IsDimmed(strLegendKeys) Then
          ReDim Preserve strLegendKeys(UBound(strLegendKeys) + 1)
          strLegendKeys(UBound(strLegendKeys)) = strKey
        Else
          ReDim strLegendKeys(0)
          strLegendKeys(0) = strKey
        End If

      Else ' ALREADY HAVE RECORD OF THIS SPECIES
        varPair = pLegendColl.Item(strKey)
        pLegendColl.Remove strKey
        If strWorkOrderOption = "Cover" Then
          Set varPair(0) = pFillSymbols.Item(strKey)
        ElseIf strWorkOrderOption = "Density" Then
          Set varPair(1) = pPointSymbols.Item(strKey)
        End If
        pLegendColl.Add varPair, strKey
      End If
    Next lngIndex
  End If

ClearMemory:
  Erase varPair

End Sub

Public Sub AddToAppropriateReport(strWorkOrderOption As String, strReport1 As String, strReport2 As String, _
    strTextToAdd As String)

  If strWorkOrderOption = "Cover" Then
    strReport1 = strReport1 & strTextToAdd
  ElseIf strWorkOrderOption = "Density" Then
    strReport2 = strReport2 & strTextToAdd
  End If

End Sub

Public Function ReturnInMemoryFClassFromPolys(varPolys() As Variant, strSpecies As String, _
    strWorkOption As String) As IFeatureClass

  Dim pNewFClass As IFeatureClass
  Dim pPolys As esriSystem.IArray
  Set pPolys = New esriSystem.Array

  Dim pPolygon As IPolygon
  Dim pArea As IArea
  Dim pEnv As IEnvelope
  Dim pPoint As IPoint

  Dim pFields As esriSystem.IVariantArray
  Dim pVals As esriSystem.IVariantArray

  Set pVals = New esriSystem.varArray
  Set pFields = New esriSystem.varArray

  Dim pField As IField
  Dim pFieldEdit As IFieldEdit
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Species"
    .Type = esriFieldTypeString
    .length = Len(strSpecies)
  End With
  pFields.Add pField

  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Sq_Meters"
    .Type = esriFieldTypeDouble
  End With
  pFields.Add pField

  Dim pSubArray As esriSystem.IVariantArray

  Dim lngIndex As Long
  For lngIndex = 0 To UBound(varPolys)
    Set pPolygon = varPolys(lngIndex)
    Set pArea = pPolygon

    Set pSubArray = New esriSystem.varArray
    pSubArray.Add strSpecies
    pSubArray.Add pArea.Area
    pVals.Add pSubArray

    If strWorkOption = "Density" Then
      Set pEnv = pPolygon.Envelope
      Set pPoint = New Point
      Set pPoint.SpatialReference = pPolygon.SpatialReference
      pPoint.PutCoords pEnv.XMin + (pEnv.Width / 2), pEnv.YMin + (pEnv.Height / 2)
      pPolys.Add pPoint
    Else
      pPolys.Add pPolygon
    End If
  Next lngIndex

  Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pPolys, pVals, pFields)
  Set ReturnInMemoryFClassFromPolys = pNewFClass

ClearMemory:
  Set pNewFClass = Nothing
  Set pPolys = Nothing
  Set pPolygon = Nothing
  Set pArea = Nothing
  Set pFields = Nothing
  Set pVals = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSubArray = Nothing

End Function

Public Function ReturnUniqueSpecies(pFClass As IFeatureClass, pQueryFilt As IQueryFilter, _
    lngSpeciesIndex As Long, booFoundSomething As Boolean, strKeys() As String) As Collection

  Dim lngIndex As Long
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strSpecies As String
  Dim pPolygon As IPolygon
  Dim pColl As New Collection
  Dim varPolys() As Variant
  Dim lngPolyIndex As Long
  Dim varReturn() As Variant

  lngIndex = -1
  Set pFCursor = pFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strSpecies = pFeature.Value(lngSpeciesIndex)
    Set pPolygon = pFeature.ShapeCopy
    If Not pPolygon.IsEmpty Then
      If StrComp(strSpecies, "No Point Species", vbTextCompare) <> 0 And _
         StrComp(strSpecies, "No Polygon Species", vbTextCompare) <> 0 And _
         StrComp(strSpecies, "No Cover Species Observed", vbTextCompare) <> 0 And _
         StrComp(strSpecies, "No Density Species Observed", vbTextCompare) <> 0 Then
        If Not MyGeneralOperations.CheckCollectionForKey(pColl, strSpecies) Then
          ReDim varPolys(0)
          Set varPolys(0) = pPolygon
          pColl.Add varPolys, strSpecies
          lngIndex = lngIndex + 1
          ReDim Preserve strKeys(lngIndex)
          strKeys(lngIndex) = strSpecies
        Else
          varPolys = pColl.Item(strSpecies)
          pColl.Remove strSpecies
          lngPolyIndex = UBound(varPolys) + 1
          ReDim Preserve varPolys(lngPolyIndex)
          Set varPolys(lngPolyIndex) = pPolygon
          pColl.Add varPolys, strSpecies
        End If
      End If
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  If lngIndex > 0 Then
    QuickSort.StringsAscending strKeys, 0, UBound(strKeys)
  End If

  booFoundSomething = lngIndex > -1

  Set ReturnUniqueSpecies = pColl

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pPolygon = Nothing
  Set pColl = Nothing
  Erase varPolys
  Erase varReturn

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

Public Function MakeGridFLayer() As IFeatureLayer

  Dim pFClass As IFeatureClass
  Dim pFLayer As IFeatureLayer
  Dim pFields As esriSystem.IVariantArray
  Dim pField As IField
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

Public Function ReturnFillFromLine(strSymbology As String) As IFillSymbol

  Dim strType As String
  Dim strObjects() As String
  Dim lngOutlineColor As Long
  Dim lngFillColor As Long

  strObjects = Split(strSymbology, vbTab)
  strType = strObjects(0)
  lngOutlineColor = CLng(strObjects(2))
  lngFillColor = CLng(strObjects(1))

  Dim pFillColor As IRgbColor
  Set pFillColor = New RgbColor
  pFillColor.RGB = lngFillColor

  Dim pOutlineColor As IRgbColor
  Set pOutlineColor = New RgbColor
  pOutlineColor.RGB = lngOutlineColor

  Dim pCartoline As ICartographicLineSymbol
  Dim lngCartoCap As Long
  Dim lngCartoJoin As Long
  Dim lngCartoWidth As Long

  Dim pMultiSymbol As IMultiLayerFillSymbol

  Dim lngAngle1 As Long
  Dim lngAngle2 As Long
  Dim lngAngle3 As Long
  Dim dblSeparation1 As Double
  Dim dblSeparation2 As Double
  Dim dblSeparation3 As Double
  Dim dblOffset1 As Double
  Dim dblOffset2 As Double
  Dim dblOffset3 As Double
  Dim pLineFill As ILineFillSymbol
  Dim pLineFill2 As ILineFillSymbol
  Dim pLineFill3 As ILineFillSymbol
  Dim dblOutlineWidth As Double
  Dim lngOutlineStyle As Long
  Dim pOutlineLineSymbol As ISimpleLineSymbol
  Dim pInnerColor As IRgbColor
  Dim lngInnerColor As Long
  Dim pInnerSymbol As ISimpleFillSymbol
  Dim pInnerLine As ISimpleLineSymbol
  Dim pSimpleFillSymbol As ISimpleFillSymbol
  Dim pOutline As ISimpleLineSymbol

  If strType = "1" Then

    dblOutlineWidth = CDbl(strObjects(3))
    lngOutlineStyle = CLng(strObjects(4))

    Set pSimpleFillSymbol = New SimpleFillSymbol

    Set pOutline = New SimpleLineSymbol
    pOutline.Color = pOutlineColor
    pOutline.Width = dblOutlineWidth
    pOutline.Style = lngOutlineStyle

    pSimpleFillSymbol.Color = pFillColor
    pSimpleFillSymbol.Outline = pOutline

    Set ReturnFillFromLine = pSimpleFillSymbol

  ElseIf strType = "2" Then
    lngCartoCap = CLng(strObjects(3))
    lngCartoJoin = CLng(strObjects(4))
    lngCartoWidth = CLng(strObjects(5))
    lngAngle1 = CLng(strObjects(6))
    lngAngle2 = CLng(strObjects(7))
    dblSeparation1 = CDbl(strObjects(8))
    dblSeparation2 = CDbl(strObjects(9))
    dblOffset1 = CDbl(strObjects(10))
    dblOffset2 = CDbl(strObjects(11))
    dblOutlineWidth = CDbl(strObjects(12))
    lngOutlineStyle = CDbl(strObjects(13))
    lngInnerColor = CDbl(strObjects(14))

    Set pInnerColor = New RgbColor
    pInnerColor.RGB = lngInnerColor

    Set pCartoline = New CartographicLineSymbol
    With pCartoline
      .Cap = lngCartoCap
      .Join = lngCartoJoin
      .Color = pFillColor
      .Width = lngCartoWidth
    End With

    Set pMultiSymbol = New MultiLayerFillSymbol

    Set pInnerSymbol = New SimpleFillSymbol
    Set pInnerLine = New SimpleLineSymbol
    pInnerLine.Style = esriSLSNull
    pInnerSymbol.Color = pInnerColor
    pInnerSymbol.Style = esriSFSSolid
    pInnerSymbol.Outline = pInnerLine
    pMultiSymbol.AddLayer pInnerSymbol

    Set pLineFill = New LineFillSymbol
    With pLineFill
      .Angle = lngAngle1
      .Separation = dblSeparation1
      .Offset = dblOffset1
    End With
    Set pLineFill.LineSymbol = pCartoline
    pMultiSymbol.AddLayer pLineFill

    Set pLineFill2 = New LineFillSymbol
    With pLineFill2
      .Angle = lngAngle2
      .Separation = dblSeparation2
      .Offset = dblOffset2
    End With
    Set pLineFill2.LineSymbol = pCartoline
    pMultiSymbol.AddLayer pLineFill2

    Set pOutlineLineSymbol = New SimpleLineSymbol
    pOutlineLineSymbol.Color = pOutlineColor
    pOutlineLineSymbol.Width = dblOutlineWidth
    pOutlineLineSymbol.Style = lngOutlineStyle

    pMultiSymbol.Outline = pOutlineLineSymbol

    Set ReturnFillFromLine = pMultiSymbol

  ElseIf strType = "3" Then
    lngCartoCap = CLng(strObjects(3))
    lngCartoJoin = CLng(strObjects(4))
    lngCartoWidth = CLng(strObjects(5))
    lngAngle1 = CLng(strObjects(6))
    lngAngle2 = CLng(strObjects(7))
    lngAngle3 = CLng(strObjects(8))
    dblSeparation1 = CDbl(strObjects(9))
    dblSeparation2 = CDbl(strObjects(10))
    dblSeparation3 = CDbl(strObjects(11))
    dblOffset1 = CDbl(strObjects(12))
    dblOffset2 = CDbl(strObjects(13))
    dblOffset3 = CDbl(strObjects(14))
    dblOutlineWidth = CDbl(strObjects(15))
    lngOutlineStyle = CDbl(strObjects(16))

    Set pCartoline = New CartographicLineSymbol
    With pCartoline
      .Cap = lngCartoCap
      .Join = lngCartoJoin
      .Color = pFillColor
      .Width = lngCartoWidth
    End With

    Set pMultiSymbol = New MultiLayerFillSymbol

    Set pLineFill = New LineFillSymbol
    With pLineFill
      .Angle = lngAngle1
      .Separation = dblSeparation1
      .Offset = dblOffset1
    End With
    Set pLineFill.LineSymbol = pCartoline
    pMultiSymbol.AddLayer pLineFill

    Set pLineFill2 = New LineFillSymbol
    With pLineFill2
      .Angle = lngAngle2
      .Separation = dblSeparation2
      .Offset = dblOffset2
    End With
    Set pLineFill2.LineSymbol = pCartoline
    pMultiSymbol.AddLayer pLineFill2

    Set pLineFill3 = New LineFillSymbol
    With pLineFill3
      .Angle = lngAngle3
      .Separation = dblSeparation3
      .Offset = dblOffset3
    End With
    Set pLineFill3.LineSymbol = pCartoline
    pMultiSymbol.AddLayer pLineFill3

    Set pOutlineLineSymbol = New SimpleLineSymbol
    pOutlineLineSymbol.Color = pOutlineColor
    pOutlineLineSymbol.Width = dblOutlineWidth
    pOutlineLineSymbol.Style = lngOutlineStyle

    pMultiSymbol.Outline = pOutlineLineSymbol

    Set ReturnFillFromLine = pMultiSymbol
  End If

ClearMemory:
  Erase strObjects
  Set pFillColor = Nothing
  Set pOutlineColor = Nothing
  Set pCartoline = Nothing
  Set pMultiSymbol = Nothing
  Set pLineFill = Nothing
  Set pLineFill2 = Nothing
  Set pLineFill3 = Nothing
  Set pOutlineLineSymbol = Nothing
  Set pInnerColor = Nothing
  Set pInnerSymbol = Nothing
  Set pInnerLine = Nothing
  Set pSimpleFillSymbol = Nothing
  Set pOutline = Nothing

End Function

Public Function MakeSimpleFillSymbol(pFillColor As IRgbColor, pOutlineColor As IRgbColor, _
    lngOutlineWidth As Long, strFillLine As String) As IFillSymbol

  Dim pSimpleFillSymbol As ISimpleFillSymbol
  Set pSimpleFillSymbol = New SimpleFillSymbol

  Dim pOutline As ISimpleLineSymbol
  Set pOutline = New SimpleLineSymbol
  pOutline.Color = pOutlineColor
  pOutline.Width = lngOutlineWidth
  pOutline.Style = esriSLSSolid

  pSimpleFillSymbol.Color = pFillColor
  pSimpleFillSymbol.Outline = pOutline

  Set MakeSimpleFillSymbol = pSimpleFillSymbol

  strFillLine = "1" & vbTab & Format(pFillColor.RGB, "0") & vbTab & Format(pOutlineColor.RGB, "0") & vbTab
  strFillLine = strFillLine & Format(pOutline.Width, "0.000") & vbTab & Format(pOutline.Style, "0")

ClearMemory:
  Set pSimpleFillSymbol = Nothing
  Set pOutline = Nothing

End Function

Public Function MakeCrosshatchSymbol(pColor As IRgbColor, pOutlineColor As IRgbColor, strFillLine As String) As IFillSymbol

  Dim pInnerColor As IRgbColor
  Set pInnerColor = New RgbColor
  pInnerColor.Red = pColor.Red + ((255 - pColor.Red) * 0.95)
  pInnerColor.Green = pColor.Green + ((255 - pColor.Green) * 0.95)
  pInnerColor.Blue = pColor.Blue + ((255 - pColor.Blue) * 0.95)

  Randomize
  Dim pTemp As IRgbColor
  Set pTemp = New RgbColor
  If Rnd() > 0.5 Then
    pTemp.RGB = pColor.RGB
    pColor.RGB = pInnerColor.RGB
    pInnerColor.RGB = pTemp.RGB
  End If

  Dim pCartoline As ICartographicLineSymbol
  Set pCartoline = New CartographicLineSymbol
  With pCartoline
    .Cap = esriLCSButt
    .Join = esriLJSBevel
    .Color = pColor
    .Width = 1
  End With

  Dim pMultiSymbol As IMultiLayerFillSymbol
  Set pMultiSymbol = New MultiLayerFillSymbol

  Dim pInnerSymbol As ISimpleFillSymbol
  Set pInnerSymbol = New SimpleFillSymbol
  Dim pInnerLine As ISimpleLineSymbol
  Set pInnerLine = New SimpleLineSymbol
  pInnerLine.Style = esriSLSNull
  pInnerSymbol.Color = pInnerColor
  pInnerSymbol.Style = esriSFSSolid
  pInnerSymbol.Outline = pInnerLine

  pMultiSymbol.AddLayer pInnerSymbol

  Dim lngAngle1 As Long
  Dim lngAngle2 As Long
  Randomize
  lngAngle1 = Rnd() * 180
  lngAngle2 = lngAngle1 + 90

  Dim pLineFill As ILineFillSymbol
  Set pLineFill = New LineFillSymbol
  With pLineFill
    .Angle = lngAngle1
    .Separation = 5
    .Offset = 1
  End With
  Set pLineFill.LineSymbol = pCartoline
  pMultiSymbol.AddLayer pLineFill

  Dim pLineFill2 As ILineFillSymbol
  Set pLineFill2 = New LineFillSymbol
  With pLineFill2
    .Angle = lngAngle2
    .Separation = 5
    .Offset = 3
  End With
  Set pLineFill2.LineSymbol = pCartoline
  pMultiSymbol.AddLayer pLineFill2

  Dim pOutlineLineSymbol As ISimpleLineSymbol
  Set pOutlineLineSymbol = New SimpleLineSymbol

  pOutlineLineSymbol.Color = pOutlineColor
  pOutlineLineSymbol.Width = 1.5
  pOutlineLineSymbol.Style = esriSLSSolid

  pMultiSymbol.Outline = pOutlineLineSymbol

  Set MakeCrosshatchSymbol = pMultiSymbol

  strFillLine = "2" & vbTab & Format(pColor.RGB, "0") & vbTab & Format(pOutlineColor.RGB, "0") & vbTab
  strFillLine = strFillLine & Format(pCartoline.Cap, "0") & vbTab & Format(pCartoline.Join, "0") & vbTab & _
      Format(pCartoline.Width, "0.000") & vbTab & Format(lngAngle1, "0") & vbTab & Format(lngAngle2, "0") & vbTab & _
      Format(pLineFill.Separation, "0.000") & vbTab & Format(pLineFill2.Separation, "0.000") & vbTab & _
      Format(pLineFill.Offset, "0.000") & vbTab & Format(pLineFill2.Offset, "0.000") & vbTab & _
      Format(pOutlineLineSymbol.Width, "0.000") & vbTab & _
      Format(pOutlineLineSymbol.Style, "0") & vbTab & Format(pInnerColor.RGB, "0")

  GoTo ClearMemory

ClearMemory:
  Set pCartoline = Nothing
  Set pInnerColor = Nothing
  Set pTemp = Nothing
  Set pMultiSymbol = Nothing
  Set pInnerSymbol = Nothing
  Set pInnerLine = Nothing
  Set pLineFill = Nothing
  Set pLineFill2 = Nothing
  Set pOutlineLineSymbol = Nothing

End Function

Public Function Make3CrosshatchSymbol(pColor As IRgbColor, pOutlineColor As IRgbColor, strFillLine As String) As IFillSymbol

  Dim pCartoline As ICartographicLineSymbol
  Set pCartoline = New CartographicLineSymbol
  With pCartoline
    .Cap = esriLCSButt
    .Join = esriLJSBevel
    .Color = pColor
    .Width = 1
  End With

  Dim pMultiSymbol As IMultiLayerFillSymbol
  Set pMultiSymbol = New MultiLayerFillSymbol

  Dim lngAngle1 As Long
  Dim lngAngle2 As Long
  Dim lngAngle3 As Long
  Randomize
  lngAngle1 = Rnd() * 180
  lngAngle2 = lngAngle1 + 60
  lngAngle3 = lngAngle2 + 60

  Dim pLineFill As ILineFillSymbol
  Set pLineFill = New LineFillSymbol
  With pLineFill
    .Angle = lngAngle1
    .Separation = 5
    .Offset = 0
  End With
  Set pLineFill.LineSymbol = pCartoline
  pMultiSymbol.AddLayer pLineFill

  Dim pLineFill2 As ILineFillSymbol
  Set pLineFill2 = New LineFillSymbol
  With pLineFill2
    .Angle = lngAngle2
    .Separation = 5
    .Offset = 0
  End With
  Set pLineFill2.LineSymbol = pCartoline
  pMultiSymbol.AddLayer pLineFill2

  Dim pLineFill3 As ILineFillSymbol
  Set pLineFill3 = New LineFillSymbol
  With pLineFill3
    .Angle = lngAngle3
    .Separation = 5
    .Offset = 0
  End With
  Set pLineFill3.LineSymbol = pCartoline
  pMultiSymbol.AddLayer pLineFill3

  Dim pOutlineLineSymbol As ISimpleLineSymbol
  Set pOutlineLineSymbol = New SimpleLineSymbol

  pOutlineLineSymbol.Color = pOutlineColor
  pOutlineLineSymbol.Width = 2
  pOutlineLineSymbol.Style = esriSLSSolid

  pMultiSymbol.Outline = pOutlineLineSymbol

  Set Make3CrosshatchSymbol = pMultiSymbol

  strFillLine = "3" & vbTab & Format(pColor.RGB, "0") & vbTab & Format(pOutlineColor.RGB, "0") & vbTab
  strFillLine = strFillLine & Format(pCartoline.Cap, "0") & vbTab & Format(pCartoline.Join, "0") & vbTab & _
      Format(pCartoline.Width, "0") & vbTab & Format(lngAngle1, "0") & vbTab & Format(lngAngle2, "0") & vbTab & _
      Format(lngAngle3, "0") & vbTab & Format(pLineFill.Separation, "0.000") & vbTab & _
      Format(pLineFill2.Separation, "0.000") & vbTab & Format(pLineFill3.Separation, "0.000") & vbTab & _
      Format(pLineFill.Offset, "0.000") & vbTab & Format(pLineFill2.Offset, "0.000") & vbTab & _
      Format(pLineFill3.Offset, "0.000") & vbTab & Format(pOutlineLineSymbol.Width, "0.000") & vbTab & _
      Format(pOutlineLineSymbol.Style, "0")

  GoTo ClearMemory

ClearMemory:
  Set pCartoline = Nothing
  Set pMultiSymbol = Nothing
  Set pLineFill = Nothing
  Set pLineFill2 = Nothing
  Set pOutlineLineSymbol = Nothing
  Set pOutlineColor = Nothing

End Function

Public Function MakeMarkerSymbol(pColor As IRgbColor, strMarkerLine As String) As IMarkerSymbol

  Dim pMarkerSymbol As ICharacterMarkerSymbol
  Set pMarkerSymbol = New CharacterMarkerSymbol

  Dim pFont As IFontDisp
  Set pFont = New StdFont
  pFont.Name = "ESRI Default Marker"
  pFont.size = 10
  pFont.Bold = True

  Randomize
  Dim lngUnicode As Long
  Do Until lngUnicode >= 33 And lngUnicode <= 110 And (lngUnicode < 97 Or lngUnicode > 105)
    lngUnicode = CLng(Rnd() * 111)
  Loop

  pMarkerSymbol.size = 13
  pMarkerSymbol.Font = pFont
  pMarkerSymbol.CharacterIndex = lngUnicode
  pMarkerSymbol.Color = pColor

  strMarkerLine = Format(pMarkerSymbol.size, "0") & vbTab & pFont.Name & vbTab & Format(pFont.size, "0") & vbTab & _
      CStr(pFont.Bold) & vbTab & Format(lngUnicode, "0") & vbTab & Format(pColor.RGB, "0")

  Set MakeMarkerSymbol = pMarkerSymbol

ClearMemory:
  Set pMarkerSymbol = Nothing
  Set pFont = Nothing

End Function

Public Function ReturnMarkerFromLine(strLine As String) As IMarkerSymbol
  Dim strObjects() As String
  strObjects = Split(strLine, vbTab)

  Dim lngMarkerSize As Long
  Dim strFontName As String
  Dim lngFontSize As Long
  Dim booFontBold As Boolean
  Dim lngUnicode As Long
  Dim lngRGB As Long
  Dim pColor As IRgbColor

  lngMarkerSize = CLng(strObjects(0))
  strFontName = strObjects(1)
  lngFontSize = CLng(strObjects(2))
  booFontBold = CBool(strObjects(3))
  lngUnicode = CLng(strObjects(4))
  lngRGB = CLng(strObjects(5))

  Dim pMarkerSymbol As ICharacterMarkerSymbol
  Set pMarkerSymbol = New CharacterMarkerSymbol

  Dim pFont As IFontDisp
  Set pFont = New StdFont
  pFont.Name = strFontName
  pFont.size = lngFontSize
  pFont.Bold = booFontBold

  Set pColor = New RgbColor
  pColor.RGB = lngRGB

  pMarkerSymbol.size = lngMarkerSize
  pMarkerSymbol.Font = pFont
  pMarkerSymbol.CharacterIndex = lngUnicode
  pMarkerSymbol.Color = pColor

  Set ReturnMarkerFromLine = pMarkerSymbol

ClearMemory:
  Erase strObjects
  Set pColor = Nothing
  Set pMarkerSymbol = Nothing
  Set pFont = Nothing

End Function

Public Sub InsertHeaderInfo(strSite As String, strPlot As String, strCrew As String, _
    strPhoto As String, strDate As String, strUTME As String, strUTMN As String, strComment As String, _
    strElev As String, strQuadrat As String)

  Debug.Print "-------------------"
  Dim pMxDoc As IMxDocument
  Dim pElements As esriSystem.IArray
  Dim pElement As IElement
  Dim pEnv As IEnvelope
  Dim pPolyline As IPolyline
  Dim pPtColl As IPointCollection
  Dim pDataElement As IElement
  Dim pPoint As IPoint
  Dim pText As ITextElement
  Dim pElementProps As IElementProperties

  Set pEnv = New Envelope
  Set pMxDoc = ThisDocument
  Dim pClone As IClone

  Dim pCalibri10 As ITextSymbol
  Set pCalibri10 = New TextSymbol
  Dim pBlack As IRgbColor
  Set pBlack = New RgbColor
  pBlack.Red = 0
  pBlack.Green = 0
  pBlack.Blue = 0
  Dim pMaroon As IRgbColor
  Set pMaroon = New RgbColor
  pMaroon.RGB = RGB(149, 32, 32)
  Dim pFont As IFont
  Set pFont = New SystemFont
  pFont.Name = "Calibri"
  pFont.Bold = True
  pFont.Italic = False
  pFont.size = 9.75
  pFont.Weight = 800
  pCalibri10.Font = pFont
  pCalibri10.Font.Bold = True
  pCalibri10.Angle = 0
  pCalibri10.Color = pMaroon
  pCalibri10.HorizontalAlignment = esriTHALeft
  pCalibri10.VerticalAlignment = esriTVABottom
  pCalibri10.RightToLeft = False
  pCalibri10.size = 10

  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMxDoc.PageLayout

  Call ReplaceTextElement("SiteLine", "SiteText", pCalibri10, pMxDoc, strSite, pGContainer)

  Call ReplaceTextElement("PlotLine", "PlotText", pCalibri10, pMxDoc, strPlot, pGContainer)

  Call ReplaceTextElement("CrewLine", "CrewText", pCalibri10, pMxDoc, strUTMN & "N, " & strUTME & "E; NAD 1983, Zone 12", pGContainer)

  Call ReplaceTextElement("PhotoLine", "PhotoText", pCalibri10, pMxDoc, strElev, pGContainer)

  Call ReplaceTextElement("DateLine", "DateText", pCalibri10, pMxDoc, strDate, pGContainer)

  Call ReplaceTextElement("CommentLine", "CommentText", pCalibri10, pMxDoc, strComment, pGContainer)

  pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pElements = Nothing
  Set pElement = Nothing
  Set pEnv = Nothing
  Set pPolyline = Nothing
  Set pPtColl = Nothing
  Set pPoint = Nothing
  Set pClone = Nothing
  Set pCalibri10 = Nothing
  Set pBlack = Nothing
  Set pFont = Nothing
  Set pGContainer = Nothing
  Set pDataElement = Nothing
  Set pText = Nothing

End Sub

Public Sub CreateCornerTab()

  Dim pMxDoc As IMxDocument
  Dim pElements As esriSystem.IArray
  Dim pElement As IElement
  Dim pEnv As IEnvelope
  Dim pPolyline As IPolyline
  Dim pPtColl As IPointCollection
  Dim pTextElement As IElement
  Dim pPoint As IPoint
  Dim pText As ITextElement
  Dim pElementProps As IElementProperties

  Set pEnv = New Envelope
  Set pMxDoc = ThisDocument
  Dim pClone As IClone

  Dim pGaramond As ITextSymbol
  Set pGaramond = New TextSymbol
  Dim pBlack As IRgbColor
  Set pBlack = New RgbColor
  pBlack.RGB = RGB(0, 0, 0)
  Dim pWhite As IRgbColor
  Set pWhite = New RgbColor
  pWhite.RGB = RGB(255, 255, 255)
  Dim pBlue As IRgbColor
  Set pBlue = New RgbColor
  pBlue.RGB = RGB(0, 0, 127)
  Dim pFont As IFont
  Set pFont = New SystemFont
  pFont.Name = "Arial"
  pFont.Bold = True
  pFont.Italic = False
  pFont.size = 9.75
  pFont.Weight = 800
  pGaramond.Font = pFont
  pGaramond.Font.Bold = True
  pGaramond.Angle = 0
  pGaramond.Color = pBlue
  pGaramond.HorizontalAlignment = esriTHALeft
  pGaramond.VerticalAlignment = esriTVABottom
  pGaramond.RightToLeft = False
  pGaramond.size = 11

  Dim strElementName As String
  Dim strLineName As String
  strElementName = "Corner Tag"
  strLineName = "Corner Line"

  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMxDoc.PageLayout

  MyGeneralOperations.DeleteGraphicsByName pMxDoc, strElementName, True
  MyGeneralOperations.DeleteGraphicsByName pMxDoc, strLineName, True

  Dim pTransform As ITransform2D
  Dim dblPI As Double
  dblPI = 3.14159265358979
  Set pEnv = New Envelope

  Set pTextElement = New TextElement
  Set pPoint = New Point
  pPoint.PutCoords 0.16, 9.726
  Set pText = pTextElement
  pText.Symbol = pGaramond
  pText.Text = "Tag"
  pTextElement.Geometry = pPoint
  Set pElementProps = pTextElement
  pElementProps.Name = strElementName

  Set pTransform = pTextElement
  pTransform.Rotate pPoint, dblPI / 4

  pGContainer.AddElement pTextElement, 0

  Set pPolyline = New Polyline
  Set pPtColl = pPolyline
  Set pPoint = New Point
  pPoint.PutCoords 0.08, 9.745
  pPtColl.AddPoint pPoint
  pPoint.PutCoords 0.35, 9.745
  pPtColl.AddPoint pPoint
  pPoint.PutCoords 0.35, 10.015
  pPtColl.AddPoint pPoint

  Dim pLineSym As IMultiLayerLineSymbol
  Dim pLine1 As ISimpleLineSymbol
  Dim pLine2 As ISimpleLineSymbol
  Set pLineSym = New MultiLayerLineSymbol
  Set pLine1 = New SimpleLineSymbol
  pLine1.Color = pBlack
  pLine1.Style = esriSLSSolid
  pLine1.Width = 3
  pLineSym.AddLayer pLine1
  Set pLine2 = New SimpleLineSymbol
  pLine2.Color = pWhite
  pLine2.Style = esriSLSSolid
  pLine2.Width = 0.75
  pLineSym.AddLayer pLine2

  Dim pLine As IElement
  Dim pLineElement As ILineElement
  Set pLineElement = New LineElement
  Set pLine = pLineElement
  pLine.Geometry = pPolyline
  pLineElement.Symbol = pLineSym
  Set pElementProps = pLine
  pElementProps.Name = strLineName
  pGContainer.AddElement pLine, 0

  pMxDoc.ActiveView.Refresh

ClearMemory:
  Set pMxDoc = Nothing
  Set pElements = Nothing
  Set pElement = Nothing
  Set pEnv = Nothing
  Set pPolyline = Nothing
  Set pPtColl = Nothing
  Set pTextElement = Nothing
  Set pPoint = Nothing
  Set pText = Nothing
  Set pElementProps = Nothing
  Set pClone = Nothing
  Set pGaramond = Nothing
  Set pBlack = Nothing
  Set pWhite = Nothing
  Set pBlue = Nothing
  Set pFont = Nothing
  Set pGContainer = Nothing
  Set pTransform = Nothing
  Set pLineSym = Nothing
  Set pLine1 = Nothing
  Set pLine2 = Nothing
  Set pLine = Nothing
  Set pLineElement = Nothing

End Sub

Public Sub ReplaceTextElement(strElementName As String, strNewElementName As String, pTextSymbol As ITextSymbol, _
    pMxDoc As IMxDocument, strText As String, pGContainer As IGraphicsContainer)

  MyGeneralOperations.DeleteGraphicsByName pMxDoc, strNewElementName, True

  Dim pElements As esriSystem.IArray
  Dim pElement As IElement
  Dim pEnv As IEnvelope
  Dim pDataElement As IElement
  Dim pPoint As IPoint
  Dim pText As ITextElement
  Dim pElementProps As IElementProperties

  Set pEnv = New Envelope

  Set pElements = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, strElementName, True)
  Set pElement = pElements.Element(0)
  pElement.QueryBounds pMxDoc.ActiveView.ScreenDisplay, pEnv

  Set pDataElement = New TextElement
  Set pPoint = New Point
  pPoint.PutCoords pEnv.XMin + 0.1, pEnv.YMax + 0.01
  Set pText = pDataElement

  pText.Symbol = pTextSymbol
  pText.Text = strText
  pDataElement.Geometry = pPoint
  Set pElementProps = pDataElement
  pElementProps.Name = strNewElementName
  pGContainer.AddElement pDataElement, 0

  Dim dblWidth As Double
  Dim dblTextWidth As Double
  Dim pClone As IClone
  Dim pCloneSymbol As ITextSymbol
  Dim pNewEnv As IEnvelope
  Dim pNewElement As IElement
  Set pNewEnv = New Envelope

  Set pClone = pTextSymbol
  Set pCloneSymbol = pClone.Clone

  Dim lngCounter As Long
  lngCounter = 0
  Dim lngResponse As Long

  If strElementName = "CommentLine" Then
    Set pNewElement = pDataElement
    dblWidth = pEnv.Width
    pElement.QueryBounds pMxDoc.ActiveView.ScreenDisplay, pNewEnv
    dblTextWidth = pNewEnv.Width
    Do Until dblTextWidth <= dblWidth
      lngCounter = lngCounter + 1
      If lngCounter Mod 20 = 0 Then
        Debug.Print "Having trouble fitting comment text:  Iteration = " & Format(lngCounter, "#,##0") & ", Text Symbol Size = " & _
            Format(pCloneSymbol.size, "0.0000")
        DoEvents
      End If
      MyGeneralOperations.DeleteGraphicsByName pMxDoc, strNewElementName, True
      pCloneSymbol.size = pCloneSymbol.size * 0.95

      Set pDataElement = New TextElement
      Set pText = pDataElement

      pText.Symbol = pCloneSymbol
      pText.Text = strText
      pDataElement.Geometry = pPoint
      Set pElementProps = pDataElement
      pElementProps.Name = strNewElementName
      pGContainer.AddElement pDataElement, 0
    Loop
  End If

ClearMemory:
  Set pElements = Nothing
  Set pElement = Nothing
  Set pEnv = Nothing
  Set pDataElement = Nothing
  Set pPoint = Nothing
  Set pText = Nothing
  Set pElementProps = Nothing

End Sub

Public Sub ConstructLegend(pLegendColl As Collection, strLegendKeys() As String, pAreaColl As Collection, _
    pMxDoc As IMxDocument)

  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Legend_Elements", True

  Dim pGaramond12 As ITextSymbol
  Dim pBlack As IRgbColor
  Dim pMaroon As IRgbColor
  Dim pFont As IFont
  Dim lngIndex As Long
  Dim dblX As Double
  Dim dblY As Double
  Dim strSpecies As String
  Dim pElements As esriSystem.IArray
  Dim pElement As IElement
  Dim pEnv As IEnvelope
  Dim pDataElement As IElement
  Dim pGeomElement As IElement
  Dim pPoint As IPoint
  Dim pText As ITextElement
  Dim pElementProps As IElementProperties
  Dim pGContainer As IGraphicsContainer
  Dim pGContainerSelect As IGraphicsContainerSelect
  Dim dblArea As Double
  Dim strArea As String
  Dim pMarkerPoint As IPoint
  Dim pPolygon As IPolygon
  Dim pMarkerSymbol As IMarkerSymbol
  Dim pPolygonSymbol As IFillSymbol
  Dim varPair() As Variant
  Dim dblMarkerX As Double
  Dim dblFillX As Double
  Dim pSubEnv As IEnvelope
  Dim pClone As IClone
  Dim pWorkingMarkerSymbol As IMarkerSymbol

  Dim pActiveView As IActiveView
  Dim pPageLayout As IPageLayout
  Dim pDisplay As IDisplay

  Set pActiveView = pMxDoc.ActiveView ' <-- Better be page layout if this function is being called
  Set pPageLayout = pActiveView
  Set pDisplay = pActiveView.ScreenDisplay

  Set pGaramond12 = New TextSymbol
  Set pBlack = New RgbColor
  pBlack.Red = 0
  pBlack.Green = 0
  pBlack.Blue = 0
  Set pMaroon = New RgbColor
  pMaroon.RGB = RGB(149, 32, 32)
  Set pFont = New SystemFont
  pFont.Name = "Garamond"
  pFont.Bold = False
  pFont.Italic = False
  pFont.size = 9.75
  pFont.Weight = 400
  pGaramond12.Font = pFont
  pGaramond12.Font.Bold = False
  pGaramond12.Angle = 0
  pGaramond12.Color = pBlack
  pGaramond12.HorizontalAlignment = esriTHALeft
  pGaramond12.VerticalAlignment = esriTVABottom
  pGaramond12.RightToLeft = False
  pGaramond12.size = 11 / 0.95

  Set pGContainer = pMxDoc.PageLayout
  Set pGContainerSelect = pGContainer

  Dim pOverallLegendExtent As IEnvelope
  Set pOverallLegendExtent = New Envelope
  pOverallLegendExtent.PutCoords 0, 0, 8, -10

  Dim dblYInc As Double
  dblYInc = 0.2
  Dim dblRunningMarkerSizeRatio As Double
  dblRunningMarkerSizeRatio = 1# / 0.95

  Dim pAllElements As esriSystem.IArray
  Set pAllElements = New esriSystem.Array

  If pLegendColl.Count > 0 Then

    QuickSort.StringsAscending strLegendKeys, 0, UBound(strLegendKeys)

    Do Until pOverallLegendExtent.YMin >= 0.1
      Set pOverallLegendExtent = New Envelope
      dblYInc = dblYInc - 0.01
      pGaramond12.size = pGaramond12.size * 0.95
      dblRunningMarkerSizeRatio = dblRunningMarkerSizeRatio * 0.95
      pAllElements.RemoveAll

      dblX = 1.2
      dblY = 1.9

      For lngIndex = 1 To UBound(strLegendKeys) + 1
        strSpecies = strLegendKeys(lngIndex - 1)
        dblArea = pAreaColl.Item(strSpecies)
        strArea = " [% Cover = " & Format(dblArea, "0.000%") & "]"
        varPair = pLegendColl.Item(strSpecies)

        If (UBound(strLegendKeys) + 1) Mod 2 = 0 Then
          If (lngIndex) Mod (CLng((UBound(strLegendKeys) + 1) / 2) + 1) = 0 Then
            dblX = dblX + 3.9
            dblY = 1.9 - dblYInc ' 1.71
          Else
            dblY = dblY - dblYInc
          End If
        Else
          If (lngIndex) Mod (CLng(((UBound(strLegendKeys) + 1) / 2) + 0.5) + 1) = 0 Then
            dblX = dblX + 3.9
            dblY = 1.9 - dblYInc ' 1.71
          Else
            dblY = dblY - dblYInc
          End If
        End If

        Set pDataElement = New TextElement
        Set pPoint = New Point
        pPoint.PutCoords dblX, dblY
        Set pText = pDataElement

        pText.Symbol = pGaramond12
        pText.Text = strSpecies & strArea
        pDataElement.Geometry = pPoint
        Set pElementProps = pDataElement
        pElementProps.Name = "Legend_Elements"

        Set pSubEnv = New Envelope
        pDataElement.QueryBounds pDisplay, pSubEnv
        pOverallLegendExtent.Union pSubEnv

        pAllElements.Add pDataElement

        If Not IsNull(varPair(0)) Then
          dblFillX = dblX - 0.5
          dblMarkerX = dblX - 0.68
        Else
          If Not IsNull(varPair(1)) Then
            dblMarkerX = dblX - 0.18
          End If
        End If

        If Not IsNull(varPair(0)) Then  ' IF POLYGON
          Set pPolygonSymbol = varPair(0)
          Set pEnv = New Envelope
          pEnv.PutCoords dblX - 0.5, dblY + 0.01, dblX - 0.1, dblY + (dblYInc - 0.03)  ' 0.16
          Set pPolygon = MyGeometricOperations.EnvelopeToPolygon(pEnv)

          Set pGeomElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
              "Legend_Elements", pPolygonSymbol, True, False)
          Set pSubEnv = New Envelope
          pGeomElement.QueryBounds pDisplay, pSubEnv
          pOverallLegendExtent.Union pSubEnv
          pAllElements.Add pGeomElement
        End If

        If Not IsNull(varPair(1)) Then  ' IF POINT
          Set pMarkerSymbol = varPair(1)
          Set pClone = pMarkerSymbol
          Set pWorkingMarkerSymbol = pClone.Clone
          pWorkingMarkerSymbol.size = pWorkingMarkerSymbol.size * dblRunningMarkerSizeRatio
          Set pMarkerPoint = New Point
          pMarkerPoint.PutCoords dblMarkerX, dblY + 0.095

          Set pGeomElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pMarkerPoint, _
              "Legend_Elements", pWorkingMarkerSymbol, True, False)
          Set pSubEnv = New Envelope
          pGeomElement.QueryBounds pDisplay, pSubEnv
          pOverallLegendExtent.Union pSubEnv

          pAllElements.Add pGeomElement
        End If

      Next lngIndex

      MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Legend_Elements", True
      For lngIndex = 0 To pAllElements.Count - 1
        Set pDataElement = pAllElements.Element(lngIndex)
        pGContainer.AddElement pDataElement, 0
      Next lngIndex
      pActiveView.Refresh

      Debug.Print "X: " & Format(pOverallLegendExtent.XMin, "0.0") & " to " & Format(pOverallLegendExtent.XMax, "0.0")
      Debug.Print "Y: " & Format(pOverallLegendExtent.YMin, "0.0") & " to " & Format(pOverallLegendExtent.YMax, "0.0")
      DoEvents

    Loop

  Else
    dblX = 1.2
    dblY = 1.71

    Set pDataElement = New TextElement
    Set pPoint = New Point
    pPoint.PutCoords dblX, dblY
    Set pText = pDataElement

    pText.Symbol = pGaramond12
    pText.Text = "No species observed"
    pDataElement.Geometry = pPoint
    Set pElementProps = pDataElement
    pElementProps.Name = "Legend_Elements"
    pGContainer.AddElement pDataElement, 0

  End If

  pGContainerSelect.UnselectAllElements

ClearMemory:
  Set pGaramond12 = Nothing
  Set pBlack = Nothing
  Set pMaroon = Nothing
  Set pFont = Nothing
  Set pElements = Nothing
  Set pElement = Nothing
  Set pEnv = Nothing
  Set pDataElement = Nothing
  Set pGeomElement = Nothing
  Set pPoint = Nothing
  Set pText = Nothing
  Set pElementProps = Nothing
  Set pGContainer = Nothing
  Set pGContainerSelect = Nothing
  Set pMarkerPoint = Nothing
  Set pPolygon = Nothing
  Set pMarkerSymbol = Nothing
  Set pPolygonSymbol = Nothing
  Erase varPair
  Set pSubEnv = Nothing
  Set pActiveView = Nothing
  Set pPageLayout = Nothing
  Set pDisplay = Nothing
  Set pOverallLegendExtent = Nothing

End Sub


