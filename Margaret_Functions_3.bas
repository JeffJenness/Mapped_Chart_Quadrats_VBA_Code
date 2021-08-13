Attribute VB_Name = "Margaret_Functions_3"
Option Explicit
Public Sub RunExport()
  ExportSubsetsOfSpeciesShapefiles
End Sub

Public Sub ExportSubsetsOfSpeciesShapefiles(Optional booDoAll As Boolean = False, Optional booDo10 As Boolean = False)
  
  Dim lngStartYear As Long
  Dim lngEndYear As Long
  lngStartYear = 2002
  lngEndYear = 2020
  
  ' DON'T BRING IN ANY EMPTY GEOMETRIES OR "No Point Species" OR "No Polygon Species"
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
      
  Dim pQuadData As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim strQuadratNames() As String
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev(strQuadratNames, pPlotToQuadratConversion, _
      pQuadratToPlotConversion, varSites, varSitesSpecific)
  
  Dim pSitesSurveyedByYearColl As Collection
  Dim pYearsSiteSurveyed As Collection
  Dim booSurveyedThisYear As Boolean
  Set pSitesSurveyedByYearColl = More_Margaret_Functions.ReturnCollectionOfYearsSurveyedByQuadrat(lngStartYear, lngEndYear)
  
  Dim strCombinePath As String
  Dim strRecreatedModifiedRoot As String
  Dim strExtractShapefileFolder As String
  Dim pWS As IWorkspace
  Dim pFeatWS As IFeatureWorkspace
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pEnumDataset As IEnumDataset
  Dim pFClass As IFeatureClass
  Dim strPrefix As String
  Dim strSuffix As String
  Dim pQueryFilt As IQueryFilter
  Dim pNewFClass As IFeatureClass
  Dim pDataset As IDataset
  Dim strSpecies As String
  
  Call DeclareWorkspaces(strCombinePath, , , , strRecreatedModifiedRoot, , strExtractShapefileFolder)
  
'  strRecreatedModifiedRoot = strRecreatedModifiedRoot & "_Temp_Apr_16"
'  If Not aml_func_mod.ExistFileDir(strRecreatedModifiedRoot & "\Shapefiles") Then
'    MyGeneralOperations.CreateNestedFoldersByPath (strRecreatedModifiedRoot & "\Shapefiles")
'  End If
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strRecreatedModifiedRoot & "\Shapefiles", 0)
  Set pFeatWS = pWS
  
  
  
  Dim varSpeciesArray() As Variant
  ' FROM OCTOBER 4, 2019
'  varSpeciesArray = Array("Artemisia ludoviciana")
  ' FROM FEB. 4 2021
  If booDo10 Then
    strExtractShapefileFolder = Replace(strExtractShapefileFolder, "Shapefile_Extractions", "Shapefile_Extractions_10_Focal_Species")
    varSpeciesArray = Array("Poa fendleriana", "Carex geophila", "Elymus elymoides", "Koeleria macrantha", "Festuca arizonica", _
      "Bromus ciliatus", "Muhlenbergia tricholepis", "Sporobolus interruptus", "Muhlenbergia montana", "Bouteloua gracilis")
  End If
  ' FROM November 18, 2019

'  varSpeciesArray = Array("Poa fendleriana")

  If booDoAll Then
    varSpeciesArray = CreateArrayOfAllSpeciesNames
  End If
  
  Dim lngIndex As Long
  Dim strCurrentDate As String
  strCurrentDate = Format(Now, "mmm_d_yyyy")
  Dim strFolderName As String
  Dim varFieldIndexArray() As Variant
  Dim pSrcFCursor As IFeatureCursor
  Dim pSrcFeature As IFeature
  Dim pDestFCursor As IFeatureCursor
  Dim pDestFBuffer As IFeatureBuffer
  Dim lngIndex2 As Long
  Dim lngIndex3 As Long
  Dim lngIndex4 As Long
  Dim strSites() As String
  Dim strPlots() As String
  Dim lngSiteArrayIndex As Long
  Dim lngPlotArrayIndex As Long
  Dim lngSiteIDIndex As Long
  Dim lngPlotIDIndex As Long
  Dim lngQuadratIDIndex As Long
  Dim strQuadrat As String
  Dim pDoneSites As Collection
  Dim pDonePlots As Collection
  Dim strSite As String
  Dim strPlot As String
  Dim strSiteAsName As String
  Dim strPlotAsName As String
  Dim lngYearIndex As Long
  Dim strNewShapefileName As String
  Dim lngSpeciesCounter As Long
  Dim lngSiteCounter As Long
  Dim lngPlotCounter As Long
  
  Dim strReport As String
  
  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose
  
'  Big Fill is both Site 1 and site 8?
    
  Set pQueryFilt = New QueryFilter
  ' Split by Species\Site\Plot\Yearly_shapefiles!  Make separate folders for each plot.
  
  For lngIndex = 0 To UBound(varSpeciesArray)
    strSpecies = CStr(varSpeciesArray(lngIndex))
    lngSpeciesCounter = lngSpeciesCounter + 1
    lngSiteCounter = 0
    lngPlotCounter = 0
    Debug.Print "Species #" & Format(lngSpeciesCounter, "0") & "] Working on Species '" & strSpecies & "'"
    strReport = strReport & "Species #" & Format(lngSpeciesCounter, "0") & "] Working on Species '" & strSpecies & "'" & vbCrLf
    DoEvents
    
    
    ' GET SITES THAT CONTAIN SPECIES
    Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
    pEnumDataset.Reset
    Set pFClass = pEnumDataset.Next
    
    Do Until pFClass Is Nothing
      lngSiteIDIndex = pFClass.FindField("Site")
      lngPlotIDIndex = pFClass.FindField("Plot")
      lngQuadratIDIndex = pFClass.FindField("Quadrat")
      Set pDataset = pFClass
      If pDataset.BrowseName <> "Cover_All" And pDataset.BrowseName <> "Density_All" Then
        MyGeneralOperations.ReturnQuerySpecialCharacters pFClass, strPrefix, strSuffix
        
        ' NARROW DOWN SITES
        lngSiteArrayIndex = -1
        Set pDoneSites = New Collection
        
        pQueryFilt.WhereClause = strPrefix & "species" & strSuffix & " = '" & CStr(varSpeciesArray(lngIndex)) & "'"
        If pFClass.FeatureCount(pQueryFilt) > 0 Then
          Set pSrcFCursor = pFClass.Search(pQueryFilt, False)
          Set pSrcFeature = pSrcFCursor.NextFeature
          Do Until pSrcFeature Is Nothing
            strSite = pSrcFeature.Value(lngSiteIDIndex)
            If InStr(1, strSite, "/") > 0 Then
              DoEvents
            End If
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneSites, strSite) Then
              strSiteAsName = Replace(strSite, "/", "bar")
              strSiteAsName = MyGeneralOperations.ReplaceBadChars(Trim(strSiteAsName), True, True, True, False)
              lngSiteArrayIndex = lngSiteArrayIndex + 1
              ReDim Preserve strSites(1, lngSiteArrayIndex)
              strSites(0, lngSiteArrayIndex) = strSite
              strSites(1, lngSiteArrayIndex) = strSiteAsName
              pDoneSites.Add True, strSite
              
              Debug.Print "...Added Site '" & strSite & "' from shapefile '" & pDataset.BrowseName & "'..."
              strReport = strReport & "...Added Site '" & strSite & "' from shapefile '" & pDataset.BrowseName & "'..." & vbCrLf
            End If
              
            Set pSrcFeature = pSrcFCursor.NextFeature
          Loop
          
          ' NARROW DOWN PLOTS
          If lngSiteArrayIndex > -1 Then
            For lngIndex2 = 0 To lngSiteArrayIndex
              strSite = strSites(0, lngIndex2)
              strSiteAsName = strSites(1, lngIndex2)
              lngSiteCounter = lngSiteCounter + 1
              Debug.Print "    Site " & Format(lngSiteCounter, "0") & "] Working on Site '" & strSite & "'"
              strReport = strReport & "    Site " & Format(lngSiteCounter, "0") & "] Working on Site '" & strSite & "'" & vbCrLf
              DoEvents
              
              lngPlotArrayIndex = -1
              Set pDonePlots = New Collection
              
              
              pQueryFilt.WhereClause = strPrefix & "species" & strSuffix & " = '" & CStr(varSpeciesArray(lngIndex)) & "' AND " & _
                  strPrefix & "Site" & strSuffix & " = '" & strSite & "'"
              If pFClass.FeatureCount(pQueryFilt) > 0 Then
                Set pSrcFCursor = pFClass.Search(pQueryFilt, False)
                Set pSrcFeature = pSrcFCursor.NextFeature
                Do Until pSrcFeature Is Nothing
                  strPlot = pSrcFeature.Value(lngPlotIDIndex)
                  strQuadrat = pSrcFeature.Value(lngQuadratIDIndex)
                  
                  If Not MyGeneralOperations.CheckCollectionForKey(pDonePlots, strPlot) Then
                    strPlotAsName = Replace(strPlot, "/", "bar")
                    strPlotAsName = MyGeneralOperations.ReplaceBadChars(Trim(strPlotAsName), True, True, True, False)
                    lngPlotArrayIndex = lngPlotArrayIndex + 1
                    ReDim Preserve strPlots(2, lngPlotArrayIndex)
                    strPlots(0, lngPlotArrayIndex) = strPlot
                    strPlots(1, lngPlotArrayIndex) = strPlotAsName
                    strPlots(2, lngPlotArrayIndex) = strQuadrat
                    pDonePlots.Add True, strPlot
                  End If
                    
                  Set pSrcFeature = pSrcFCursor.NextFeature
                Loop
                
                If lngPlotArrayIndex > -1 Then     ' export separate shapefiles by year, making folders for each plot
                  For lngIndex3 = 0 To lngPlotArrayIndex
                    strPlot = strPlots(0, lngIndex3)
                    strPlotAsName = strPlots(1, lngIndex3)
                    strQuadrat = strPlots(2, lngIndex3)
                    Set pYearsSiteSurveyed = pSitesSurveyedByYearColl.Item(strQuadrat)
                    
                    lngPlotCounter = lngPlotCounter + 1
                    Debug.Print "        Plot " & Format(lngPlotCounter, "0") & "] Working on Plot '" & strPlot & "'"
                    strReport = strReport & "        Plot " & Format(lngPlotCounter, "0") & "] Working on Plot '" & strPlot & "'" & vbCrLf
                    DoEvents
                    
                    strFolderName = strExtractShapefileFolder & "\" & Replace(strSpecies, " ", "_") & "_" & strCurrentDate & "\" & _
                        strSiteAsName & "_" & strCurrentDate & "\Quadrat_" & strPlotAsName & "_" & strCurrentDate
'                    strFolderName = MyGeneralOperations.MakeUniquedBASEName(strFolderName)
                    
                    MyGeneralOperations.CreateNestedFoldersByPath strFolderName
                    Set pNewWS = pWSFact.OpenFromFile(strFolderName, 0)
                    Set pNewFeatWS = pNewWS
                    
                    For lngYearIndex = lngStartYear To lngEndYear
                      booSurveyedThisYear = pYearsSiteSurveyed.Item(Format(lngYearIndex, "0000"))
                      strNewShapefileName = "Quadrat_" & strPlotAsName & "_" & Format(lngYearIndex, "0000")
                      
                      pQueryFilt.WhereClause = strPrefix & "species" & strSuffix & " = '" & CStr(varSpeciesArray(lngIndex)) & _
                          "' AND " & strPrefix & "Site" & strSuffix & " = '" & strSite & "' AND " & _
                          strPrefix & "Plot" & strSuffix & " = '" & strPlot & "' AND " & _
                          strPrefix & "z_Year" & strSuffix & " = '" & Format(lngYearIndex, "0000") & "'"
                      '"No Cover Species Observed" No Density Species Observed
'                      Debug.Print "              --> Found " & Format(pFClass.FeatureCount(pQueryFilt), "0") & _
                          " records for Year " & Format(lngYearIndex, "0000")
                      ' If, for some reason, shapefile already exists, just write new records to it.
                      ' This could happen if the species was in both Cover and Density feature classes.
                      If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewShapefileName) Then
'                        Debug.Print "************ Already created '" & strNewShapefileName & "'!"
                        Set pNewFClass = pNewFeatWS.OpenFeatureClass(strNewShapefileName)
                      Else
                        If booSurveyedThisYear Then
                          Set pNewFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pFClass, pNewWS, varFieldIndexArray, _
                              strNewShapefileName, True)
                          Set pSrcFCursor = pFClass.Search(pQueryFilt, False)
                          Set pSrcFeature = pSrcFCursor.NextFeature
                          Set pDestFCursor = pNewFClass.Insert(True)
                          Set pDestFBuffer = pNewFClass.CreateFeatureBuffer
              
                          strPurpose = "This shapefile contains all records of '" & strSpecies & "' observed in " & _
                              "Quadrat '" & strPlot & "', Site '" & strSite & "', in " & Format(lngYearIndex, "0000") & ". " & _
                              "These single-species shapefiles are formatted to input directly into Integral Projection " & _
                              "Modeling (IPM) analysis in R."
                          
                          Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewFClass, strAbstract, strPurpose)
    
                          Do Until pSrcFeature Is Nothing
                            For lngIndex4 = 0 To UBound(varFieldIndexArray, 2)
                              pDestFBuffer.Value(varFieldIndexArray(3, lngIndex4)) = pSrcFeature.Value(varFieldIndexArray(1, lngIndex4))
                            Next lngIndex4
                            Set pDestFBuffer.Shape = pSrcFeature.ShapeCopy
                            pDestFCursor.InsertFeature pDestFBuffer
                          
                            Set pSrcFeature = pSrcFCursor.NextFeature
                          Loop
                          pDestFCursor.Flush
                          Set pDestFCursor = Nothing
                          Set pDestFBuffer = Nothing
                        End If
                      End If
                    Next lngYearIndex   ' DONE ROTATING THROUGH YEARS
                  Next lngIndex3        ' DONE ROTATING THROUGH PLOTS
                End If                  ' DONE CHECKING IF SPECIES EXISTS ON ANY PLOTS AT THIS SITE
              End If                    ' DONE CHECKING IF ANY RECORDS FOUND FOR THIS PARTICULAR SITE
            Next lngIndex2              ' DONE ROTATING THROUGH SITES FIND FOR THIS SPECIES
          End If                        ' DONE CHECKING IF ANY SITES FOUND FOR THIS SPECIES
        End If                          ' DONE CHECKING WHETHER SPECIES EXISTS IN CURRENT FEATURE CLASS
      End If
      Set pFClass = pEnumDataset.Next
    Loop
  Next lngIndex
    
'
'
'
'    strFolderName = strExtractShapefileFolder & "\" & Replace(strSpecies, " ", "_") & "_" & strCurrentDate
'    strFolderName = MyGeneralOperations.MakeUniquedBASEName(strFolderName)
'
'    MyGeneralOperations.CreateNestedFoldersByPath strFolderName
'    Set pNewWS = pWSFact.OpenFromFile(strFolderName, 0)
'
'    Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
'    pEnumDataset.Reset
'    Set pFClass = pEnumDataset.Next
'    Set pDataset = pFClass
'
'
'    Do Until pFClass Is Nothing
'      If pDataset.BrowseName <> "Cover_All" And pDataset.BrowseName <> "Density_All" Then
'        MyGeneralOperations.ReturnQuerySpecialCharacters pFClass, strPrefix, strSuffix
'        pQueryFilt.WhereClause = strPrefix & "species" & strSuffix & " = '" & CStr(varSpeciesArray(lngIndex)) & "'"
'        If pFClass.FeatureCount(pQueryFilt) > 0 Then
'
'          Set pNewFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pFClass, pNewWS, varFieldIndexArray, _
'              pDataset.Name, True)
'
'          Set pSrcFCursor = pFClass.Search(pQueryFilt, False)
'          Set pSrcFeature = pSrcFCursor.NextFeature
'          Set pDestFCursor = pNewFClass.Insert(True)
'          Set pDestFBuffer = pNewFClass.CreateFeatureBuffer
'
'          Do Until pSrcFeature Is Nothing
'            For lngIndex2 = 0 To UBound(varFieldIndexArray, 2)
'              pDestFBuffer.Value(varFieldIndexArray(3, lngIndex2)) = pSrcFeature.Value(varFieldIndexArray(1, lngIndex2))
'            Next lngIndex2
'            Set pDestFBuffer.Shape = pSrcFeature.ShapeCopy
'            pDestFCursor.InsertFeature pDestFBuffer
'
'            Set pSrcFeature = pSrcFCursor.NextFeature
'          Loop
'          pDestFCursor.Flush
'          Set pDestFCursor = Nothing
'          Set pDestFBuffer = Nothing
'
'        End If
'      End If
'      Set pFClass = pEnumDataset.Next
'    Loop
'
'  Next lngIndex
  
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  strReport = strReport & MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  Dim pDataObj As New MSForms.DataObject
  pDataObj.SetText strReport
'  pDataObj.PutInClipboard
  Set pDataObj = Nothing

ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pQuadData = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Erase strQuadratNames
  Erase varSites
  Erase varSitesSpecific
  Set pSitesSurveyedByYearColl = Nothing
  Set pYearsSiteSurveyed = Nothing
  Set pWS = Nothing
  Set pFeatWS = Nothing
  Set pNewWS = Nothing
  Set pWSFact = Nothing
  Set pEnumDataset = Nothing
  Set pFClass = Nothing
  Set pQueryFilt = Nothing
  Set pNewFClass = Nothing
  Set pDataset = Nothing
  Erase varSpeciesArray
  Erase varFieldIndexArray
  Set pSrcFCursor = Nothing
  Set pSrcFeature = Nothing
  Set pDestFCursor = Nothing
  Set pDestFBuffer = Nothing
  Erase strSites
  Erase strPlots
  Set pDoneSites = Nothing
  Set pDonePlots = Nothing




  
End Sub

Public Sub TestCreateArray()
  Dim varReturn() As Variant
  varReturn = CreateArrayOfAllSpeciesNames
  Debug.Print "Done..."
  Debug.Print UBound(varReturn)
End Sub

Public Function CreateArrayOfAllSpeciesNames() As Variant()
  
  Dim pApp As IApplication
  Dim pMxDoc As IMxDocument
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngCount As Long
  Dim pCoverFClass As IFeatureClass
  Dim pDensityFClass As IFeatureClass
  
  Dim strCombinePath As String
  Dim strRecreatedModifiedRoot As String
  Dim strExtractShapefileFolder As String
  Dim strFinalFolder As String
  Dim pWS As IWorkspace
  Dim pFeatWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pEnumDataset As IEnumDataset
  Dim pFClass As IFeatureClass
  Dim strPrefix As String
  Dim strSuffix As String
  Dim pQueryFilt As IQueryFilter
  Dim strSpecies As String
  Dim pDoneColl As New Collection
  Dim lngArrayCounter As Long
  Dim strReturn() As String
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngSpeciesIndex As Long
  
  Call DeclareWorkspaces(strCombinePath, , , , strRecreatedModifiedRoot, , strExtractShapefileFolder, strFinalFolder)
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
  Set pFeatWS = pWS
  Set pCoverFClass = pFeatWS.OpenFeatureClass("Cover_All")
  Set pDensityFClass = pFeatWS.OpenFeatureClass("Density_All")
  
  lngCount = pCoverFClass.FeatureCount(Nothing) + pDensityFClass.FeatureCount(Nothing)
  pSBar.ShowProgressBar "Reviewing Species...", 0, lngCount, 1, True
  pProg.position = 0
  
  lngArrayCounter = -1
  
  lngSpeciesIndex = pCoverFClass.FindField("Species")
  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    strSpecies = pFeature.Value(lngSpeciesIndex)
    If InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strSpecies) Then
        DoEvents
        pDoneColl.Add True, strSpecies
        lngArrayCounter = lngArrayCounter + 1
        ReDim Preserve strReturn(lngArrayCounter)
        strReturn(lngArrayCounter) = strSpecies
      End If
    End If
    Set pFeature = pFCursor.NextFeature
  Loop

  lngSpeciesIndex = pDensityFClass.FindField("Species")
  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    strSpecies = pFeature.Value(lngSpeciesIndex)
    If InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strSpecies) Then
        DoEvents
        pDoneColl.Add True, strSpecies
        lngArrayCounter = lngArrayCounter + 1
        ReDim Preserve strReturn(lngArrayCounter)
        strReturn(lngArrayCounter) = strSpecies
      End If
    End If
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.StringsAscending strReturn, 0, UBound(strReturn)
  Dim lngIndex As Long
  Dim varReturn() As Variant
  ReDim varReturn(UBound(strReturn))
  For lngIndex = 0 To UBound(strReturn)
    varReturn(lngIndex) = strReturn(lngIndex)
  Next lngIndex
  
  CreateArrayOfAllSpeciesNames = varReturn
  
  Set pCoverFClass = Nothing
  Set pDensityFClass = Nothing
  Set pFeatWS = Nothing
  Set pWS = Nothing
  pProg.position = 0
  pSBar.HideProgressBar
  
End Function

Public Sub ListShapefileNamesInFolder()
  
  Debug.Print "------------------------------"
  Dim pPaths As esriSystem.IStringArray
  Set pPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2( _
      "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_As_of_May_14_2020", ".shp")
  
  Dim strSubPath As String
  Dim lngIndex As Long
  Dim strSortArray() As String
  Dim lngCounter As Long
  Dim pDoneColl As New Collection
  
  lngCounter = -1
  For lngIndex = 0 To pPaths.Count - 1
    strSubPath = pPaths.Element(lngIndex)
    strSubPath = aml_func_mod.ReturnFilename2(strSubPath)
    If Right(strSubPath, 4) <> ".xml" Then
      strSubPath = Replace(strSubPath, "_CF", "", , , vbTextCompare)
      strSubPath = Replace(strSubPath, "_DF", "", , , vbTextCompare)
      strSubPath = Replace(strSubPath, ".shp", "", , , vbTextCompare)
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strSubPath) Then
        pDoneColl.Add True, strSubPath
        
        lngCounter = lngCounter + 1
        ReDim Preserve strSortArray(lngCounter)
        strSortArray(lngCounter) = strSubPath
      End If
    End If
  Next lngIndex
  
  QuickSort.StringsAscending strSortArray, 0, lngCounter
  
  For lngIndex = 0 To lngCounter
    strSubPath = strSortArray(lngIndex)
    Debug.Print "  varNameLinks_2019(" & Format(lngIndex) & ") = Array(""" & strSubPath & """, """")"
  Next lngIndex
  
    
    
End Sub
