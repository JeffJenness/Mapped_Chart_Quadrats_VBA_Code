Attribute VB_Name = "Margaret_Functions_3"
Option Explicit

Public Sub ExportSubsetsOfSpeciesShapefiles(Optional booDoAll As Boolean = False, Optional booDo10 As Boolean = False, _
    Optional booDoAllBut10 As Boolean = False, Optional booCreateOtherPolygon As Boolean = False, _
    Optional booDoSpecialCase As Boolean = False, Optional strSpecifiedDateText1 As String = "", _
    Optional strSpecifiedDateText2 As String = "", Optional varSpecialCaseArray As Variant)

  Dim strCurrentDate As String

  If Not booDoSpecialCase Then
    strCurrentDate = "Nov_9_2024"
    strCurrentDate = "2002-2024"
  Else
    strCurrentDate = strSpecifiedDateText1 & "_" & strSpecifiedDateText2
  End If

  Dim lngStartYear As Long
  Dim lngEndYear As Long
  lngStartYear = 2002
  lngEndYear = 2032

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
  Dim pReverseQueryFilt As IQueryFilter
  Dim pNewFClass As IFeatureClass
  Dim pDataset As IDataset
  Dim strSpecies As String

  Dim varTempArray() As Variant
  varTempArray = Array("Poa fendleriana", "Carex geophila", "Elymus elymoides", "Koeleria macrantha", "Festuca arizonica", _
      "Bromus ciliatus", "Muhlenbergia tricholepis", "Sporobolus interruptus", "Muhlenbergia montana", "Bouteloua gracilis")
  Dim lngIndex As Long
  Dim p10Coll As New Collection
  For lngIndex = 0 To UBound(varTempArray)
    p10Coll.Add True, varTempArray(lngIndex)
  Next lngIndex

  Call DeclareWorkspaces(strCombinePath, , , , strRecreatedModifiedRoot, , strExtractShapefileFolder)
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strRecreatedModifiedRoot & "\Shapefiles", 0)
  Set pFeatWS = pWS

  Dim varSpeciesArray() As Variant
  If booDo10 Then
    strExtractShapefileFolder = Replace(strExtractShapefileFolder, "Shapefile_Extractions", "Shapefile_Extractions_10_Focal_Species")
    varSpeciesArray = varTempArray
  ElseIf booDoSpecialCase Then
    strExtractShapefileFolder = Replace(strExtractShapefileFolder, "Shapefile_Extractions", "Shapefile_Extractions_Special_Cases")
    varSpeciesArray = varSpecialCaseArray
  End If

  If booDoAll Or booDoAllBut10 Then
    varSpeciesArray = CreateArrayOfAllSpeciesNames(p10Coll, booDoAllBut10)
  End If

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
  Dim pReversePolygon As IPolygon
  Dim varAttributes() As Variant

  Dim strReport As String

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Set pQueryFilt = New QueryFilter
  Set pReverseQueryFilt = New QueryFilter

  For lngIndex = 0 To UBound(varSpeciesArray)
    strSpecies = CStr(varSpeciesArray(lngIndex))
    lngSpeciesCounter = lngSpeciesCounter + 1
    lngSiteCounter = 0
    lngPlotCounter = 0
    Debug.Print "Species #" & Format(lngSpeciesCounter, "0") & "] Working on Species '" & strSpecies & "'"
    strReport = strReport & "Species #" & Format(lngSpeciesCounter, "0") & "] Working on Species '" & strSpecies & "'" & vbCrLf
    DoEvents

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

                    If booDoSpecialCase Then
                      strFolderName = strExtractShapefileFolder & "\" & Replace(strSpecies, " ", "_") & "_" & strCurrentDate & "\" & _
                          strSiteAsName & "_" & strSpecifiedDateText2 & "\Quadrat_" & strPlotAsName & "_" & strSpecifiedDateText2
                    Else
                      strFolderName = strExtractShapefileFolder & "\" & Replace(strSpecies, " ", "_") & "_" & strCurrentDate & "\" & _
                          strSiteAsName & "_" & strCurrentDate & "\Quadrat_" & strPlotAsName & "_" & strCurrentDate
                    End If

                    Do Until InStr(1, strFolderName, "__") = 0
                      strFolderName = Replace(strFolderName, "__", "_")
                    Loop

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
                          " records for Year " & Format(lngYearIndex, "0000")
                      If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewShapefileName) Then
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
                          If booCreateOtherPolygon Then
                            strPurpose = strPurpose & "  This shapefile also contains a single polygon representing " & _
                            "all other species observed on this quadrat, dissolved into a single object.  This 'Other' " & _
                            "polygon can be used to investigate crowding and neighbor effects on '" & strSpecies & _
                            "' in this quadrat."

                            pReverseQueryFilt.WhereClause = strPrefix & "species" & strSuffix & " <> '" & CStr(varSpeciesArray(lngIndex)) & _
                                "' AND " & strPrefix & "Site" & strSuffix & " = '" & strSite & "' AND " & _
                                strPrefix & "Plot" & strSuffix & " = '" & strPlot & "' AND " & _
                                strPrefix & "z_Year" & strSuffix & " = '" & Format(lngYearIndex, "0000") & "'"
                            Set pReversePolygon = ReturnOtherPolygon(pFClass, pReverseQueryFilt, varFieldIndexArray, varAttributes)
                            If Not pReversePolygon.IsEmpty Then
                              For lngIndex4 = 0 To UBound(varFieldIndexArray, 2)
                                pDestFBuffer.Value(varFieldIndexArray(3, lngIndex4)) = varAttributes(lngIndex4)
                              Next lngIndex4
                              Set pDestFBuffer.Shape = pReversePolygon
                              pDestFCursor.InsertFeature pDestFBuffer
                            End If
                          End If

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

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  strReport = strReport & MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  Dim pDataObj As New MSForms.DataObject
  pDataObj.SetText strReport
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

Public Function ReturnOtherPolygon(pFClass As IFeatureClass, pQueryFilt As IQueryFilter, _
    varFieldIndexArray() As Variant, varAttributes() As Variant) As IPolygon

  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFClass
  Dim pDissolveArray As esriSystem.IVariantArray
  Dim booAddedAttributes As Boolean
  Set pDissolveArray = New esriSystem.varArray
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngIndex As Long
  Dim strFieldName As String

  ReDim varAttributes(UBound(varFieldIndexArray, 2))

  booAddedAttributes = False
  Set pFCursor = pFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pDissolveArray.Add pFeature.ShapeCopy

    If Not booAddedAttributes Then

      For lngIndex = 0 To UBound(varFieldIndexArray, 2)
        strFieldName = pFeature.Fields.Field(varFieldIndexArray(1, lngIndex)).Name
        If strFieldName = "Verb_Spcs" Or strFieldName = "species" Then
          varAttributes(lngIndex) = "All Other Species"
        ElseIf strFieldName = "area" Or strFieldName = "x" Or strFieldName = "y" Then
          varAttributes(lngIndex) = -999
        Else
          varAttributes(lngIndex) = pFeature.Value(varFieldIndexArray(1, lngIndex))
        End If
      Next lngIndex

      booAddedAttributes = True
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  Dim pReturn As IPolygon
  If pDissolveArray.Count > 0 Then
    Set pReturn = MyGeometricOperations.UnionGeometries4(pDissolveArray, 200)
  Else
    Set pReturn = New Polygon
  End If
  Set pReturn.SpatialReference = pGeoDataset.SpatialReference

  Set ReturnOtherPolygon = pReturn

  Set pGeoDataset = Nothing
  Set pDissolveArray = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pReturn = Nothing

End Function

Public Function CreateArrayOfAllSpeciesNames(p10Coll As Collection, Optional booSkip10 As Boolean = False) As Variant()

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
  If aml_func_mod.ExistFileDir(strFinalFolder & "\Combined_by_Site.gdb") Then
    Set pWS = pWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
  Else
    Set pWS = pWSFact.OpenFromFile(strFinalFolder & "\Data\Quadrat_Spatial_Data\Combined_by_Site.gdb", 0)
  End If
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
    If Not booSkip10 Or Not MyGeneralOperations.CheckCollectionForKey(p10Coll, strSpecies) Then
      If InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
        If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strSpecies) Then
          DoEvents
          pDoneColl.Add True, strSpecies
          lngArrayCounter = lngArrayCounter + 1
          ReDim Preserve strReturn(lngArrayCounter)
          strReturn(lngArrayCounter) = strSpecies
        End If
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
    If Not booSkip10 Or Not MyGeneralOperations.CheckCollectionForKey(p10Coll, strSpecies) Then
      If InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
        If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strSpecies) Then
          DoEvents
          pDoneColl.Add True, strSpecies
          lngArrayCounter = lngArrayCounter + 1
          ReDim Preserve strReturn(lngArrayCounter)
          strReturn(lngArrayCounter) = strSpecies
        End If
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


