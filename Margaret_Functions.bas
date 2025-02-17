Attribute VB_Name = "Margaret_Functions"
Option Explicit

Public Sub FillCollections(strQuadrat As String, strYear As String, pQuadratColl As Collection, _
    strQuadrats() As String, dblArea As Double, strSpecies As String)

  Dim pYearSubColl As Collection        ' HOLDS AREA VALUES FOR EACH YEAR, FOR A SINGLE QUADRAT.  HELD IN varPair, in pQuadratColl
  Dim strYears() As String
  Dim varPair() As Variant
  Dim dblRunningArea As Double
  Dim pYearSpeciesColl As Collection    ' HOLDS ALL YARS FOR EACH QUADRAT, SO WE CAN SEE IF SPECIES WAS FOUND IN ANY YEAR
  Dim pSubYearSpeciesColl As Collection ' TRUE OR FALSE VALUES STATING WHETHER SPECIES WAS FOUND THAT YEAR

  Dim pCountAndAreaPerYearColl As Collection ' HOLD COUNT AND AREA PER YEAR, FOR EACH SPECIES.
  Dim pCountAreaAllYears As Collection       ' HOLD ALL YEARS FOR EACH QUADRAT
  Dim varCountAreaPair() As Variant          ' Count/Area Pair, to be held in pCountAndAreaPerYearColl
  Dim dblRunningAreaForCount As Double
  Dim lngRunningCount As Long

  If strYear = "2006" And strQuadrat = "Q2" Then
    DoEvents
  End If

  If Not MyGeneralOperations.CheckCollectionForKey(pQuadratColl, strQuadrat) Then
    If MyGeneralOperations.IsDimmed(strQuadrats) Then
      ReDim Preserve strQuadrats(UBound(strQuadrats) + 1)
    Else
      ReDim Preserve strQuadrats(0)
    End If
    strQuadrats(UBound(strQuadrats)) = strQuadrat

    Set pYearSubColl = New Collection
    ReDim strYears(0)
    ReDim varPair(3)

    strYears(0) = strYear
    pYearSubColl.Add dblArea, strYear

    Set pYearSpeciesColl = New Collection
    Set pSubYearSpeciesColl = New Collection
    pSubYearSpeciesColl.Add True, strSpecies
    pYearSpeciesColl.Add pSubYearSpeciesColl, strYear
    Set varPair(2) = pYearSpeciesColl

    ReDim varCountAreaPair(1)
    varCountAreaPair(0) = 1         ' COUNT
    varCountAreaPair(1) = dblArea   ' AREA
    Set pCountAndAreaPerYearColl = New Collection
    pCountAndAreaPerYearColl.Add varCountAreaPair, strSpecies
    Set pCountAreaAllYears = New Collection
    pCountAreaAllYears.Add pCountAndAreaPerYearColl, strYear

  Else
    varPair = pQuadratColl.Item(strQuadrat)
    pQuadratColl.Remove strQuadrat
    strYears = varPair(0)
    Set pYearSubColl = varPair(1)
    Set pYearSpeciesColl = varPair(2)

    Set pCountAreaAllYears = varPair(3)
    If MyGeneralOperations.CheckCollectionForKey(pCountAreaAllYears, strYear) Then
      Set pCountAndAreaPerYearColl = pCountAreaAllYears.Item(strYear)
      pCountAreaAllYears.Remove strYear

      If MyGeneralOperations.CheckCollectionForKey(pCountAndAreaPerYearColl, strSpecies) Then
        varCountAreaPair = pCountAndAreaPerYearColl.Item(strSpecies)
        pCountAndAreaPerYearColl.Remove strSpecies

        lngRunningCount = varCountAreaPair(0)
        dblRunningAreaForCount = varCountAreaPair(1)
      Else
        ReDim varCountAreaPair(1)
        lngRunningCount = 0
        dblRunningAreaForCount = 0
      End If

    Else
      Set pCountAndAreaPerYearColl = New Collection
      ReDim varCountAreaPair(1)
    End If

    lngRunningCount = lngRunningCount + 1
    dblRunningAreaForCount = dblRunningAreaForCount + dblArea
    varCountAreaPair(0) = lngRunningCount
    varCountAreaPair(1) = dblRunningAreaForCount
    pCountAndAreaPerYearColl.Add varCountAreaPair, strSpecies
    pCountAreaAllYears.Add pCountAndAreaPerYearColl, strYear

    If MyGeneralOperations.CheckCollectionForKey(pYearSpeciesColl, strYear) Then
      Set pSubYearSpeciesColl = pYearSpeciesColl.Item(strYear)
      pYearSpeciesColl.Remove strYear
      If Not MyGeneralOperations.CheckCollectionForKey(pSubYearSpeciesColl, strSpecies) Then
        pSubYearSpeciesColl.Add True, strSpecies
      End If
    Else
      Set pSubYearSpeciesColl = New Collection
      pSubYearSpeciesColl.Add True, strSpecies
    End If
    pYearSpeciesColl.Add pSubYearSpeciesColl, strYear

    If MyGeneralOperations.CheckCollectionForKey(pYearSubColl, strYear) Then
      dblRunningArea = pYearSubColl.Item(strYear)
      pYearSubColl.Remove (strYear)
      pYearSubColl.Add dblRunningArea + dblArea, strYear
    Else
      pYearSubColl.Add dblArea, strYear
      ReDim Preserve strYears(UBound(strYears) + 1)
      strYears(UBound(strYears)) = strYear
    End If
  End If

  varPair(0) = strYears
  Set varPair(1) = pYearSubColl
  Set varPair(2) = pYearSpeciesColl
  Set varPair(3) = pCountAreaAllYears

  pQuadratColl.Add varPair, strQuadrat

ClearMemory:
  Set pYearSubColl = Nothing
  Erase strYears
  Erase varPair

End Sub

Public Sub ShiftPolygon(pPolygon As IPolygon, dblCentroidX As Double, dblCentroidY As Double)

  Dim pTransform As ITransform2D

  If dblCentroidX <> 0 And dblCentroidY <> 1 Then
    Set pTransform = pPolygon
    pTransform.Move dblCentroidX, dblCentroidY - 1
  End If

  Set pTransform = Nothing

End Sub

Public Sub FillQuadratCenter(strQuadrat As String, pPlotLocColl As Collection, dblCentroidX As Double, dblCentroidY As Double)

  Dim varArray() As Variant
  If MyGeneralOperations.CheckCollectionForKey(pPlotLocColl, strQuadrat) Then
    varArray = pPlotLocColl.Item(strQuadrat)
    dblCentroidY = varArray(1)
    dblCentroidX = varArray(0)
  Else
    dblCentroidY = 0.5
    dblCentroidX = 0.5
  End If

  Erase varArray

End Sub

Public Function FillQuadratNameColl_Rev(strQuadrats() As String, Optional pPlotToQuadratConversion As Collection, _
    Optional pQuadratToPlotConversion As Collection, Optional varSites As Variant, _
    Optional varSitesSpecific As Variant) As Collection

  Dim strPath As String
  strPath = "D:\arcGIS_stuff\consultation\Margaret_Moore\From_Margaret_Dec_2017_Jan_2018\" & _
      "Hill-Wild Bill_Old and New Quadrat Numbers by Site_2016_mod_Feb_2018.txt"

  Dim strFile As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strQuadrat As String
  Dim strFolder As String
  Dim strFileHeader As String
  Dim lngIndex As Long
  Dim strItem() As String
  Dim pReturn As Collection
  Dim lngArrayIndex As Long
  Dim lngSitesArrayIndex As Long
  Dim lngSitesSpecificArrayIndex As Long
  Dim lngQuadrats() As Long
  Dim strQuadratForSorting As String
  Dim strQuadratSplit() As String
  Dim pQuadrats As New Collection
  Dim strJustQuadrat As String
  Dim lngRunningHighVal As String
  Dim strPlot_2004 As String
  lngRunningHighVal = 999990

  Set pReturn = New Collection
  lngArrayIndex = -1
  lngSitesArrayIndex = -1
  lngSitesSpecificArrayIndex = -1

  Set pPlotToQuadratConversion = New Collection
  Set pQuadratToPlotConversion = New Collection

  Dim pDoneSites As New Collection
  Dim pDoneSitesSpecifics As New Collection

  strFile = MyGeneralOperations.ReadTextFile(strPath)
  strLines = Split(strFile, vbCrLf)
  For lngIndex = 1 To UBound(strLines) ' skip first line of field names
    strLine = Trim(strLines(lngIndex))
    If strLine <> "" Then
      strLineSplit = Split(strLine, vbTab)
      strSite = strLineSplit(0)
      strSiteSpecific = strLineSplit(1)
      strPlot = strLineSplit(2)
      strQuadrat = strLineSplit(4)

      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSites, strSite) Then
        pDoneSites.Add True, strSite
        lngSitesArrayIndex = lngSitesArrayIndex + 1
        ReDim Preserve varSites(lngSitesArrayIndex)
        varSites(lngSitesArrayIndex) = strSite
      End If

      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSitesSpecifics, strSiteSpecific) Then
        pDoneSitesSpecifics.Add True, strSiteSpecific
        lngSitesSpecificArrayIndex = lngSitesSpecificArrayIndex + 1
        ReDim Preserve varSitesSpecific(lngSitesSpecificArrayIndex)
        varSitesSpecific(lngSitesSpecificArrayIndex) = strSiteSpecific
      End If

      If StrComp(Left(strSite, 4), "COC-", vbTextCompare) = 0 Or StrComp(Left(strSite, 16), "Fort Valley COC-", vbTextCompare) = 0 Then
        strFolder = "Woolsey"
        strFileHeader = "COC"
      Else
        strFolder = strSite
        strFileHeader = strSite
      End If

      If strSite = "Big Fill" Then
        strFileHeader = "BF"
      ElseIf strSite = "Black Springs" Then
        strFileHeader = "BS"
      ElseIf strSite = "Fry Park" Then
        strFileHeader = "FP"
      ElseIf strSite = "Reese Tank" Then
        strFileHeader = "RT"
      ElseIf strSite = "Rogers Lake" Then
        strFileHeader = "RL"
      ElseIf strSite = "Wild Bill" Then
        strFileHeader = "WB"
      ElseIf strSite = "FS 9009H" Then
        strFileHeader = "FS_9009H"
      Else
        If strFileHeader <> "COC" Then
          Debug.Print "Check This:  strSite = " & strSite & "..."
          DoEvents
        End If
      End If

      lngArrayIndex = lngArrayIndex + 1
      If strQuadrat = "not yet assigned" Then
        strQuadrat = strQuadrat & " (" & Format(lngArrayIndex, "0") & ")"
      End If
      If strPlot = "not yet assigned" Then
        strPlot = strPlot & " (" & Format(lngArrayIndex, "0") & ")"
      End If

      ReDim strItem(5)
      strItem(0) = strSite
      strItem(1) = strSiteSpecific
      strItem(2) = strPlot
      strItem(3) = strQuadrat
      strItem(4) = strFolder
      strItem(5) = strFileHeader
      pReturn.Add strItem, strQuadrat

      If Not MyGeneralOperations.CheckCollectionForKey(pPlotToQuadratConversion, strPlot) Then
        pPlotToQuadratConversion.Add strQuadrat, strPlot
        Select Case strPlot
          Case "10 / 30710"
            pPlotToQuadratConversion.Add strQuadrat, "30710"
          Case "16 / 30716"
            pPlotToQuadratConversion.Add strQuadrat, "30716"
          Case "18 / 30718"
            pPlotToQuadratConversion.Add strQuadrat, "30718"
          Case "8 / 30708"
            pPlotToQuadratConversion.Add strQuadrat, "30708"
        End Select
      End If
      If Not MyGeneralOperations.CheckCollectionForKey(pQuadratToPlotConversion, strQuadrat) Then
        pQuadratToPlotConversion.Add strPlot, strQuadrat
      End If

      ReDim Preserve lngQuadrats(lngArrayIndex)
      strQuadratForSorting = strQuadrat
      If InStr(1, strQuadratForSorting, "/", vbTextCompare) > 0 Then
        strQuadratSplit = Split(strQuadratForSorting, "/")
        strJustQuadrat = Trim(strQuadratSplit(1))
      Else
        strJustQuadrat = strQuadrat
      End If

      If IsNumeric(strJustQuadrat) Then
        lngQuadrats(lngArrayIndex) = CLng(strJustQuadrat)
        pQuadrats.Add strQuadrat, strJustQuadrat
      Else
        lngRunningHighVal = lngRunningHighVal + 1
        lngQuadrats(lngArrayIndex) = lngRunningHighVal
        pQuadrats.Add strQuadrat, Format(lngRunningHighVal, "0")
      End If
    End If
  Next lngIndex

  QuickSort.LongAscending lngQuadrats, 0, UBound(lngQuadrats)
  ReDim strQuadrats(UBound(lngQuadrats))
  For lngIndex = 0 To UBound(lngQuadrats)
    strQuadrats(lngIndex) = pQuadrats.Item(CStr(lngQuadrats(lngIndex)))
  Next lngIndex

  Set FillQuadratNameColl_Rev = pReturn

ClearMemory:
  Erase strLines
  Erase strLineSplit
  Erase strItem
  Set pReturn = Nothing
  Erase lngQuadrats
  Erase strQuadratSplit
  Set pQuadrats = Nothing

End Function

Public Sub ReturnQuadratVegSoilData(pCollection_To_Fill As Collection, strPlotNames() As String)

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ExcelWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\" & _
      "NAZ_quad_plot_info.xlsx", 0)

  Dim pTable As ITable
  Set pTable = pWS.OpenTable("For_ArcGIS$")

  Set pWS = Nothing
  Set pWSFact = Nothing

  Dim lngIndex As Long
  Dim strReportDim As String
  Dim strReportDeclare As String
  Dim strValueDim As String
  Dim strValueDeclare As String
  Dim strVarName As String
  Dim strCreateArray As String
  Dim strReadArray As String

  strCreateArray = "  varArray = Array("

  For lngIndex = 0 To pTable.Fields.FieldCount - 1
    strVarName = MyGeneralOperations.ReturnTitleCase(pTable.Fields.Field(lngIndex).Name)
    strReportDim = strReportDim & "  dim lng" & strVarName & "Index as Long" & vbCrLf
    strReportDeclare = strReportDeclare & "  lng" & strVarName & "Index = ptable.findfield(""" & strVarName & """)" & vbCrLf
    If StrComp("site", strVarName, vbTextCompare) = 0 Or StrComp("plot", strVarName, vbTextCompare) = 0 Then
      strValueDim = strValueDim & "  dim str" & strVarName & " as string" & vbCrLf
      strValueDeclare = strValueDeclare & "  str" & strVarName & _
          " = trim(cstr(prow.value(lng" & strVarName & "Index)))" & vbCrLf
      strCreateArray = strCreateArray & "str" & strVarName & IIf(lngIndex = pTable.Fields.FieldCount - 1, ")", _
          IIf(lngIndex > 0 And lngIndex Mod 5 = 0, ", _" & vbCrLf & "      ", ", "))
    Else
      strValueDim = strValueDim & "  dim dbl" & strVarName & " as double" & vbCrLf
      strValueDeclare = strValueDeclare & "  dbl" & strVarName & _
          " = cdbl(prow.value(lng" & strVarName & "Index))" & vbCrLf
      strCreateArray = strCreateArray & "dbl" & strVarName & IIf(lngIndex = pTable.Fields.FieldCount - 1, ")", _
          IIf(lngIndex > 0 And lngIndex Mod 5 = 0, ", _" & vbCrLf & "      ", ", "))
    End If
  Next lngIndex

  Set pCollection_To_Fill = New Collection

  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim varArray() As Variant

  Dim lngSiteIndex As Long
  Dim lngPlotIndex As Long
  Dim lngPipo_density_trees_haIndex As Long
  Dim lngTotal_ba_m2_haIndex As Long
  Dim lngPipo_ba_m2_haIndex As Long
  Dim lngQuga_ba_m2_haIndex As Long
  Dim lngJumo_ba_m2_haIndex As Long
  Dim lngJude_ba_m2_haIndex As Long
  Dim lngCanopy_cover_spherical_percIndex As Long
  Dim lngCanopy_cover_vertical_percIndex As Long
  Dim lngCanopy_cover_avg_percIndex As Long
  Dim lngO_horizon_depth_cmIndex As Long
  Dim lngSoil_organic_matter_percIndex As Long
  Dim lngSand_percIndex As Long
  Dim lngSilt_percIndex As Long
  Dim lngClay_percIndex As Long
  Dim lngPhIndex As Long
  Dim lngSoil_total_p_percIndex As Long
  Dim lngSoil_total_c_percIndex As Long
  Dim lngSoil_total_n_percIndex As Long

  lngSiteIndex = pTable.FindField("Site")
  lngPlotIndex = pTable.FindField("Plot")
  lngPipo_density_trees_haIndex = pTable.FindField("Pipo_density_trees_ha")
  lngTotal_ba_m2_haIndex = pTable.FindField("Total_ba_m2_ha")
  lngPipo_ba_m2_haIndex = pTable.FindField("Pipo_ba_m2_ha")
  lngQuga_ba_m2_haIndex = pTable.FindField("Quga_ba_m2_ha")
  lngJumo_ba_m2_haIndex = pTable.FindField("Jumo_ba_m2_ha")
  lngJude_ba_m2_haIndex = pTable.FindField("Jude_ba_m2_ha")
  lngCanopy_cover_spherical_percIndex = pTable.FindField("Canopy_cover_spherical_perc")
  lngCanopy_cover_vertical_percIndex = pTable.FindField("Canopy_cover_vertical_perc")
  lngCanopy_cover_avg_percIndex = pTable.FindField("Canopy_cover_avg_perc")
  lngO_horizon_depth_cmIndex = pTable.FindField("O_horizon_depth_cm")
  lngSoil_organic_matter_percIndex = pTable.FindField("Soil_organic_matter_perc")
  lngSand_percIndex = pTable.FindField("Sand_perc")
  lngSilt_percIndex = pTable.FindField("Silt_perc")
  lngClay_percIndex = pTable.FindField("Clay_perc")
  lngPhIndex = pTable.FindField("Ph")
  lngSoil_total_p_percIndex = pTable.FindField("Soil_total_p_perc")
  lngSoil_total_c_percIndex = pTable.FindField("Soil_total_c_perc")
  lngSoil_total_n_percIndex = pTable.FindField("Soil_total_n_perc")

  Dim strSite As String
  Dim strPlot As String
  Dim dblPipo_density_trees_ha As Double
  Dim dblTotal_ba_m2_ha As Double
  Dim dblPipo_ba_m2_ha As Double
  Dim dblQuga_ba_m2_ha As Double
  Dim dblJumo_ba_m2_ha As Double
  Dim dblJude_ba_m2_ha As Double
  Dim dblCanopy_cover_spherical_perc As Double
  Dim dblCanopy_cover_vertical_perc As Double
  Dim dblCanopy_cover_avg_perc As Double
  Dim dblO_horizon_depth_cm As Double
  Dim dblSoil_organic_matter_perc As Double
  Dim dblSand_perc As Double
  Dim dblSilt_perc As Double
  Dim dblClay_perc As Double
  Dim dblPh As Double
  Dim dblSoil_total_p_perc As Double
  Dim dblSoil_total_c_perc As Double
  Dim dblSoil_total_n_perc As Double

  Dim varVal As Variant

  Dim lngCounter As Long

  lngCounter = -1
  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    varVal = pRow.Value(lngPlotIndex)
    If Not IsNull(varVal) Then
      strPlot = Trim(CStr(varVal))

      strSite = Trim(CStr(pRow.Value(lngSiteIndex)))
      dblPipo_density_trees_ha = CDbl(pRow.Value(lngPipo_density_trees_haIndex))
      dblTotal_ba_m2_ha = CDbl(pRow.Value(lngTotal_ba_m2_haIndex))
      dblPipo_ba_m2_ha = CDbl(pRow.Value(lngPipo_ba_m2_haIndex))
      dblQuga_ba_m2_ha = CDbl(pRow.Value(lngQuga_ba_m2_haIndex))
      dblJumo_ba_m2_ha = CDbl(pRow.Value(lngJumo_ba_m2_haIndex))
      dblJude_ba_m2_ha = CDbl(pRow.Value(lngJude_ba_m2_haIndex))
      dblCanopy_cover_spherical_perc = CDbl(pRow.Value(lngCanopy_cover_spherical_percIndex))
      dblCanopy_cover_vertical_perc = CDbl(pRow.Value(lngCanopy_cover_vertical_percIndex))
      dblCanopy_cover_avg_perc = CDbl(pRow.Value(lngCanopy_cover_avg_percIndex))
      dblO_horizon_depth_cm = CDbl(pRow.Value(lngO_horizon_depth_cmIndex))
      dblSoil_organic_matter_perc = CDbl(pRow.Value(lngSoil_organic_matter_percIndex))
      dblSand_perc = CDbl(pRow.Value(lngSand_percIndex))
      dblSilt_perc = CDbl(pRow.Value(lngSilt_percIndex))
      dblClay_perc = CDbl(pRow.Value(lngClay_percIndex))
      dblPh = CDbl(pRow.Value(lngPhIndex))
      dblSoil_total_p_perc = CDbl(pRow.Value(lngSoil_total_p_percIndex))
      dblSoil_total_c_perc = CDbl(pRow.Value(lngSoil_total_c_percIndex))
      dblSoil_total_n_perc = CDbl(pRow.Value(lngSoil_total_n_percIndex))

      lngCounter = lngCounter + 1
      ReDim Preserve strPlotNames(lngCounter)
      strPlotNames(lngCounter) = strPlot
      varArray = Array(strSite, strPlot, dblPipo_density_trees_ha, dblTotal_ba_m2_ha, dblPipo_ba_m2_ha, dblQuga_ba_m2_ha, _
          dblJumo_ba_m2_ha, dblJude_ba_m2_ha, dblCanopy_cover_spherical_perc, dblCanopy_cover_vertical_perc, dblCanopy_cover_avg_perc, _
          dblO_horizon_depth_cm, dblSoil_organic_matter_perc, dblSand_perc, dblSilt_perc, dblClay_perc, _
          dblPh, dblSoil_total_p_perc, dblSoil_total_c_perc, dblSoil_total_n_perc)
      pCollection_To_Fill.Add varArray, strPlot

    End If
    Set pRow = pCursor.NextRow
  Loop

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  varVal = Null

End Sub

Public Sub ReturnQuadratCoordsAndNames(pCollection_To_Fill As Collection, strPlotNames() As String)

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ExcelWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\From_Margaret_Dec_2017_Jan_2018\" & _
      "Hill-WildBill and FVEF_Quadrat Locations by Site_NAD83_UTMs_Fall Sampling_2016_2017.xlsx", 0)

  Dim pTable As ITable
  Set pTable = pWS.OpenTable("For_ArcGIS$")
  Set pWS = Nothing
  Set pWSFact = Nothing

  Dim lngIndex As Long

  Set pCollection_To_Fill = New Collection

  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngQuadIndex As Long
  Dim lngEastingIndex As Long
  Dim lngNorthingIndex As Long
  Dim lngNameIndex As Long
  Dim lngNotesIndex As Long
  Dim lngCommentIndex As Long
  Dim lng2016Index As Long
  Dim lng2017Index As Long

  lngQuadIndex = pTable.FindField("Quadrat_or_Plot")
  lngEastingIndex = pTable.FindField("NAD_83_UTM_E")
  lngNorthingIndex = pTable.FindField("NAD_83_UTM_N")
  lngNameIndex = pTable.FindField("Name")
  lngCommentIndex = pTable.FindField("Comment")
  lngNotesIndex = pTable.FindField("Notes")
  lng2016Index = pTable.FindField("Surveyed_2016")
  lng2017Index = pTable.FindField("Surveyed_2017")

  Dim strQuad As String
  Dim dblEasting As Double
  Dim dblNorthing As Double
  Dim strName As String
  Dim strComment As String
  Dim varVal As Variant
  Dim strNote As String
  Dim str2016 As String
  Dim str2017 As String

  Dim lngCounter As Long

  lngCounter = -1
  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    varVal = pRow.Value(lngQuadIndex)
    If Not IsNull(varVal) Then
      strQuad = Trim(CStr(varVal))
      strQuad = Replace(strQuad, "*", "")
      strQuad = Trim(strQuad)
      If strQuad = "6" Then
        DoEvents
      End If
      dblEasting = CDbl(pRow.Value(lngEastingIndex))
      dblNorthing = CDbl(pRow.Value(lngNorthingIndex))
      strName = Trim(CStr(pRow.Value(lngNameIndex)))
      varVal = pRow.Value(lngCommentIndex)
      If IsNull(varVal) Then
        strComment = ""
      Else
        strComment = Trim(CStr(varVal))
      End If
      varVal = pRow.Value(lngNotesIndex)
      If IsNull(varVal) Then
        strNote = ""
      Else
        strNote = Trim(CStr(varVal))
      End If
      varVal = pRow.Value(lng2016Index)
      If IsNull(varVal) Then
        str2016 = ""
      Else
        str2016 = Trim(CStr(varVal))
      End If
      varVal = pRow.Value(lng2017Index)
      If IsNull(varVal) Then
        str2017 = ""
      Else
        str2017 = Trim(CStr(varVal))
      End If

      lngCounter = lngCounter + 1
      ReDim Preserve strPlotNames(lngCounter)
      strPlotNames(lngCounter) = strQuad
      pCollection_To_Fill.Add Array(dblEasting, dblNorthing, strName, strComment, strNote, str2016, str2017), strQuad
    End If
    Set pRow = pCursor.NextRow
  Loop

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  varVal = Null

End Sub

Public Sub ReturnVegDataElevAndNames(pVegDataAndElevations As Collection, strPlotNames() As String, pPlotLocColl As Collection)

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New TextFileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Data_to_include_in_publication\", 0)

  Dim pTable As ITable
  Set pTable = pWS.OpenTable("Plot_Area_Data.csv")

  Set pWS = Nothing
  Set pWSFact = Nothing
  Dim pRastWSEx As IRasterWorkspaceEx
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pRastWSEx = pWSFact.OpenFromFile("D:\GIS_Data\DEM_Stuff\Full_DEM_Data.gdb", 0)
  Dim pRDataset As IRasterDataset2
  Set pRDataset = pRastWSEx.OpenRasterDataset("All_US_NoNull")
  Dim pRaster As IRaster2
  Set pRaster = pRDataset.CreateFullRaster

  Dim lngIndex As Long

  Set pVegDataAndElevations = New Collection

  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngQuadIndex As Long
  Dim lngSiteIndex As Long
  Dim lngAspectIndex As Long
  Dim lngSlopeIndex As Long
  Dim lngCanopyCoverIndex As Long
  Dim lngBasalAreaIndex As Long
  Dim lngAltBasalAreaIndex As Long
  Dim lngSoilIndex As Long

  lngSiteIndex = pTable.FindField("Site")
  lngQuadIndex = pTable.FindField("Plot")
  lngAspectIndex = pTable.FindField("Aspect")
  lngSlopeIndex = pTable.FindField("Perc_Slope")
  lngCanopyCoverIndex = pTable.FindField("Avg_Canopy_Cover")
  lngBasalAreaIndex = pTable.FindField("BA_per_ha")
  lngAltBasalAreaIndex = pTable.FindField("BA_m2_per_ha_1998")
  lngSoilIndex = pTable.FindField("Soil")

  Dim strQuad As String
  Dim strSite As String
  Dim dblAspect As Double
  Dim dblSlope As Double
  Dim dblCanopyCover As Double
  Dim varBA As Variant
  Dim varVal As Variant
  Dim varSlope As Variant
  Dim strSoil As String

  Dim varArray() As Variant
  Dim dblNorthing As Double
  Dim dblEasting As Double
  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pRDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pPrjSpRef As IProjectedCoordinateSystem
  Set pPrjSpRef = pSpRef
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001
  Dim pPoint As IPoint
  Dim dblElev As Double
  Dim lngYear As Long
  Dim pGeoPoint As IPoint
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  Dim pClone As IClone

  Dim lngCounter As Long

  lngCounter = -1
  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    varVal = pRow.Value(lngQuadIndex)
    If Not IsNull(varVal) Then
      strQuad = Trim(CStr(varVal))
      strQuad = Replace(strQuad, "*", "")
      strQuad = Trim(strQuad)
      If strQuad = "6" Then
        DoEvents
      End If

      varArray = pPlotLocColl.Item(strQuad)
      dblEasting = varArray(0)
      dblNorthing = varArray(1)
      Set pPoint = New Point
      pPoint.PutCoords dblEasting, dblNorthing
      Set pPoint.SpatialReference = pSpRef
      pPoint.Project pGeoDataset.SpatialReference
      dblElev = GridFunctions.CellValue4CellInterp(pPoint, pRaster)

      Set pClone = pPoint
      Set pGeoPoint = pClone.Clone
      pGeoPoint.Project pPrjSpRef.GeographicCoordinateSystem
      dblLongitude = pGeoPoint.x
      dblLatitude = pGeoPoint.Y
      strSite = pRow.Value(lngSiteIndex)

      If strSite = "Big Fill" Or strSite = "Black Springs" Or strSite = "Fry Park" Or strSite = "Reese Tank" Then
        lngYear = 2004
      ElseIf strSite = "Rogers Lake" Or strSite = "Wild Bill" Then
        lngYear = 2006
      Else
        lngYear = 2002 ' changed from 1998 based on 7/27/2021 email discussion with Margaret and Jon Bakker
      End If

      dblAspect = CDbl(pRow.Value(lngAspectIndex))
      If IsNull(pRow.Value(lngSlopeIndex)) Then
        varSlope = Null
      Else
        varSlope = CDbl(pRow.Value(lngSlopeIndex))
      End If
      dblCanopyCover = CDbl(pRow.Value(lngCanopyCoverIndex))

      If IsNull(pRow.Value(lngBasalAreaIndex)) Then
        If IsNull(pRow.Value(lngAltBasalAreaIndex)) Then
          varBA = Null
        Else
          varBA = CDbl(pRow.Value(lngAltBasalAreaIndex))
        End If
      Else
        varBA = CDbl(pRow.Value(lngBasalAreaIndex))
      End If

      strSoil = Trim(pRow.Value(lngSoilIndex))

      lngCounter = lngCounter + 1
      ReDim Preserve strPlotNames(lngCounter)
      strPlotNames(lngCounter) = strQuad
      pVegDataAndElevations.Add Array(strSite, dblElev, dblAspect, varSlope, dblCanopyCover, varBA, strSoil, pPoint, _
          dblNorthing, dblEasting, lngYear, dblLatitude, dblLongitude, pGeoPoint), strQuad
    End If
    Set pRow = pCursor.NextRow
  Loop

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pRastWSEx = Nothing
  Set pRDataset = Nothing
  Set pRaster = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  varBA = Null
  varVal = Null
  varSlope = Null
  Erase varArray
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pPoint = Nothing

End Sub

Public Function ReturnQuadratData(pPlotLocColl As Collection) As Collection

  Dim strPath As String
  strPath = "D:\arcGIS_stuff\consultation\Margaret_Moore\From_Margaret_Dec_2017_Jan_2018\" & _
      "Hill-Wild Bill_Old and New Quadrat Numbers by Site_2016_mod_Feb_2018.txt"

  Dim strFile As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strQuadrat As String
  Dim strFolder As String
  Dim strFileHeader As String
  Dim lngIndex As Long
  Dim strItem() As String
  Dim pReturn As Collection
  Dim lngArrayIndex As Long
  Dim lngSitesArrayIndex As Long
  Dim lngSitesSpecificArrayIndex As Long
  Dim lngQuadrats() As Long
  Dim strQuadratForSorting As String
  Dim strQuadratSplit() As String
  Dim pQuadrats As New Collection
  Dim strJustQuadrat As String
  Dim lngRunningHighVal As String
  Dim strPlot_2004 As String
  lngRunningHighVal = 999990

  Dim pReturnColl As New Collection
  Dim dblEasting As Double
  Dim dblNorthing As Double
  Dim strName As String
  Dim strComment As String
  Dim strNote As String
  Dim strAKA As String
  Dim strExclosure As String
  Dim strComment2 As String

  Set pReturn = New Collection
  lngArrayIndex = -1
  lngSitesArrayIndex = -1
  lngSitesSpecificArrayIndex = -1
  Dim varData() As Variant

  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pRastWSEx As IRasterWorkspaceEx
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pRastWSEx = pWSFact.OpenFromFile("D:\GIS_Data\DEM_Stuff\Full_DEM_Data.gdb", 0)
  Dim pRDataset As IRasterDataset2
  Set pRDataset = pRastWSEx.OpenRasterDataset("All_US_NoNull")
  Dim pRaster As IRaster2
  Set pRaster = pRDataset.CreateFullRaster

  Dim dblElev As Double
  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pRDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pPrjSpRef As IProjectedCoordinateSystem
  Set pPrjSpRef = pSpRef
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001
  Dim pPoint As IPoint
  Dim pGeoPoint As IPoint
  Dim pClone As IClone

  strFile = MyGeneralOperations.ReadTextFile(strPath)
  strLines = Split(strFile, vbCrLf)
  For lngIndex = 1 To UBound(strLines) ' skip first line of field names
    strLine = Trim(strLines(lngIndex))
    If strLine <> "" Then
      strLineSplit = Split(strLine, vbTab)
      strSite = strLineSplit(0)
      strSiteSpecific = strLineSplit(1)
      strPlot = strLineSplit(2)
      strQuadrat = strLineSplit(4)
      strAKA = "Q" & strQuadrat
      strExclosure = strLineSplit(5)
      strComment2 = strLineSplit(16)

      varData = pPlotLocColl.Item(strPlot)
      dblEasting = varData(0)
      dblNorthing = varData(1)
      strName = varData(2)
      strComment = varData(3)
      If InStr(1, strComment, "?") = 0 Then strComment = ""
      strNote = varData(4)

      Set pPoint = New Point
      pPoint.PutCoords dblEasting, dblNorthing
      Set pPoint.SpatialReference = pSpRef
      Set pClone = pPoint
      Set pGeoPoint = pClone.Clone
      pGeoPoint.Project pPrjSpRef.GeographicCoordinateSystem
      pPoint.Project pGeoDataset.SpatialReference
      dblElev = GridFunctions.CellValue4CellInterp(pPoint, pRaster)

      pReturnColl.Add Array(dblEasting, dblNorthing, strSite, strName, strAKA, strExclosure, strNote, _
          strComment, strComment2, dblElev, pGeoPoint), strPlot
    End If
  Next lngIndex

  Set ReturnQuadratData = pReturnColl

End Function

Public Function FillPolygonWithPointArray(pPolygon As IPolygon, dblSeparationDist As Double) As Variant()

  Dim varReturn() As Variant
  Dim booFoundPoint As Boolean
  Dim pArea As IArea
  Dim pPoint As IPoint
  varReturn = CreatePointsTriangularPattern(pPolygon, dblSeparationDist, booFoundPoint)
  If booFoundPoint = False Then
    Set pArea = pPolygon
    Set pPoint = pArea.LabelPoint
    ReDim varReturn(0)
    Set varReturn(0) = pPoint
  End If

  FillPolygonWithPointArray = varReturn

End Function

Public Function ReturnCircleClippedToQuadrat(pPoint As IPoint, dblRadius As Double, lngVertexCount As Long, _
    pQuadrat As IPolygon, Optional pCoverPolygon As IPolygon = Nothing)

  Dim pPolygon As IPolygon
  Dim pTopoOp As ITopologicalOperator
  Dim pSpRefRes As ISpatialReferenceResolution
  Dim pSpRef As ISpatialReference
  Set pSpRef = pQuadrat.SpatialReference
  Set pSpRefRes = pSpRef
  pSpRefRes.XYResolution(True) = 0.00001
  Set pPoint.SpatialReference = pSpRef
  Set pPolygon = MyGeometricOperations.CreateCircleAroundPoint(pPoint, dblRadius, lngVertexCount)
  Set pPolygon.SpatialReference = pQuadrat.SpatialReference

  Dim pTempPoly As IPolygon
  If Not pCoverPolygon Is Nothing Then
    Set pTopoOp = pCoverPolygon
    Set pTempPoly = pTopoOp.Intersect(pPolygon, pPolygon.Dimension)
  Else
    Set pTempPoly = pPolygon
  End If

  Set pTopoOp = pQuadrat
  Set ReturnCircleClippedToQuadrat = pTopoOp.Intersect(pTempPoly, pPolygon.Dimension)

  Set pPolygon = Nothing
  Set pTopoOp = Nothing

End Function

Public Function ReturnQuadratPolygon(pSpRef As ISpatialReference) As IPolygon

  Dim pQuadrat As IPolygon
  Set pQuadrat = New Polygon
  Set pQuadrat.SpatialReference = pSpRef
  Dim pQPtColl As IPointCollection
  Set pQPtColl = pQuadrat
  Dim pQPoint As IPoint
  Set pQPoint = New Point
  Set pQPoint.SpatialReference = pSpRef
  pQPoint.PutCoords 0, 0
  pQPtColl.AddPoint pQPoint
  Set pQPoint = New Point
  Set pQPoint.SpatialReference = pSpRef
  pQPoint.PutCoords 0, 1
  pQPtColl.AddPoint pQPoint
  Set pQPoint = New Point
  Set pQPoint.SpatialReference = pSpRef
  pQPoint.PutCoords 1, 1
  pQPtColl.AddPoint pQPoint
  Set pQPoint = New Point
  Set pQPoint.SpatialReference = pSpRef
  pQPoint.PutCoords 1, 0
  pQPtColl.AddPoint pQPoint
  pQuadrat.Close

  Set ReturnQuadratPolygon = pQuadrat

  Set pQuadrat = Nothing
  Set pQPtColl = Nothing
  Set pQPoint = Nothing
End Function

Public Function CreatePointsTriangularPattern(pPolygon As IPolygon, dblSeparationDist As Double, _
  booFoundPoint As Boolean) As Variant()

  Dim pPoint As IPoint
  Dim pSpRef As ISpatialReference
  Dim pClone As IClone
  Dim pTopoOp As ITopologicalOperator

  Dim pExpandedPolygon As IPolygon
  Set pClone = pPolygon
  Set pExpandedPolygon = pClone.Clone
  Set pTopoOp = pExpandedPolygon
  Set pExpandedPolygon = pTopoOp.Buffer(dblSeparationDist * 2)

  Set pSpRef = pPolygon.SpatialReference
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument

  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pSpRef
  pSpRefRes.XYResolution(True) = 0.00001

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pPolygon.Envelope

  Dim varPolygon() As Variant
  varPolygon = MyGeometricOperations.ReturnPolygonRingsAsDoubleArray(pPolygon)

  Dim varReturn() As Variant
  Dim lngReturnCounter As Long
  lngReturnCounter = -1

  Dim dblOrient As Double
  Randomize
  dblOrient = Rnd() * 30
  booFoundPoint = False

  Dim theExtent As IEnvelope
  Set pClone = pEnvelope
  Set theExtent = pClone.Clone
  theExtent.Expand dblSeparationDist * 2, dblSeparationDist * 2, False
  Dim theStartPoint As IPoint
  Set theStartPoint = New Point
  Set theStartPoint.SpatialReference = pSpRef
  theStartPoint.PutCoords theExtent.XMin, theExtent.YMin

  Do While dblOrient > 60
    dblOrient = dblOrient - 60
  Loop

  Dim strOrientForReport As String
  strOrientForReport = CStr(dblOrient)

  If dblOrient > 30 Then dblOrient = dblOrient - 60

  Dim theGeneralWidth As Double
  Dim theGeneralRight As Double
  Dim theGeneralTop As Double
  Dim theGeneralTopLimit As Double
  Dim theGeneralLeft As Double
  Dim theGeneralBottom As Double
  Dim theGeneralBottomLimit As Double
  Dim theOrientTan As Double

  theGeneralWidth = theExtent.Width / Cos(MyGeometricOperations.AsRadians(dblOrient))
  theGeneralRight = theExtent.XMax
  theGeneralTop = theExtent.YMax
  theGeneralTopLimit = theGeneralTop + dblSeparationDist
  theGeneralLeft = theExtent.XMin
  theGeneralBottom = theExtent.YMin
  theGeneralBottomLimit = theGeneralBottom - dblSeparationDist
  theOrientTan = Tan(MyGeometricOperations.AsRadians(dblOrient))

  Dim theVerticalAdjust As Double
  Dim theGeneralHeight As Double
  Dim theHorizontalAdjust As Double

  theVerticalAdjust = theOrientTan * theGeneralWidth
  theGeneralHeight = theExtent.Height + Abs(theVerticalAdjust) + dblSeparationDist
  theHorizontalAdjust = theOrientTan * theGeneralHeight

  If dblOrient < 0 Then
    theStartPoint.PutCoords theGeneralLeft + theHorizontalAdjust, theGeneralBottom
  Else
    theStartPoint.PutCoords theGeneralLeft, theGeneralBottom - theVerticalAdjust
  End If

  Dim theNextBearing As Double
  theNextBearing = 30 - dblOrient
  If theNextBearing < 0 Then theNextBearing = theNextBearing + 360
  Dim theNextBearingBack As Double
  theNextBearingBack = theNextBearing - 60
  If theNextBearingBack < 0 Then theNextBearingBack = theNextBearingBack + 360

  Dim theNextPoint As IPoint
  Call MyGeometricOperations.CalcPointLine(theStartPoint, dblSeparationDist, theNextBearing, theNextPoint)

  Dim theDelta As IPoint

  Set theDelta = MyGeometricOperations.PointSubtract(theNextPoint, theStartPoint)

  Dim theNextRowDeltaForward As IPoint
  Set theNextRowDeltaForward = theDelta
  Dim theNextRowDeltaBack As IPoint
  Dim thePointDeltaEdge As IPoint
  Dim theAddStep As Long
  Dim theCloner As IClone
  Dim theCurrentPoint As IPoint
  Dim thePointList As Collection
  Dim thePoint As IPoint
  Dim theCurrentX As Double

  Dim theTempPoint As IPoint

  Set theTempPoint = New Point
  Call MyGeometricOperations.CalcPointLine(theStartPoint, dblSeparationDist, theNextBearingBack, theTempPoint)
  Set theNextRowDeltaBack = MyGeometricOperations.PointSubtract(theTempPoint, theStartPoint)

  Set theTempPoint = New Point
  Call MyGeometricOperations.CalcPointLine(theStartPoint, dblSeparationDist, 90 - dblOrient, theTempPoint)
  Set thePointDeltaEdge = MyGeometricOperations.PointSubtract(theTempPoint, theStartPoint)

  Dim theDeltaY As Double
  theDeltaY = theDelta.Y
  Dim theStartY As Double
  theStartY = theStartPoint.Y

  Dim theRowOrigin As IPoint
  Set theRowOrigin = theStartPoint
  Set theRowOrigin.SpatialReference = pSpRef

  Dim theTotal1 As Long
  theTotal1 = Int(theGeneralHeight / theDeltaY) + 1

  Dim FoundAPoint As Boolean
  FoundAPoint = True

  Dim theCurrentY As Double
  theCurrentY = theStartY

  Dim aNum As Long
  aNum = 0
  Dim aNum2 As Long

  Dim theDateString As String
  Dim theElapsedTimeString As String
  Dim ShouldContinue As Boolean

  Dim dblMaxX As Double
  Dim dblMinX As Double
  Dim dblMaxY As Double
  Dim dblMinY As Double
  Dim dblPointX As Double
  Dim dblPointY As Double
  dblMaxX = theExtent.XMax
  dblMinX = theExtent.XMin
  dblMaxY = theExtent.YMax
  dblMinY = theExtent.YMin

  Dim theEstimateArea As Double
  Dim pArea As IArea
  Set pArea = theExtent
  theEstimateArea = pArea.Area

  Dim thePropString As String

  Do While FoundAPoint

    aNum = aNum + 1
    If aNum Mod 2 = 0 Then
      Set theRowOrigin = MyGeometricOperations.PointAdd(theRowOrigin, theNextRowDeltaForward)
      theAddStep = 0
    Else
      Set theRowOrigin = MyGeometricOperations.PointAdd(theRowOrigin, theNextRowDeltaBack)
      theAddStep = 1
    End If

    FoundAPoint = theRowOrigin.Y < theGeneralTopLimit

    Set theCloner = theRowOrigin
    Set theCurrentPoint = theCloner.Clone

    aNum2 = 0
    theCurrentX = theCurrentPoint.x
    Do While theCurrentX < theGeneralRight
      aNum2 = aNum2 + 1

      Set theCurrentPoint = MyGeometricOperations.PointAdd(theCurrentPoint, thePointDeltaEdge)
      theCurrentX = theCurrentPoint.x
      If dblOrient > 0 And (theCurrentPoint.Y > theGeneralTopLimit) Then
        Exit Do
      ElseIf dblOrient < 0 And (theCurrentPoint.Y < theGeneralBottomLimit) Then
        Exit Do
      End If

      Set theCloner = theCurrentPoint
      Set thePoint = theCloner.Clone
      Set thePoint.SpatialReference = pSpRef
      dblPointX = thePoint.x
      dblPointY = thePoint.Y

      If dblPointX >= dblMinX And dblPointX <= dblMaxX And dblPointY >= dblMinY And dblPointY <= dblMaxY Then

        If MyGeometricOperations.PointInPoly_Winding(dblPointX, dblPointY, varPolygon) Then
          lngReturnCounter = lngReturnCounter + 1
          ReDim Preserve varReturn(lngReturnCounter)
          Set varReturn(lngReturnCounter) = thePoint
          booFoundPoint = True

        End If

      End If
      If thePoint.Y < theGeneralTopLimit Then FoundAPoint = True

    Loop
  Loop

  CreatePointsTriangularPattern = varReturn

ClearMemory:
  Set pPoint = Nothing
  Set pSpRef = Nothing
  Set pClone = Nothing
  Set pTopoOp = Nothing
  Set pExpandedPolygon = Nothing
  Set pMxDoc = Nothing
  Set pSpRefRes = Nothing
  Set pEnvelope = Nothing
  Erase varPolygon
  Erase varReturn
  Set theExtent = Nothing
  Set theStartPoint = Nothing
  Set theNextPoint = Nothing
  Set theDelta = Nothing
  Set theNextRowDeltaForward = Nothing
  Set theNextRowDeltaBack = Nothing
  Set thePointDeltaEdge = Nothing
  Set theCloner = Nothing
  Set theCurrentPoint = Nothing
  Set thePointList = Nothing
  Set thePoint = Nothing
  Set theTempPoint = Nothing
  Set theRowOrigin = Nothing
  Set pArea = Nothing

End Function

Public Sub Metadata_pNewFClass(pMxDoc As IMxDocument, pNewFClass As IDataset, _
    strAbstract As String, strPurpose As String)

  Dim pDataset As IDataset
  Dim pPropSet As IPropertySet
  Dim strResponse As String
  Dim booSucceeded As Boolean
  Dim lngIndex As Long
  Dim pTable As ITable

  Dim pKeyWords As esriSystem.IStringArray
  Dim pIncludeThemeKeys As esriSystem.IStringArray
  Dim pIncludeSearchKeys As esriSystem.IStringArray
  Dim pIncludeDescKeys As esriSystem.IStringArray
  Dim pIncludeStratKeys As esriSystem.IStringArray
  Dim pIncludeThemeSlashThemekeys As esriSystem.IStringArray
  Dim pIncludePlaceKeys As esriSystem.IStringArray
  Dim pIncludeTemporalKeys As esriSystem.IStringArray
  Dim pCombinedKeyWords As esriSystem.IStringArray

  Dim strFormatVersion As String
  Dim lngVersion As Long
  Dim strDescription As String
  Dim strName As String
  Dim datCreated As Date
  Dim datPublished As Date
  Dim strUseLimitations As String
  strUseLimitations = "This dataset is intended for research, planning, and conservation purposes. " & _
        "Contact Dr. Margaret Moore, Northern Arizona University School of Forestry for information regarding the use of these data."
  Dim strCredits As String
  strCredits = "Northern Arizona University School of Forestry"

  Set pDataset = pNewFClass
  Set pTable = pNewFClass

  strResponse = Metadata_Functions.SynchronizeMetadataPropSet(pDataset)

  strResponse = Metadata_Functions.SetMetadataAbstract(pDataset, strAbstract)

  strResponse = Metadata_Functions.SetMetadataPurpose(pDataset, strPurpose)

  strResponse = Metadata_Functions.SetMetadataCredits(pDataset, strCredits)

  strResponse = Metadata_Functions.AddMetadataUseLimitations(pDataset, strUseLimitations)

  Set pIncludeThemeKeys = New esriSystem.strArray
  Set pIncludeSearchKeys = New esriSystem.strArray
  Set pIncludeDescKeys = New esriSystem.strArray
  Set pIncludeStratKeys = New esriSystem.strArray
  Set pIncludeThemeSlashThemekeys = New esriSystem.strArray
  Set pIncludePlaceKeys = New esriSystem.strArray
  Set pIncludeTemporalKeys = New esriSystem.strArray

  Set pCombinedKeyWords = Metadata_Functions.ReturnExistingMetadataKeyWords(pDataset, _
      pKeyWords, booSucceeded, pIncludeThemeKeys, pIncludeSearchKeys, pIncludeDescKeys, pIncludeStratKeys, _
      pIncludeThemeSlashThemekeys, _
      pIncludePlaceKeys, pIncludeTemporalKeys)

  pIncludeThemeKeys.Add "Northern Arizona University"
  pIncludeThemeKeys.Add "NAU"
  pIncludeThemeKeys.Add "Climate"
  pIncludeThemeKeys.Add "Competition"
  pIncludeThemeKeys.Add "Demography"
  pIncludeThemeKeys.Add "GIS"
  pIncludeThemeKeys.Add "Geographic Information Systems"
  pIncludeThemeKeys.Add "Grassland"
  pIncludeThemeKeys.Add "Plant Community"

  For lngIndex = 0 To pIncludeThemeKeys.Count - 1
    pIncludeSearchKeys.Add pIncludeThemeKeys.Element(lngIndex)
    pIncludeDescKeys.Add pIncludeThemeKeys.Element(lngIndex)
  Next lngIndex

  strResponse = Metadata_Functions.SetMetadataKeyWords(pDataset, pIncludeThemeKeys, pIncludeSearchKeys, _
        pIncludeDescKeys, pIncludeStratKeys, pIncludeThemeSlashThemekeys, pIncludePlaceKeys, pIncludeTemporalKeys)

  strDescription = "Organized feature classes, corrected misspelled or revised species names and converted " & _
      "misclassified species from Cover to Density or vice versa.  Placed data in real-world coordinates."
  strName = aml_func_mod.GetTheUserName
  strResponse = Metadata_Functions.AddNewLineageStep(pDataset, strDescription, Now, JenMetadata_Processor, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)

  strResponse = Metadata_Functions.AddNewGeoProcStep(pDataset, "NOTE:  This is not Python code! " & _
     "Data processing on Quadrat Shapefiles...", _
     "Custom VBA functions in MXD 'Analyze_Shapefiles_for_VM.mxd'", Now, _
     "ThisDocument_for_VM/ConvertPointShapefiles and ../ReviseShapefiles", False)

  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_Processor, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_Author, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)

  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_PointOfContact, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)
  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_PointOfContact, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)
  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_Processor, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)
  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_Custodian, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)

  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_PointOfContact, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)
  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_PointOfContact, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)
  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_Processor, _
    "Jeff Jenness", False, "Jenness Enterprises", , "(928) 526-4139", "3020 N. Schevene Blvd.", "Flagstaff", "AZ", _
    "86004", "USA", , JenMetadata_Postal)
  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_Custodian, _
    "Margaret Moore", False, "Northern Arizona University School of Forestry", "Professor", "(928) 523-3031", _
    "200 E. Pine Knoll Drive", "Flagstaff", "AZ", _
    "86011", "USA", "Margaret.Moore@nau.edu", JenMetadata_both)

  datCreated = Now
  datPublished = Now
  strResponse = Metadata_Functions.AddCitationDates(pDataset, datCreated, datPublished)

  strResponse = Metadata_Functions.AddResourceDetailsStatus(pDataset, JenMetadata_Ongoing)

  strResponse = Metadata_Functions.AddResourceMaintenance(pDataset, JenMetadata_Maint_Daily)

  If pTable.FindField("SPCODE") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "SPCODE", _
      "Species Code", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("SP_CODE") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "SP_CODE", _
      "Species Code", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("SP_CPDE") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "SP_CPDE", _
      "Species Code", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("SPP_") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "SPP_", _
      "Species Code", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("SPP") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "SPP", _
      "Species Code", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("SP") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "SP", _
      "Species Code", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Species") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Species", _
      "Genus and Species", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Seedling") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Seedling", _
      "Whether observed plant was a seedling:  Yes or No.  Cannot reliably track herbaceous seedlings, so this field is only used to track ponderosa pine seedlings", _
      "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Coords_x1") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Coords_x1", _
      "X-Coordinate of observed plant species", "Northern Arizona University", "", "", _
      "", "Meter, in custom local coordinate system", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Coords_x2") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Coords_x2", _
      "Y-Coordinate of observed plant species", "Northern Arizona University", "", "", _
      "", "Meter, in custom local coordinate system", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("x") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "x", _
      "X-Coordinate of representative point within polygon of observed plant species", "Northern Arizona University", "", "", _
      "", "Meter, in custom local coordinate system", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("y") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "y", _
      "Y-Coordinate of representative point within polygon of observed plant species", "Northern Arizona University", "", "", _
      "", "Meter, in custom local coordinate system", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("area") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "area", _
      "Area of polygon of observed plant species, in square meters", "Northern Arizona University", "", "", _
      "", "Square Meter", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("ID") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "ID", _
      "Original ID Value of this observation, extracted from source shapefiles", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("FClassName") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "FClassName", _
      "Source Shapefile or feature class name", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Quadrat") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Quadrat", _
      "Quadrat Name or Number", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Year") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Year", _
      "Year of Observation", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Type") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Type", _
      "Whether observation is of type Density or Cover", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Orig_FID") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Orig_FID", _
      "Original Feature ID Value of this observation, extracted from source shapefiles", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("IsEmpty") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "IsEmpty", _
      "Whether this is an empty geometry, indicating it was probably a placeholder for a year with no observed species on the quadrat", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Verb_Spcs") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Verb_Spcs", _
      "Verbatim Species Value.  This is the original species recorded on the datasheet.  It might be changed in post-processing " & _
          "if it is misspelled or misidentified in the field.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Verb_Type") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Verb_Type", _
      "Verbatim Type Value.  This is the original type recorded on the datasheet (cover vs. density).  This might be changed in post-processing " & _
          "if it is misidentified in the field.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Revise_Rtn") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Revise_Rtn", _
      "Rotation value.  In a few cases, the quadrat was mapped with the wrong orientation on the datasheet.  This value indicates the " & _
      "rotation applied to the datasheet to correct it.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Site") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Site", _
      "Site name of quadrat.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Site_Specific") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Site_Specific", _
      "Site name of quadrat, with additional details about the sub-region within the larger site.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Plot") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Plot", _
      "Quadrat name.  These values are based on Quadrat numbers.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Tree_Perc_Canopy_Cover") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Tree_Perc_Canopy_Cover", _
      "Percent Canopy Cover [between 1 and 100] of overstory trees within 20m x 20m overstory plots surrounding quadrat.", "Northern Arizona University", "0", "100", _
      "", "Percent", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Tree_Basal_Area_per_Ha") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Tree_Basal_Area_per_Ha", _
      "Basal Area per Hectare of overstory trees within 20m x 20m overstory plots surrounding quadrat.", "Northern Arizona University", "0", "10000", _
      "", "Square Meters", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Soil") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Soil", _
      "Soil parent material; either 'Sed' (for sedimentary) or 'Bas' (for basalt).", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Parent_Material_Class") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Parent_Material_Class", _
      "Soil parent material; either 'Sed' (for sedimentary) or 'Bas' (for basalt).", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Elevation_m") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Elevation_m", _
      "Elevation of plot center, in meters, calculated using bilinear interpolation from 30m DEM downloaded from " & _
          "USGS 3D Elevation Program (3DEP).", "Northern Arizona University", "", "", _
      "", "Meter", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Aspect") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Aspect", _
      "Bearing in direction of steepest slope at plot center, in degrees.", "Northern Arizona University", "", "", _
      "", "Degrees", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Slope_Percent") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Slope_Percent", _
      "Slope at plot center, in percent.", "Northern Arizona University", "", "", _
      "", "Percent", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Easting_NAD_1983_UTM_12") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Easting_NAD_1983_UTM_12", _
      "Plot center Easting (X) coordinate, in UTM Zone 12, projected from NAD 1983", "Northern Arizona University", "", "", _
      "", "Meter; X-Coordinate", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Northing_NAD_1983_UTM_12") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Northing_NAD_1983_UTM_12", _
      "Plot center Northing (Y) coordinate, in UTM Zone 12, projected from NAD 1983", "Northern Arizona University", "", "", _
      "", "Meter; Y-Coordinate", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Longitude_NAD_1983") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Longitude_NAD_1983", _
      "Plot center Longitude (X) coordinate, in the North American Datum of 1983", "Northern Arizona University", "", "", _
      "", "Degrees Longitude", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Latitude_NAD_1983") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Latitude_NAD_1983", _
      "Plot center Latitude (Y) coordinate, in the North American Datum of 1983", "Northern Arizona University", "", "", _
      "", "Degrees Latitude", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Abbreviation") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Abbreviation", _
      "6-Letter abbreviation for plant species, consisting of first 3 letters of Genus name and " & _
          "first three letters of Species name.", "Northern Arizona University", "", "", _
      "", "Meter; Y-Coordinate", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Number_Observations") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Number_Observations", _
      "The number of times this species was recorded in this quadrat this year.", "Northern Arizona University", "", "", _
      "", "Count", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Area_Sq_Cm") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Area_Sq_Cm", _
      "The total area within the quadrat covered by this species in this quadrat this year.", "Northern Arizona University", "", "", _
      "", "Square Centimeters", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Proportion_Quadrat") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Area_Sq_Cm", _
      "The proportion of the quadrat covered by this species in this quadrat this year.", "Northern Arizona University", "", "", _
      "", "Percent", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("AKA") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "AKA", _
      "Alternate quadrat naming system.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Exclosure") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Exclosure", _
      "Comment indicating whether quadrat was inside or outside of a fenced cattle exclosure.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Note") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Note", _
      "Comment describing various aspects of quadrat.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Comment") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Comment", _
      "Comment describing various aspects of quadrat.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If
  If pTable.FindField("Comment_2") > -1 Then
    strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "Comment_2", _
      "Comment describing various aspects of quadrat.", "Northern Arizona University", "", "", _
      "", "", "", "", Null, _
      "", "", "", True)
  End If

  strResponse = Metadata_Functions.SetMetadataFormatVersion(pDataset, "10.8.1", False, pMxDoc)

  strResponse = Metadata_Functions.SynchronizeMetadataPropSet(pDataset)

ClearMemory:
  Set pDataset = Nothing
  Set pPropSet = Nothing
  Set pKeyWords = Nothing
  Set pIncludeThemeKeys = Nothing
  Set pIncludeSearchKeys = Nothing
  Set pIncludeDescKeys = Nothing
  Set pIncludeStratKeys = Nothing
  Set pIncludeThemeSlashThemekeys = Nothing
  Set pIncludePlaceKeys = Nothing
  Set pIncludeTemporalKeys = Nothing
  Set pCombinedKeyWords = Nothing

End Sub


