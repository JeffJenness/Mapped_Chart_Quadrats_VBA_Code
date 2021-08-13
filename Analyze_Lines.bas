Attribute VB_Name = "Analyze_Lines"
Option Explicit

Public Sub AnalyzeLines_FromTables_Phase3()
  
  ' May 4, 2016:
  ' Analyzing only those values in which they were predicted to be caves but were not known to be caves
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim pCSVFiles As esriSystem.IStringArray
  Set pCSVFiles = MyGeneralOperations.ReturnFilesFromNestedFolders2( _
      "D:\arcGIS_stuff\consultation\Jut_Wynne\aaa_Phase_3\CSV_Files", ".csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Jut_Wynne\aaa_Phase_3\" & _
      "Cross_Data_May_4_2016.gdb", 0)
  
  Dim lngIndex As Long
  Dim strFilename As String
  Dim pTransform2D As ITransform2D
  Dim strDayThermal As String
  Dim strDayCurve As String
  Dim strDaySlope As String
  Dim strDayTPI As String
  Dim strNightThermal As String
  Dim strNightCurve As String
  Dim strNightSlope As String
  Dim strNightTPI As String
  Dim strDiffThermal As String
  Dim strDiffCurve As String
  Dim strDiffSlope As String
  Dim strDiffTPI As String
  
  strDayThermal = "Day_Misclass_As_Cave_Thermal_May_4_2016"
  strDaySlope = "Day_Misclass_As_Cave_Slope_May_4_2016"
  strDayCurve = "Day_Misclass_As_Cave_Curve_May_4_2016"
  strDayTPI = "Day_Misclass_As_Cave_TPI_May_4_2016"
  strNightThermal = "Night_Misclass_As_Cave_Thermal_May_4_2016"
  strNightSlope = "Night_Misclass_As_Cave_Slope_May_4_2016"
  strNightCurve = "Night_Misclass_As_Cave_Curve_May_4_2016"
  strNightTPI = "Night_Misclass_As_Cave_TPI_May_4_2016"
  strDiffThermal = "Diff_Misclass_As_Cave_Thermal_May_4_2016"
  strDiffSlope = "Diff_Misclass_As_Cave_Slope_May_4_2016"
  strDiffCurve = "Diff_Misclass_As_Cave_Curve_May_4_2016"
  strDiffTPI = "Diff_Misclass_As_Cave_TPI_May_4_2016"
  
  Dim pTableOfMeans As ITable
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim strTableName As String
  Dim pMeanFields As esriSystem.IVariantArray
  
  Set pMeanFields = New esriSystem.varArray
  strTableName = MyGeneralOperations.MakeUniqueGDBTableName(pWS, "Mean_Values")
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Variable"
    .Type = esriFieldTypeString
    .length = 25
  End With
  pMeanFields.Add pField
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Day_Night"
    .Type = esriFieldTypeString
    .length = 25
  End With
  pMeanFields.Add pField
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Cave_Random"
    .Type = esriFieldTypeString
    .length = 25
  End With
  pMeanFields.Add pField
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Mean_Value"
    .Type = esriFieldTypeDouble
  End With
  pMeanFields.Add pField
  
  Dim pMeanTable As ITable
  Set pMeanTable = MyGeneralOperations.CreateGDBTable(pWS, strTableName, pMeanFields)
  
  Dim pMeanCursor As ICursor
  Dim pMeanBuffer As IRowBuffer
  Dim lngCaveRandomIndex As Long
  Dim lngDayNightIndex As Long
  Dim lngVariableIndex As Long
  Dim lngMeanValueIndex As Long
    
  lngCaveRandomIndex = pMeanTable.FindField("Cave_Random")
  lngDayNightIndex = pMeanTable.FindField("Day_Night")
  lngVariableIndex = pMeanTable.FindField("Variable")
  lngMeanValueIndex = pMeanTable.FindField("Mean_Value")
  Set pMeanCursor = pMeanTable.Insert(True)
  Set pMeanBuffer = pMeanTable.CreateRowBuffer
  
  Dim pDayThermalTable As ITable
  Dim pDayCurveTable As ITable
  Dim pDaySlopeTable As ITable
  Dim pDayTPITable As ITable
  Dim pNightThermalTable As ITable
  Dim pNightCurveTable As ITable
  Dim pNightSlopeTable As ITable
  Dim pNightTPITable As ITable
  Dim pDiffThermalTable As ITable
  Dim pDiffCurveTable As ITable
  Dim pDiffSlopeTable As ITable
  Dim pDiffTPITable As ITable
  
  Dim pMeanPolylineArray As esriSystem.IArray
  Dim pStDevPolygonArray As esriSystem.IArray
  Dim p95ConfPolygonArray As esriSystem.IArray
  Dim pStatMeanArray As esriSystem.IVariantArray
  Dim pStatStDevArray As esriSystem.IVariantArray
  Dim pStatConfIntArray As esriSystem.IVariantArray
  
  Set pMeanPolylineArray = New esriSystem.Array
  Set pStDevPolygonArray = New esriSystem.Array
  Set p95ConfPolygonArray = New esriSystem.Array
  Set pStatMeanArray = New esriSystem.varArray
  Set pStatStDevArray = New esriSystem.varArray
  Set pStatConfIntArray = New esriSystem.varArray
  
  Dim pPolylineArray As esriSystem.IArray
  Set pPolylineArray = New esriSystem.Array
  Dim pGeomArray As esriSystem.IArray
  Set pGeomArray = New esriSystem.Array
  Dim pPolylineValArray As esriSystem.IVariantArray
  Set pPolylineValArray = New esriSystem.varArray
  Dim pPolylineSubArray As esriSystem.IVariantArray
  Dim pPolylineFieldArray As esriSystem.IVariantArray
  Set pPolylineFieldArray = New esriSystem.varArray
  Dim pPolylineField As iField
  Dim pPolylineFieldEdit As IFieldEdit
  Dim pPolyline As IPolyline
  
  Dim pGraphMap As IMap
  Set pGraphMap = MyGeneralOperations.ReturnMapByName("Graphs", pMxDoc)
  
  Set pPolylineField = New Field
  Set pPolylineFieldEdit = pPolylineField
  With pPolylineFieldEdit
    .Name = "Name"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pPolylineFieldArray.Add pPolylineField
  
  Set pPolylineField = New Field
  Set pPolylineFieldEdit = pPolylineField
  With pPolylineFieldEdit
    .Name = "Cave_Or_Random"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pPolylineFieldArray.Add pPolylineField
      
  Set pDayThermalTable = pWS.OpenTable(strDayThermal)
  Set pDaySlopeTable = pWS.OpenTable(strDaySlope)
  Set pDayCurveTable = pWS.OpenTable(strDayCurve)
  Set pDayTPITable = pWS.OpenTable(strDayTPI)
  Set pNightThermalTable = pWS.OpenTable(strNightThermal)
  Set pNightSlopeTable = pWS.OpenTable(strNightSlope)
  Set pNightCurveTable = pWS.OpenTable(strNightCurve)
  Set pNightTPITable = pWS.OpenTable(strNightTPI)
  Set pDiffThermalTable = pWS.OpenTable(strDiffThermal)
  Set pDiffSlopeTable = pWS.OpenTable(strDiffSlope)
  Set pDiffCurveTable = pWS.OpenTable(strDiffCurve)
  Set pDiffTPITable = pWS.OpenTable(strDiffTPI)
  
  Dim dblCaveMean As Double
  Dim dblRandomMean As Double
  
  ' DAY THERMAL
  Call RunDatasetAnalysis_FromTable_Phase_3(pDayThermalTable, "Daytime Thermal", 0, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Missclassified As Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer
  
  ' DAY Slope
  Call RunDatasetAnalysis_FromTable_Phase_3(pDaySlopeTable, "Daytime Slope", 10, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Missclassified As Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

'   Day curve
  Call RunDatasetAnalysis_FromTable_Phase_3(pDayCurveTable, "Daytime Curvature", 20, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Missclassified As Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' DAY TPI
  Call RunDatasetAnalysis_FromTable_Phase_3(pDayTPITable, "Daytime TPI", 30, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Missclassified As Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer


  ' NIGHT THERMAL
  Call RunDatasetAnalysis_FromTable_Phase_3(pNightThermalTable, "Night Thermal", 40, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' NIGHT Slope
  Call RunDatasetAnalysis_FromTable_Phase_3(pNightSlopeTable, "Night Slope", 50, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' NIGHT curve
  Call RunDatasetAnalysis_FromTable_Phase_3(pNightCurveTable, "Night Curvature", 60, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' NIGHT TPI
  Call RunDatasetAnalysis_FromTable_Phase_3(pNightTPITable, "Night TPI", 70, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer


  ' DIFFERENCE THERMAL
  Call RunDatasetAnalysis_FromTable_Phase_3(pDiffThermalTable, "Diff Thermal", 40, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' DIFFERENCE Slope
  Call RunDatasetAnalysis_FromTable_Phase_3(pDiffSlopeTable, "Diff Slope", 50, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' DIFFERENCE curve
  Call RunDatasetAnalysis_FromTable_Phase_3(pDiffCurveTable, "Diff Curvature", 60, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' DIFFERENCE TPI
  Call RunDatasetAnalysis_FromTable_Phase_3(pDiffTPITable, "Diff TPI", 70, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Diff"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer
    
  
'  CreatePolylineFClass pWS, "Mean_Polylines", pMeanPolylineArray, pStatMeanArray, pPolylineFieldArray, True
'  CreatePolylineFClass pWS, "Polyline_Standard_Deviation", pStDevPolygonArray, pStatStDevArray, pPolylineFieldArray, False
'  CreatePolylineFClass pWS, "Polyline_95_Percent_CI", p95ConfPolygonArray, pStatConfIntArray, pPolylineFieldArray, False
  
  pMeanCursor.Flush
  
  
  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pCSVFiles = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTransform2D = Nothing
  Set pTableOfMeans = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pMeanFields = Nothing
  Set pMeanTable = Nothing
  Set pMeanCursor = Nothing
  Set pMeanBuffer = Nothing
  Set pDayThermalTable = Nothing
  Set pDayCurveTable = Nothing
  Set pDaySlopeTable = Nothing
  Set pDayTPITable = Nothing
  Set pNightThermalTable = Nothing
  Set pNightCurveTable = Nothing
  Set pNightSlopeTable = Nothing
  Set pNightTPITable = Nothing
  Set pDiffThermalTable = Nothing
  Set pDiffCurveTable = Nothing
  Set pDiffSlopeTable = Nothing
  Set pDiffTPITable = Nothing
  Set pMeanPolylineArray = Nothing
  Set pStDevPolygonArray = Nothing
  Set p95ConfPolygonArray = Nothing
  Set pStatMeanArray = Nothing
  Set pStatStDevArray = Nothing
  Set pStatConfIntArray = Nothing
  Set pPolylineArray = Nothing
  Set pGeomArray = Nothing
  Set pPolylineValArray = Nothing
  Set pPolylineSubArray = Nothing
  Set pPolylineFieldArray = Nothing
  Set pPolylineField = Nothing
  Set pPolylineFieldEdit = Nothing
  Set pPolyline = Nothing
  Set pGraphMap = Nothing





End Sub

Public Sub RunDatasetAnalysis_FromTable_Phase_3(pTable As ITable, strName As String, dblXOffset As Double, _
    pMxDoc As IMxDocument, pPolylineArray As esriSystem.IArray, _
    pPolylineValArray As esriSystem.IVariantArray, pGeomArray As IArray, pWS As IWorkspace, _
    pPolylineFieldArray As esriSystem.IVariantArray, pMeanPolylineArray As esriSystem.IArray, _
    pStDevPolygonArray As esriSystem.IArray, p95ConfPolygonArray As esriSystem.IArray, _
    pStatMeanArray As esriSystem.IVariantArray, pStatStDevArray As esriSystem.IVariantArray, _
    pStatConfIntArray As esriSystem.IVariantArray, dblCaveMean As Double, _
    dblRandomMean As Double)
  
  ' MAY 4, 2016
  ' MODIFIED TO NOT EXAMINE RANDOM VALUES
  ' ALSO DON'T LOOK FOR "AT_CAVE" BUT RATHER "Actual_Raster_Predicted_Cave"
  
  ' DAY Slope
  Dim strText As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String
  Dim lngCount As Long
  Dim lngInterval As Long
  Dim lngCounter As Long
  Dim booFirst As Boolean
  Dim lngCaveCounter As Long
  Dim lngRandomCounter As Long
  Dim booIsFirst As Boolean
  Dim lngIndex As Long
  Dim strCaveRandom As String
'  Dim dblCaveVals() As Double
  Dim varCaveVals() As Variant
  Dim dblMinVal As Double
  Dim dblMaxVal As Double
'  Dim dblCaveMean As Double
'  Dim dblRandomMean As Double
  Dim dblRescaleCave() As Double
  Dim dblRescaleRandom() As Double
  Dim dblGraphMax As Double
  Dim dblGraphArray() As Double
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim varVal As Variant
  Dim dblRandomVals() As Double
  
  Dim pGraphMap As IMap
  Set pGraphMap = MyGeneralOperations.ReturnMapByName("Graphs", pMxDoc)
  
  Dim lngCaveOrRandomIndex As Long
  lngCaveOrRandomIndex = pTable.FindField("CaveRandom")
  Dim lngValueIndices() As Long
  ReDim lngValueIndices(32)
  lngCounter = -1
  For lngIndex = -400 To 400 Step 25
    lngCounter = lngCounter + 1
    If lngIndex < 0 Then
      lngValueIndices(lngCounter) = pTable.FindField("n" & Format(Abs(lngIndex), "000"))
    Else
      lngValueIndices(lngCounter) = pTable.FindField("p" & Format(lngIndex, "000"))
    End If
  Next lngIndex
  
  lngCount = pTable.RowCount(Nothing)
  lngInterval = lngCount / 10
  lngCounter = 0
  booFirst = True
  lngCaveCounter = -1
  lngRandomCounter = -1
  booIsFirst = True
  Dim lngTotalCounter As Long
  lngTotalCounter = -1
  Dim booFoundValue As Boolean
  
  Debug.Print strName & ":"
  
  Dim pCursor As ICursor
  Dim pRow As IRow
  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    lngCounter = lngCounter + 1
    lngTotalCounter = lngTotalCounter + 1
    If lngCounter >= lngInterval Then
      lngCounter = 0
      Debug.Print "  --> " & Format(lngTotalCounter, "#,##0"); " of " & Format(lngCount, "#,##0") & _
          " [" & Format(lngTotalCounter * 100 / lngCount, "0") & "%]"
      DoEvents
    End If
    strCaveRandom = pRow.Value(lngCaveOrRandomIndex)
    If strCaveRandom = "Actual_Raster_Predicted_Cave" Then
      lngCaveCounter = lngCaveCounter + 1
      ReDim Preserve varCaveVals(32, lngCaveCounter)
      booFoundValue = False
      For lngIndex2 = 0 To 32
'        dblVal = pRow.Value(lngValueIndices(lngIndex2))
        varVal = pRow.Value(lngValueIndices(lngIndex2))
        If IsNumeric(varVal) Then
          dblVal = CDbl(varVal)
          If booFirst Then
            dblMaxVal = dblVal
            dblMinVal = dblVal
            booFirst = False
          Else
            If dblVal > dblMaxVal Then dblMaxVal = dblVal
            If dblVal < dblMinVal Then dblMinVal = dblVal
          End If
          varCaveVals(lngIndex2, lngCaveCounter) = dblVal
        Else
          varCaveVals(lngIndex2, lngCaveCounter) = Null
        End If
      Next lngIndex2
      
'    Else
'      lngRandomCounter = lngRandomCounter + 1
'      ReDim Preserve dblRandomVals(32, lngRandomCounter)
'
'      For lngIndex2 = 0 To 32
'        dblVal = pRow.Value(lngValueIndices(lngIndex2))
'        If booFirst Then
'          dblMaxVal = dblVal
'          dblMinVal = dblVal
'          booFirst = False
'        Else
'          If dblVal > dblMaxVal Then dblMaxVal = dblVal
'          If dblVal < dblMinVal Then dblMinVal = dblVal
'        End If
'        dblRandomVals(lngIndex2, lngRandomCounter) = dblVal
'      Next lngIndex2
    End If
    Set pRow = pCursor.NextRow
  Loop
      
  Dim strRow As String
  Dim strPolylineName As String
  Dim pMeanPolyline As IPolyline
  Dim pStDevPolygon As IPolygon
  Dim pConfIntPolygon As IPolygon
  Dim pStatValSubArray As esriSystem.IVariantArray
  Dim varRescaleCave() As Variant
  

'  For lngIndex = 0 To UBound(varCaveVals, 2)
'    strRow = ""
'    For lngIndex2 = 0 To UBound(varCaveVals, 1)
'      strRow = strRow & IIf(IsNull(varCaveVals(lngIndex2, lngIndex)), " Null, ", _
'          Format(varCaveVals(lngIndex2, lngIndex), "0.0") & ", ")
'    Next lngIndex2
''    Debug.Print strRow
'  Next lngIndex
  
  dblCaveMean = ReturnMeanValue_ExcludeNulls(varCaveVals)
  
''  dblRandomMean = ReturnMeanValue(dblRandomVals)
  varRescaleCave = ShiftToMean_ExcludeNulls(varCaveVals, dblCaveMean, dblMinVal, dblMaxVal, True)
  
  
'  For lngIndex = 0 To UBound(varCaveVals, 2)
'    strRow = ""
'    For lngIndex2 = 0 To UBound(varCaveVals, 1)
'      strRow = strRow & IIf(IsNull(varRescaleCave(lngIndex2, lngIndex)), " Null, ", _
'          Format(varRescaleCave(lngIndex2, lngIndex), "0.0") & ", ")
'    Next lngIndex2
''    If Left(strRow, 5) <> " Null" Then
''      Debug.Print strRow
''    End If
'
'  Next lngIndex
  
'  dblRescaleRandom = ShiftToMean(dblRandomVals, dblRandomMean, dblMinVal, dblMaxVal, False)
  varRescaleCave = RescaleTo4High_ExcludeNulls(varRescaleCave, dblMinVal, dblMaxVal)
'  dblRescaleRandom = RescaleTo4High(dblRescaleRandom, dblMinVal, dblMaxVal)

'  For lngIndex = 0 To UBound(varCaveVals, 2)
'    strRow = ""
'    For lngIndex2 = 0 To UBound(varCaveVals, 1)
'      strRow = strRow & IIf(IsNull(varRescaleCave(lngIndex2, lngIndex)), " Null, ", _
'          Format(varRescaleCave(lngIndex2, lngIndex), "0.0") & ", ")
'    Next lngIndex2
'    If Left(strRow, 5) <> " Null" Then
'      Debug.Print strRow
'    End If
'
'  Next lngIndex

  dblGraphArray = FillGraphArray2_ExcludeNulls(varRescaleCave, dblGraphMax, pMxDoc, pPolylineArray, pMeanPolyline, _
      pStDevPolygon, pConfIntPolygon, dblXOffset, 6)

  pMeanPolylineArray.Add pMeanPolyline
  Set pStatValSubArray = New esriSystem.varArray
  pStatValSubArray.Add strName & " Mean"
  pStatValSubArray.Add "At Cave"
  pStatMeanArray.Add pStatValSubArray

  pStDevPolygonArray.Add pStDevPolygon
  Set pStatValSubArray = New esriSystem.varArray
  pStatValSubArray.Add strName & " Standard Deviation"
  pStatValSubArray.Add "At Cave"
  pStatStDevArray.Add pStatValSubArray

  p95ConfPolygonArray.Add pConfIntPolygon
  Set pStatValSubArray = New esriSystem.varArray
  pStatValSubArray.Add strName & " 95% Confidence Interval"
  pStatValSubArray.Add "At Cave"
  pStatConfIntArray.Add pStatValSubArray
    
  AddToPolylineArray pPolylineArray, strName, "At Cave", pPolylineValArray, pGeomArray, dblXOffset, 6
  strPolylineName = Replace(strName, " ", "_") & "Cave_Profiles"
  CreatePolylineFClass pWS, strPolylineName, pGeomArray, pPolylineValArray, pPolylineFieldArray, True
  ConvertToPointArray dblGraphArray, dblXOffset, 6, dblGraphMax, strName, "At Cave", pGraphMap, _
      dblMinVal, dblMaxVal, dblCaveMean, pWS
      
  
'  dblGraphArray = FillGraphArray2(dblRescaleRandom, dblGraphMax, pMxDoc, pPolylineArray, pMeanPolyline, _
'      pStDevPolygon, pConfIntPolygon, dblXOffset, 0)
'
'  pMeanPolylineArray.Add pMeanPolyline
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " Mean"
'  pStatValSubArray.Add "Random"
'  pStatMeanArray.Add pStatValSubArray
'
'  pStDevPolygonArray.Add pStDevPolygon
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " Standard Deviation"
'  pStatValSubArray.Add "Random"
'  pStatStDevArray.Add pStatValSubArray
'
'  p95ConfPolygonArray.Add pConfIntPolygon
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " 95% Confidence Interval"
'  pStatValSubArray.Add "Random"
'  pStatConfIntArray.Add pStatValSubArray
  
'  AddToPolylineArray pPolylineArray, strName, "Random", pPolylineValArray, pGeomArray, dblXOffset, 0
'  strPolylineName = Replace(strName, " ", "_") & "Random_Profiles"
'  CreatePolylineFClass pWS, strPolylineName, pGeomArray, pPolylineValArray, pPolylineFieldArray, True
'  ConvertToPointArray dblGraphArray, dblXOffset, 0, dblGraphMax, strName, "Random", pGraphMap, _
      dblMinVal, dblMaxVal, dblRandomMean, pWS
  

ClearMemory:
  Erase strLines
  Erase strLineSplit
  Erase varCaveVals
  Erase dblRescaleCave
  Erase dblRescaleRandom
  Erase dblGraphArray
  Erase dblRandomVals
  Set pGraphMap = Nothing
  Erase lngValueIndices
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pMeanPolyline = Nothing
  Set pStDevPolygon = Nothing
  Set pConfIntPolygon = Nothing
  Set pStatValSubArray = Nothing


  
End Sub





Public Sub AnalyzeLines_FromTables()
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim pCSVFiles As esriSystem.IStringArray
  Set pCSVFiles = MyGeneralOperations.ReturnFilesFromNestedFolders2( _
      "D:\arcGIS_stuff\consultation\Jut_Wynne\Analysis_Files\October_22_2015_Outputs", ".csv")
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Jut_Wynne\Analysis_Files\" & _
      "October_22_2015_Outputs\Cross_Data_Oct_31_2015.gdb", 0)
  
  Dim lngIndex As Long
  Dim strDayThermal As String
  Dim strDayCurve As String
  Dim strDaySlope As String
  Dim strDayTPI As String
  Dim strNightThermal As String
  Dim strNightCurve As String
  Dim strNightSlope As String
  Dim strNightTPI As String
  Dim strFilename As String
  Dim pTransform2D As ITransform2D
  
  strDayThermal = "Mojave_Day_Cave_Data_Thermal_Oct_24_2015"
  strDaySlope = "Mojave_Day_Cave_Data_Slope_Oct_24_2015"
  strDayCurve = "Mojave_Day_Cave_Data_Curve_Oct_24_2015"
  strDayTPI = "Mojave_Day_Cave_Data_TPI_Oct_24_2015"
  strNightThermal = "Mojave_night_Cave_Data_Thermal_Oct_24_2015"
  strNightSlope = "Mojave_night_Cave_Data_Slope_Oct_24_2015"
  strNightCurve = "Mojave_night_Cave_Data_Curve_Oct_24_2015"
  strNightTPI = "Mojave_night_Cave_Data_TPI_Oct_24_2015"
  
  Dim pTableOfMeans As ITable
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim strTableName As String
  Dim pMeanFields As esriSystem.IVariantArray
  
  Set pMeanFields = New esriSystem.varArray
  strTableName = MyGeneralOperations.MakeUniqueGDBTableName(pWS, "Mean_Values")
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Variable"
    .Type = esriFieldTypeString
    .length = 25
  End With
  pMeanFields.Add pField
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Day_Night"
    .Type = esriFieldTypeString
    .length = 25
  End With
  pMeanFields.Add pField
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Cave_Random"
    .Type = esriFieldTypeString
    .length = 25
  End With
  pMeanFields.Add pField
    
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Mean_Value"
    .Type = esriFieldTypeDouble
  End With
  pMeanFields.Add pField
  
  Dim pMeanTable As ITable
  Set pMeanTable = MyGeneralOperations.CreateGDBTable(pWS, strTableName, pMeanFields)
  
  Dim pMeanCursor As ICursor
  Dim pMeanBuffer As IRowBuffer
  Dim lngCaveRandomIndex As Long
  Dim lngDayNightIndex As Long
  Dim lngVariableIndex As Long
  Dim lngMeanValueIndex As Long
    
  lngCaveRandomIndex = pMeanTable.FindField("Cave_Random")
  lngDayNightIndex = pMeanTable.FindField("Day_Night")
  lngVariableIndex = pMeanTable.FindField("Variable")
  lngMeanValueIndex = pMeanTable.FindField("Mean_Value")
  Set pMeanCursor = pMeanTable.Insert(True)
  Set pMeanBuffer = pMeanTable.CreateRowBuffer
  
  Dim pDayThermalTable As ITable
  Dim pDayCurveTable As ITable
  Dim pDaySlopeTable As ITable
  Dim pDayTPITable As ITable
  Dim pNightThermalTable As ITable
  Dim pNightCurveTable As ITable
  Dim pNightSlopeTable As ITable
  Dim pNightTPITable As ITable
  
  Dim pMeanPolylineArray As esriSystem.IArray
  Dim pStDevPolygonArray As esriSystem.IArray
  Dim p95ConfPolygonArray As esriSystem.IArray
  Dim pStatMeanArray As esriSystem.IVariantArray
  Dim pStatStDevArray As esriSystem.IVariantArray
  Dim pStatConfIntArray As esriSystem.IVariantArray
  
  Set pMeanPolylineArray = New esriSystem.Array
  Set pStDevPolygonArray = New esriSystem.Array
  Set p95ConfPolygonArray = New esriSystem.Array
  Set pStatMeanArray = New esriSystem.varArray
  Set pStatStDevArray = New esriSystem.varArray
  Set pStatConfIntArray = New esriSystem.varArray
  
  Dim pPolylineArray As esriSystem.IArray
  Set pPolylineArray = New esriSystem.Array
  Dim pGeomArray As esriSystem.IArray
  Set pGeomArray = New esriSystem.Array
  Dim pPolylineValArray As esriSystem.IVariantArray
  Set pPolylineValArray = New esriSystem.varArray
  Dim pPolylineSubArray As esriSystem.IVariantArray
  Dim pPolylineFieldArray As esriSystem.IVariantArray
  Set pPolylineFieldArray = New esriSystem.varArray
  Dim pPolylineField As iField
  Dim pPolylineFieldEdit As IFieldEdit
  Dim pPolyline As IPolyline
  
  Dim pGraphMap As IMap
  Set pGraphMap = MyGeneralOperations.ReturnMapByName("Graphs", pMxDoc)
  
  Set pPolylineField = New Field
  Set pPolylineFieldEdit = pPolylineField
  With pPolylineFieldEdit
    .Name = "Name"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pPolylineFieldArray.Add pPolylineField
  
  Set pPolylineField = New Field
  Set pPolylineFieldEdit = pPolylineField
  With pPolylineFieldEdit
    .Name = "Cave_Or_Random"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pPolylineFieldArray.Add pPolylineField
      
  Set pDayThermalTable = pWS.OpenTable(strDayThermal)
  Set pDaySlopeTable = pWS.OpenTable(strDaySlope)
  Set pDayCurveTable = pWS.OpenTable(strDayCurve)
  Set pDayTPITable = pWS.OpenTable(strDayTPI)
  Set pNightThermalTable = pWS.OpenTable(strNightThermal)
  Set pNightSlopeTable = pWS.OpenTable(strNightSlope)
  Set pNightCurveTable = pWS.OpenTable(strNightCurve)
  Set pNightTPITable = pWS.OpenTable(strNightTPI)
  
  Dim dblCaveMean As Double
  Dim dblRandomMean As Double
  
  ' DAY THERMAL
  Call RunDatasetAnalysis_FromTable(pDayThermalTable, "Daytime Thermal", 0, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer
  
  ' DAY Slope
  Call RunDatasetAnalysis_FromTable(pDaySlopeTable, "Daytime Slope", 10, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

'   Day curve
  Call RunDatasetAnalysis_FromTable(pDayCurveTable, "Daytime Curvature", 20, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' DAY TPI
  Call RunDatasetAnalysis_FromTable(pDayTPITable, "Daytime TPI", 30, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Day"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer


  ' NIGHT THERMAL
  Call RunDatasetAnalysis_FromTable(pNightThermalTable, "Night Thermal", 40, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Thermal"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' NIGHT Slope
  Call RunDatasetAnalysis_FromTable(pNightSlopeTable, "Night Slope", 50, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Slope"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' NIGHT curve
  Call RunDatasetAnalysis_FromTable(pNightCurveTable, "Night Curvature", 60, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "Curvature"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer

  ' NIGHT TPI
  Call RunDatasetAnalysis_FromTable(pNightTPITable, "Night TPI", 70, pMxDoc, pPolylineArray, _
       pPolylineValArray, pGeomArray, pWS, pPolylineFieldArray, pMeanPolylineArray, pStDevPolygonArray, _
       p95ConfPolygonArray, pStatMeanArray, pStatStDevArray, pStatConfIntArray, dblCaveMean, dblRandomMean)
  pMeanBuffer.Value(lngCaveRandomIndex) = "Cave"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblCaveMean
  pMeanCursor.InsertRow pMeanBuffer
  pMeanBuffer.Value(lngCaveRandomIndex) = "Random"
  pMeanBuffer.Value(lngDayNightIndex) = "Night"
  pMeanBuffer.Value(lngVariableIndex) = "TPI"
  pMeanBuffer.Value(lngMeanValueIndex) = dblRandomMean
  pMeanCursor.InsertRow pMeanBuffer
    
  
'  CreatePolylineFClass pWS, "Mean_Polylines", pMeanPolylineArray, pStatMeanArray, pPolylineFieldArray, True
'  CreatePolylineFClass pWS, "Polyline_Standard_Deviation", pStDevPolygonArray, pStatStDevArray, pPolylineFieldArray, False
'  CreatePolylineFClass pWS, "Polyline_95_Percent_CI", p95ConfPolygonArray, pStatConfIntArray, pPolylineFieldArray, False
  
  pMeanCursor.Flush
  
  
  Debug.Print "Done..."
  

ClearMemory:
  Set pMxDoc = Nothing
  Set pCSVFiles = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTransform2D = Nothing
  Set pDayThermalTable = Nothing
  Set pDayCurveTable = Nothing
  Set pDaySlopeTable = Nothing
  Set pDayTPITable = Nothing
  Set pNightThermalTable = Nothing
  Set pNightCurveTable = Nothing
  Set pNightSlopeTable = Nothing
  Set pNightTPITable = Nothing
  Set pMeanPolylineArray = Nothing
  Set pStDevPolygonArray = Nothing
  Set p95ConfPolygonArray = Nothing
  Set pStatMeanArray = Nothing
  Set pStatStDevArray = Nothing
  Set pStatConfIntArray = Nothing
  Set pPolylineArray = Nothing
  Set pGeomArray = Nothing
  Set pPolylineValArray = Nothing
  Set pPolylineSubArray = Nothing
  Set pPolylineFieldArray = Nothing
  Set pPolylineField = Nothing
  Set pPolylineFieldEdit = Nothing
  Set pPolyline = Nothing
  Set pGraphMap = Nothing


End Sub


Public Sub RunDatasetAnalysis_FromTable(pTable As ITable, strName As String, dblXOffset As Double, _
    pMxDoc As IMxDocument, pPolylineArray As esriSystem.IArray, _
    pPolylineValArray As esriSystem.IVariantArray, pGeomArray As IArray, pWS As IWorkspace, _
    pPolylineFieldArray As esriSystem.IVariantArray, pMeanPolylineArray As esriSystem.IArray, _
    pStDevPolygonArray As esriSystem.IArray, p95ConfPolygonArray As esriSystem.IArray, _
    pStatMeanArray As esriSystem.IVariantArray, pStatStDevArray As esriSystem.IVariantArray, _
    pStatConfIntArray As esriSystem.IVariantArray, dblCaveMean As Double, _
    dblRandomMean As Double)
  
  ' DAY Slope
  Dim strText As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String
  Dim lngCount As Long
  Dim lngInterval As Long
  Dim lngCounter As Long
  Dim booFirst As Boolean
  Dim lngCaveCounter As Long
  Dim lngRandomCounter As Long
  Dim booIsFirst As Boolean
  Dim lngIndex As Long
  Dim strCaveRandom As String
  Dim dblCaveVals() As Double
  Dim dblMinVal As Double
  Dim dblMaxVal As Double
'  Dim dblCaveMean As Double
'  Dim dblRandomMean As Double
  Dim dblRescaleCave() As Double
  Dim dblRescaleRandom() As Double
  Dim dblGraphMax As Double
  Dim dblGraphArray() As Double
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim dblRandomVals() As Double
  
  Dim pGraphMap As IMap
  Set pGraphMap = MyGeneralOperations.ReturnMapByName("Graphs", pMxDoc)
  
  Dim lngCaveOrRandomIndex As Long
  lngCaveOrRandomIndex = pTable.FindField("NearestCaveID")
  Dim lngValueIndices() As Long
  ReDim lngValueIndices(32)
  lngCounter = -1
  For lngIndex = -400 To 400 Step 25
    lngCounter = lngCounter + 1
    If lngIndex < 0 Then
      lngValueIndices(lngCounter) = pTable.FindField("n" & Format(Abs(lngIndex), "000"))
    Else
      lngValueIndices(lngCounter) = pTable.FindField("p" & Format(lngIndex, "000"))
    End If
  Next lngIndex
  
  lngCount = pTable.RowCount(Nothing)
  lngInterval = lngCount / 10
  lngCounter = 0
  booFirst = True
  lngCaveCounter = -1
  lngRandomCounter = -1
  booIsFirst = True
  Dim lngTotalCounter As Long
  lngTotalCounter = -1
  
  Debug.Print strName & ":"
  
  Dim pCursor As ICursor
  Dim pRow As IRow
  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    lngCounter = lngCounter + 1
    lngTotalCounter = lngTotalCounter + 1
    If lngCounter >= lngInterval Then
      lngCounter = 0
      Debug.Print "  --> " & Format(lngTotalCounter, "#,##0"); " of " & Format(lngCount, "#,##0") & _
          " [" & Format(lngTotalCounter * 100 / lngCount, "0") & "%]"
      DoEvents
    End If
    strCaveRandom = pRow.Value(lngCaveOrRandomIndex)
    If strCaveRandom = "At_Cave" Then
      lngCaveCounter = lngCaveCounter + 1
      ReDim Preserve dblCaveVals(32, lngCaveCounter)
      For lngIndex2 = 0 To 32
        dblVal = pRow.Value(lngValueIndices(lngIndex2))
        If booFirst Then
          dblMaxVal = dblVal
          dblMinVal = dblVal
          booFirst = False
        Else
          If dblVal > dblMaxVal Then dblMaxVal = dblVal
          If dblVal < dblMinVal Then dblMinVal = dblVal
        End If
        dblCaveVals(lngIndex2, lngCaveCounter) = dblVal
      Next lngIndex2
      
    Else
      lngRandomCounter = lngRandomCounter + 1
      ReDim Preserve dblRandomVals(32, lngRandomCounter)
      
      For lngIndex2 = 0 To 32
        dblVal = pRow.Value(lngValueIndices(lngIndex2))
        If booFirst Then
          dblMaxVal = dblVal
          dblMinVal = dblVal
          booFirst = False
        Else
          If dblVal > dblMaxVal Then dblMaxVal = dblVal
          If dblVal < dblMinVal Then dblMinVal = dblVal
        End If
        dblRandomVals(lngIndex2, lngRandomCounter) = dblVal
      Next lngIndex2
    End If
    Set pRow = pCursor.NextRow
  Loop
      
  Dim strPolylineName As String
  Dim pMeanPolyline As IPolyline
  Dim pStDevPolygon As IPolygon
  Dim pConfIntPolygon As IPolygon
  Dim pStatValSubArray As esriSystem.IVariantArray
  
  dblCaveMean = ReturnMeanValue(dblCaveVals)
  dblRandomMean = ReturnMeanValue(dblRandomVals)
'  dblRescaleCave = ShiftToMean(dblCaveVals, dblCaveMean, dblMinVal, dblMaxVal, True)
'  dblRescaleRandom = ShiftToMean(dblRandomVals, dblRandomMean, dblMinVal, dblMaxVal, False)
'  dblRescaleCave = RescaleTo4High(dblRescaleCave, dblMinVal, dblMaxVal)
'  dblRescaleRandom = RescaleTo4High(dblRescaleRandom, dblMinVal, dblMaxVal)
'
'  dblGraphArray = FillGraphArray2(dblRescaleCave, dblGraphMax, pMxDoc, pPolylineArray, pMeanPolyline, _
'      pStDevPolygon, pConfIntPolygon, dblXOffset, 6)
'
'  pMeanPolylineArray.Add pMeanPolyline
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " Mean"
'  pStatValSubArray.Add "At Cave"
'  pStatMeanArray.Add pStatValSubArray
'
'  pStDevPolygonArray.Add pStDevPolygon
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " Standard Deviation"
'  pStatValSubArray.Add "At Cave"
'  pStatStDevArray.Add pStatValSubArray
'
'  p95ConfPolygonArray.Add pConfIntPolygon
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " 95% Confidence Interval"
'  pStatValSubArray.Add "At Cave"
'  pStatConfIntArray.Add pStatValSubArray
    
'  AddToPolylineArray pPolylineArray, strName, "At Cave", pPolylineValArray, pGeomArray, dblXOffset, 6
'  strPolylineName = Replace(strName, " ", "_") & "Cave_Profiles"
'  CreatePolylineFClass pWS, strPolylineName, pGeomArray, pPolylineValArray, pPolylineFieldArray, True
'  ConvertToPointArray dblGraphArray, dblXOffset, 6, dblGraphMax, strName, "At Cave", pGraphMap, _
      dblMinVal, dblMaxVal, dblCaveMean, pWS
      
  
'  dblGraphArray = FillGraphArray2(dblRescaleRandom, dblGraphMax, pMxDoc, pPolylineArray, pMeanPolyline, _
'      pStDevPolygon, pConfIntPolygon, dblXOffset, 0)
'
'  pMeanPolylineArray.Add pMeanPolyline
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " Mean"
'  pStatValSubArray.Add "Random"
'  pStatMeanArray.Add pStatValSubArray
'
'  pStDevPolygonArray.Add pStDevPolygon
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " Standard Deviation"
'  pStatValSubArray.Add "Random"
'  pStatStDevArray.Add pStatValSubArray
'
'  p95ConfPolygonArray.Add pConfIntPolygon
'  Set pStatValSubArray = New esriSystem.VarArray
'  pStatValSubArray.Add strName & " 95% Confidence Interval"
'  pStatValSubArray.Add "Random"
'  pStatConfIntArray.Add pStatValSubArray
  
'  AddToPolylineArray pPolylineArray, strName, "Random", pPolylineValArray, pGeomArray, dblXOffset, 0
'  strPolylineName = Replace(strName, " ", "_") & "Random_Profiles"
'  CreatePolylineFClass pWS, strPolylineName, pGeomArray, pPolylineValArray, pPolylineFieldArray, True
'  ConvertToPointArray dblGraphArray, dblXOffset, 0, dblGraphMax, strName, "Random", pGraphMap, _
      dblMinVal, dblMaxVal, dblRandomMean, pWS
  

ClearMemory:
  Erase strLines
  Erase strLineSplit
  Erase dblCaveVals
  Erase dblRescaleCave
  Erase dblRescaleRandom
  Erase dblGraphArray
  Erase dblRandomVals
  Set pGraphMap = Nothing
  Erase lngValueIndices
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pMeanPolyline = Nothing
  Set pStDevPolygon = Nothing
  Set pConfIntPolygon = Nothing
  Set pStatValSubArray = Nothing


  
End Sub


Public Sub CreatePolylineFClass(pWS As IWorkspace, strName As String, pPolylineArray As esriSystem.IArray, _
    pPolylineValArray As esriSystem.IVariantArray, pPolylineFieldArray As esriSystem.IVariantArray, _
    booIsPolyline As Boolean)
  
  Dim lngIndex As Long
  Dim pField As iField
  Dim pClone As IClone
  Dim pCloneFieldArray As esriSystem.IVariantArray
  Set pCloneFieldArray = New esriSystem.varArray
  For lngIndex = 0 To pPolylineFieldArray.Count - 1
    Set pField = pPolylineFieldArray.Element(lngIndex)
    Set pClone = pField
    pCloneFieldArray.Add pClone.Clone
  Next lngIndex
  
  Dim pPolygon As IPolygon
  Dim pPolyline As IPolyline
  Dim pSpRef As ISpatialReference
  
  If booIsPolyline Then
    Set pPolyline = pPolylineArray.Element(0)
    Set pSpRef = pPolyline.SpatialReference
  Else
    Set pPolygon = pPolylineArray.Element(0)
    Set pSpRef = pPolygon.SpatialReference
  End If
  
  strName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS, strName)
  Debug.Print "Creating '" & strName
  
  Dim pFClass As IFeatureClass
  If booIsPolyline Then
    Set pFClass = MyGeneralOperations.CreateGDBFeatureClass(pWS, strName, esriFTSimple, pSpRef, esriGeometryPolyline, _
        pCloneFieldArray, , , , False, ENUM_FileGDB)
  Else
    Set pFClass = MyGeneralOperations.CreateGDBFeatureClass(pWS, strName, esriFTSimple, pSpRef, esriGeometryPolygon, _
      pCloneFieldArray, , , , False, ENUM_FileGDB)
  End If
  
  Dim lngFieldIndexArray() As Long
  ReDim lngFieldIndexArray(pCloneFieldArray.Count - 1)
  For lngIndex = 0 To pCloneFieldArray.Count - 1
    Set pField = pCloneFieldArray.Element(lngIndex)
    lngFieldIndexArray(lngIndex) = pFClass.FindField(pField.Name)
  Next lngIndex
  
  Dim pFCursor As IFeatureCursor
  Dim pBuffer As IFeatureBuffer
  Dim pValSubArray As esriSystem.IVariantArray
  Dim lngFieldIndex As Long
  Set pFCursor = pFClass.Insert(True)
  Set pBuffer = pFClass.CreateFeatureBuffer
  
  Dim pGeom As IGeometry
  
  Dim lngIndex2 As Long
  For lngIndex = 0 To pPolylineArray.Count - 1
    If lngIndex Mod 1000 = 0 Then
      DoEvents
      pFCursor.Flush
    End If
    Set pGeom = pPolylineArray.Element(lngIndex)
    Set pBuffer.Shape = pGeom
    Set pValSubArray = pPolylineValArray.Element(lngIndex)
    For lngIndex2 = 0 To pValSubArray.Count - 1
      lngFieldIndex = lngFieldIndexArray(lngIndex2)
      pBuffer.Value(lngFieldIndex) = pValSubArray.Element(lngIndex2)
    Next lngIndex2
    pFCursor.InsertFeature pBuffer
  Next lngIndex
    
  pFCursor.Flush
  
  Dim pFLayer As IFeatureLayer
  Set pFLayer = New FeatureLayer
  pFLayer.Name = strName
  Set pFLayer.FeatureClass = pFClass
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  pMxDoc.AddLayer pFLayer
  
 
ClearMemory:
  Set pField = Nothing
  Set pClone = Nothing
  Set pCloneFieldArray = Nothing
  Set pPolygon = Nothing
  Set pPolyline = Nothing
  Set pSpRef = Nothing
  Set pFClass = Nothing
  Erase lngFieldIndexArray
  Set pFCursor = Nothing
  Set pBuffer = Nothing
  Set pValSubArray = Nothing
  Set pGeom = Nothing
  Set pFLayer = Nothing
  Set pMxDoc = Nothing




End Sub


Public Sub AddToPolylineArray(pPolylineArray As esriSystem.IArray, strName As String, strCaveOrRandom As String, _
  pPolylineValArray As esriSystem.IVariantArray, pGeomArray As esriSystem.IArray, dblXShift As Double, dblYSHift As Double)
    
  Dim pPolylineSubArray As esriSystem.IVariantArray
  Dim lngIndex As Long
  Dim pPolyline As IPolyline
  Dim pTransform2D As ITransform2D
  Set pPolylineValArray = New esriSystem.varArray
  Set pGeomArray = New esriSystem.Array
  
  For lngIndex = 0 To pPolylineArray.Count - 1
    Set pPolylineSubArray = New esriSystem.varArray
    pPolylineSubArray.Add strName
    pPolylineSubArray.Add strCaveOrRandom
    pPolylineValArray.Add pPolylineSubArray
    Set pPolyline = pPolylineArray.Element(lngIndex)
    Set pTransform2D = pPolyline
'    pTransform2D.Move 400004 + dblXShift, 3500006 + dblYSHift
    pTransform2D.Move dblXShift, dblYSHift
    pGeomArray.Add pPolyline
  Next lngIndex

ClearMemory:
  Set pPolylineSubArray = Nothing
  Set pPolyline = Nothing
  Set pTransform2D = Nothing


End Sub

Public Sub ConvertToPointArray(dblVals() As Double, dblXShift As Double, dblYSHift As Double, dblGraphMax As Double, _
    strName As String, strCaveRandom As String, pMap As IMap, dblYMin As Double, dblYMax As Double, dblYMean As Double, _
    pWS As IWorkspace)
  
  Dim pValArray As esriSystem.IVariantArray
  Dim pValSubArray As esriSystem.IVariantArray
  Dim pPolyArray As esriSystem.IArray
  Dim pFieldArray As esriSystem.IVariantArray
  Set pFieldArray = New esriSystem.varArray
  
  Dim dblIncrement As Double
  dblIncrement = 8 / UBound(dblVals, 1)
  
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Name"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Cave_Or_Random"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Count"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "X"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Y"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pPolyArray = New esriSystem.Array
  Set pValArray = New esriSystem.varArray
  Dim pPoly As IPolygon
  Dim pEnv As IEnvelope
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  
  Dim dblIndexX As Double
  Dim dblIndexY As Double
  
  Dim dblAdjustMin As Double
  Dim dblAdjustMax As Double
  Dim dblAdjustRange As Double
  dblAdjustRange = dblYMax - dblYMin
  dblAdjustMin = dblYMin - dblYMean
  dblAdjustMax = dblYMax - dblYMean
  
  For dblIndexY = 0 To UBound(dblVals, 2)
    
    For dblIndexX = 0 To UBound(dblVals, 1)
      Set pEnv = New Envelope
      Set pEnv.SpatialReference = pSpRef
      pEnv.PutCoords (dblIndexX * dblIncrement) + dblXShift + 400000, (dblIndexY * dblIncrement) + dblYSHift + 3500000, _
            ((dblIndexX * dblIncrement) + dblIncrement) + dblXShift + 400000, _
            ((dblIndexY * dblIncrement) + dblIncrement) + dblYSHift + 3500000
      Set pPoly = MyGeometricOperations.EnvelopeToPolygon(pEnv)
      pPolyArray.Add pPoly
      
      Set pValSubArray = New esriSystem.varArray
      pValSubArray.Add strName
      pValSubArray.Add strCaveRandom
      pValSubArray.Add dblVals(dblIndexX, dblIndexY)
      pValSubArray.Add (dblIndexX * dblIncrement) - 4
      pValSubArray.Add ((dblIndexY / UBound(dblVals, 2)) * dblAdjustRange) + dblAdjustMin
      pValArray.Add pValSubArray
      
    Next dblIndexX
  Next dblIndexY
  
  CreatePolylineFClass pWS, Replace(strName, " ", "_") & "_" & Replace(strCaveRandom, " ", "_") & "_Cells", _
      pPolyArray, pValArray, pFieldArray, False
  
'  Dim pFClass As IFeatureClass
'  Set pFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pPolyArray, pValArray, pFieldArray)
'  Dim pFLayer As IFeatureLayer
'  Set pFLayer = New FeatureLayer
'  Set pFLayer.FeatureClass = pFClass
'  pFLayer.Name = strName & ", " & strCaveRandom
'  pFLayer.Visible = True
'  pMap.AddLayer pFLayer
  
  
ClearMemory:
  Set pValArray = Nothing
  Set pValSubArray = Nothing
  Set pPolyArray = Nothing
  Set pFieldArray = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pPoly = Nothing
  Set pEnv = Nothing
  Set pSpRef = Nothing

End Sub

Public Function FillGraphArray2(dblVals() As Double, dblGraphMax As Double, pMxDoc As IMxDocument, _
      pPolylineArray As esriSystem.IArray, pMeanPolyline As IPolyline, pStDevPolygon As IPolygon, _
      pConfIntPolygon As IPolygon, dblXOffset As Double, dblYOffset As Double) As Double()
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim dblOutput() As Double
  Dim booFullyContained As Boolean
  Dim booIntersects As Boolean
  ReDim dblOutput(180, 90)
  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblRectMinX As Double
  Dim dblRectMinY As Double
  Dim dblRectMaxX As Double
  Dim dblRectMaxY As Double
  dblGraphMax = 0
  Set pPolylineArray = New esriSystem.Array
  
  Dim dblIncrement As Double
  dblIncrement = 8 / UBound(dblOutput, 1)
  
  Dim dblIndexX As Double
  Dim dblIndexY As Double
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  
  Dim pPolyline As IPointCollection
  Dim pPoint As IPoint
  Dim pGeom As IGeometry
  Dim lngIndexCounter As Long
  Dim lngIndexCount As Long
  lngIndexCount = UBound(dblOutput, 2) * UBound(dblOutput, 1)
  
  
  
  ' GET GENERAL LINE STATISTICS AND MAKE LINES
  Dim dblMeans() As Double
  Dim dblStDevs() As Double
  Dim dblStatVals() As Double
  ReDim dblMeans(UBound(dblVals, 1))            ' ONE MEAN VALUE FOR EVERY VERTEX
  ReDim dblStDevs(UBound(dblVals, 1))           ' ONE STANDARD DEVIATION VALUE FOR EVERY VERTEX
  ReDim dblStatVals(UBound(dblVals, 2))             ' ONE VALUE TO (CALCULATE MEAN AND STANDARD DEVIATION) FOR EVERY LINE
  Dim dblMean As Double
  Dim dblStDev As Double
  
  For lngIndex1 = 0 To UBound(dblVals, 1)       ' CHECK EACH SEGMENT ENDPOINT
    For lngIndex2 = 0 To UBound(dblVals, 2)     ' check each line
      dblY1 = dblVals(lngIndex1, lngIndex2)     ' THIS IS THE Y-VALUE AFTER LINE HAS BEEN SHIFTED TO MEAN
      dblStatVals(lngIndex2) = dblY1
    Next lngIndex2
    Call MyGeneralOperations.BasicStatsFromArraySimpleFast2(dblStatVals, True, , , , , dblMean, , , dblStDev)
    dblMeans(lngIndex1) = dblMean
    dblStDevs(lngIndex1) = dblStDev
  Next lngIndex1
  
  Dim lngIndex As Long
  
'  Dim pMeanPolyline As IPolyline
  Dim pMeanPolylinePts As IPointCollection
  Set pMeanPolyline = New Polyline
  Set pMeanPolylinePts = pMeanPolyline
  Set pGeom = pMeanPolyline
  Set pGeom.SpatialReference = pSpRef
  
'  Dim pStDevPolygon As IPolygon
  Dim pStDevPolygonPts As IPointCollection
  Set pStDevPolygon = New Polygon
  Set pStDevPolygonPts = pStDevPolygon
  Set pGeom = pStDevPolygon
  Set pGeom.SpatialReference = pSpRef
  
'  Dim pConfIntPolygon As IPolygon
  Dim pConfIntPolygonPts As IPointCollection
  Set pConfIntPolygon = New Polygon
  Set pConfIntPolygonPts = pConfIntPolygon
  Set pGeom = pConfIntPolygon
  Set pGeom.SpatialReference = pSpRef
  
  Dim dblYConfidence As Double
  
  For lngIndex = 0 To UBound(dblMeans)
    dblX1 = (CDbl(lngIndex) * 0.25) + 400000
    dblY1 = dblMeans(lngIndex) + 3500000
    dblStDev = dblStDevs(lngIndex)
    
    dblYConfidence = 1.96 * (dblStDev / Sqr(UBound(dblStatVals) + 1))
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1
    pMeanPolylinePts.AddPoint pPoint
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 + dblStDev
    pStDevPolygonPts.AddPoint pPoint
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 + dblYConfidence
    pConfIntPolygonPts.AddPoint pPoint
  Next lngIndex
  
  ' REVERSE TO GET LOWER BOUND ON POLYGON
  For lngIndex = UBound(dblMeans) To 0 Step -1
    dblX1 = (CDbl(lngIndex) * 0.25) + 400000
    dblY1 = dblMeans(lngIndex) + 3500000
    dblStDev = dblStDevs(lngIndex)
    
    dblYConfidence = 1.96 * (dblStDev / Sqr(UBound(dblStatVals) + 1))
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 - dblStDev
    pStDevPolygonPts.AddPoint pPoint
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 - dblYConfidence
    pConfIntPolygonPts.AddPoint pPoint
  Next lngIndex
  
  pStDevPolygon.Close
  pStDevPolygon.SimplifyPreserveFromTo
  pConfIntPolygon.Close
  pConfIntPolygon.SimplifyPreserveFromTo
      
  Dim pTransform2D As ITransform2D
  Set pTransform2D = pMeanPolyline
  pTransform2D.Move dblXOffset, dblYOffset
  Set pTransform2D = pStDevPolygon
  pTransform2D.Move dblXOffset, dblYOffset
  Set pTransform2D = pConfIntPolygon
  pTransform2D.Move dblXOffset, dblYOffset
  
  ' MAKE POLYLINES
  For lngIndex1 = 0 To UBound(dblVals, 2)
    If UBound(dblVals, 2) > 10000 Then
      If lngIndex1 Mod 10000 = 0 Then Debug.Print "Making Polylines...." & CStr(lngIndex1) & " of " & CStr(UBound(dblVals, 2))
    Else
      If lngIndex1 Mod 1000 = 0 Then Debug.Print "Making Polylines...." & CStr(lngIndex1) & " of " & CStr(UBound(dblVals, 2))
    End If
    DoEvents
'    If lngIndex1 > 8 Then Exit For
    
    Set pPolyline = New Polyline
    Set pGeom = pPolyline
    Set pGeom.SpatialReference = pSpRef
    
    For lngIndex2 = 0 To UBound(dblVals, 1) - 1
    
      If lngIndex2 = 1 Then
        DoEvents
      End If
      
      dblX1 = (CDbl(lngIndex2) * 0.25) + 400000
      dblY1 = dblVals(lngIndex2, lngIndex1) + 3500000
      dblX2 = (CDbl(lngIndex2) * 0.25) + 0.25 + 400000
      dblY2 = dblVals(lngIndex2 + 1, lngIndex1) + 3500000
      
'      Debug.Print CStr(lngIndex2) & "] dblX1 = " & CStr(dblX1) & ", dblX2 = " & CStr(dblX2)
      
      Set pPoint = New Point
      pPoint.PutCoords dblX1, dblY1
      pPolyline.AddPoint pPoint
      
'      Debug.Print "Row " & CStr(lngIndex1) & ", Column " & CStr(lngIndex2) & ": " & CStr(dblOutput(lngIndex2, lngIndex1))
'      Debug.Print "     Segment:  [" & Format(dblX1, "0.00") & "," & Format(dblY1, "0.00") & "] to [" & _
'            Format(dblX2, "0.00") & "," & Format(dblY2, "0.00") & "]"
      
    Next lngIndex2
    
    Set pPoint = New Point
    pPoint.PutCoords dblX2, dblY2
    pPolyline.AddPoint pPoint
    pPolylineArray.Add pPolyline
    
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolyline
  Next lngIndex1
  
  Debug.Print "Assembling output array..."
  
'  ' COMMENT OUT BELOW IF ONLY GENERATING LINES AND LINE STATISTICS
'  For dblIndexY = 0 To UBound(dblOutput, 2)
'    For dblIndexX = 0 To UBound(dblOutput, 1)
'      lngIndexCounter = lngIndexCounter + 1
'      If lngIndexCounter Mod 200 = 0 Then
'        Debug.Print "Analyzing Array...." & Format(lngIndexCounter, "#,##0") & " of " & Format(lngIndexCount, "#,##0")
'        DoEvents
'      End If
'
'      dblRectMinX = (dblIndexX * dblIncrement) ' + 400000
'      dblRectMaxX = (dblIndexX * dblIncrement) + dblIncrement ' + 400000
'      dblRectMinY = (dblIndexY * dblIncrement) ' + 3500000
'      dblRectMaxY = (dblIndexY * dblIncrement) + dblIncrement ' + 3500000
'
'      ' Check each line until find intersection or end of line
'
'      For lngIndex1 = 0 To UBound(dblVals, 2)
'        DoEvents
'        For lngIndex2 = 0 To UBound(dblVals, 1) - 1
'
'          dblX1 = (CDbl(lngIndex2) * 0.25) ' + 400000
'          dblY1 = dblVals(lngIndex2, lngIndex1) ' + 3500000
'          dblX2 = (CDbl(lngIndex2) * 0.25) + 0.25 ' + 400000
'          dblY2 = dblVals(lngIndex2 + 1, lngIndex1) ' + 3500000
'
'          booIntersects = MyGeometricOperations.SegmentIntersectsRectangle(dblX1, dblY1, dblX2, dblY2, dblRectMinX, _
'                dblRectMaxX, dblRectMinY, dblRectMaxY, booFullyContained)
'
'          If booIntersects Or booFullyContained Then
'            dblOutput(dblIndexX, dblIndexY) = dblOutput(dblIndexX, dblIndexY) + 1
'            If dblOutput(dblIndexX, dblIndexY) > dblGraphMax Then dblGraphMax = dblOutput(dblIndexX, dblIndexY)
'            Exit For ' MOVE ON TO NEXT LINE
'          End If
'
'        Next lngIndex2 ' MOVE ON TO NEXT SEGMENT IN LINE
'      Next lngIndex1   ' MOVE ON TO NEXT LINE
'    Next dblIndexX     ' MOVE ON TO NEXT COLUMN IN OUTPUT ARRAY
'  Next dblIndexY       ' MOVE ON TO NEXT ROW IN OUTPUT ARRAY
  
  FillGraphArray2 = dblOutput
  
ClearMemory:
  Erase dblOutput
  Set pSpRef = Nothing
  Set pPolyline = Nothing
  Set pPoint = Nothing
  Set pGeom = Nothing
  Erase dblMeans
  Erase dblStDevs
  Erase dblStatVals
  Set pMeanPolylinePts = Nothing
  Set pStDevPolygonPts = Nothing
  Set pConfIntPolygonPts = Nothing
  Set pTransform2D = Nothing


End Function


Public Function FillGraphArray2_ExcludeNulls(varVals() As Variant, dblGraphMax As Double, pMxDoc As IMxDocument, _
      pPolylineArray As esriSystem.IArray, pMeanPolyline As IPolyline, pStDevPolygon As IPolygon, _
      pConfIntPolygon As IPolygon, dblXOffset As Double, dblYOffset As Double) As Double()
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim dblOutput() As Double
  Dim booFullyContained As Boolean
  Dim booIntersects As Boolean
  ReDim dblOutput(180, 90)
  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblRectMinX As Double
  Dim dblRectMinY As Double
  Dim dblRectMaxX As Double
  Dim dblRectMaxY As Double
  dblGraphMax = 0
  Set pPolylineArray = New esriSystem.Array
  
  Dim dblIncrement As Double
  dblIncrement = 8 / UBound(dblOutput, 1)
  
  Dim dblIndexX As Double
  Dim dblIndexY As Double
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  
  Dim pPolyline As IPointCollection
  Dim pPoint As IPoint
  Dim pGeom As IGeometry
  Dim lngIndexCounter As Long
  Dim lngIndexCount As Long
  lngIndexCount = UBound(dblOutput, 2) * UBound(dblOutput, 1)
  
  Dim strRow As String
  
  
  ' GET GENERAL LINE STATISTICS AND MAKE LINES
  Dim dblMeans() As Double
  Dim dblStDevs() As Double
  Dim dblStatVals() As Double
  ReDim dblMeans(UBound(varVals, 1))            ' ONE MEAN VALUE FOR EVERY VERTEX
  ReDim dblStDevs(UBound(varVals, 1))           ' ONE STANDARD DEVIATION VALUE FOR EVERY VERTEX
  ReDim dblStatVals(UBound(varVals, 2))             ' ONE VALUE TO (CALCULATE MEAN AND STANDARD DEVIATION) FOR EVERY LINE
  Dim dblMean As Double
  Dim dblStDev As Double
  Dim varVal As Variant
  Dim lngCounter As Long
  
  For lngIndex1 = 0 To UBound(varVals, 1)       ' CHECK EACH SEGMENT ENDPOINT
    lngCounter = -1
    ReDim dblStatVals(UBound(varVals, 2))
    
'    strRow = ""
    
    For lngIndex2 = 0 To UBound(varVals, 2)     ' check each line
      varVal = varVals(lngIndex1, lngIndex2)     ' THIS IS THE Y-VALUE AFTER LINE HAS BEEN SHIFTED TO MEAN
      
'      strRow = strRow & IIf(IsNull(varVals(lngIndex1, lngIndex2)), " Null, ", _
'          Format(varVals(lngIndex1, lngIndex2), "0.0") & ", ")
          
      If Not IsNull(varVal) Then
        lngCounter = lngCounter + 1
        dblY1 = CDbl(varVal)                     ' THIS IS THE Y-VALUE AFTER LINE HAS BEEN SHIFTED TO MEAN
        dblStatVals(lngCounter) = dblY1
      End If
    Next lngIndex2
    
'    Debug.Print strRow
    
    ReDim Preserve dblStatVals(lngCounter)
    Call MyGeneralOperations.BasicStatsFromArraySimpleFast2(dblStatVals, True, , , , , dblMean, , , dblStDev)
    dblMeans(lngIndex1) = dblMean
    dblStDevs(lngIndex1) = dblStDev
  Next lngIndex1
  
  Dim lngIndex As Long
  
'  Dim pMeanPolyline As IPolyline
  Dim pMeanPolylinePts As IPointCollection
  Set pMeanPolyline = New Polyline
  Set pMeanPolylinePts = pMeanPolyline
  Set pGeom = pMeanPolyline
  Set pGeom.SpatialReference = pSpRef
  
'  Dim pStDevPolygon As IPolygon
  Dim pStDevPolygonPts As IPointCollection
  Set pStDevPolygon = New Polygon
  Set pStDevPolygonPts = pStDevPolygon
  Set pGeom = pStDevPolygon
  Set pGeom.SpatialReference = pSpRef
  
'  Dim pConfIntPolygon As IPolygon
  Dim pConfIntPolygonPts As IPointCollection
  Set pConfIntPolygon = New Polygon
  Set pConfIntPolygonPts = pConfIntPolygon
  Set pGeom = pConfIntPolygon
  Set pGeom.SpatialReference = pSpRef
  
  Dim dblYConfidence As Double
  
  For lngIndex = 0 To UBound(dblMeans)
    dblX1 = (CDbl(lngIndex) * 0.25) + 400000
    dblY1 = dblMeans(lngIndex) + 3500000
    dblStDev = dblStDevs(lngIndex)
    
    dblYConfidence = 1.96 * (dblStDev / Sqr(UBound(dblStatVals) + 1))
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1
    pMeanPolylinePts.AddPoint pPoint
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 + dblStDev
    pStDevPolygonPts.AddPoint pPoint
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 + dblYConfidence
    pConfIntPolygonPts.AddPoint pPoint
  Next lngIndex
  
  ' REVERSE TO GET LOWER BOUND ON POLYGON
  For lngIndex = UBound(dblMeans) To 0 Step -1
    dblX1 = (CDbl(lngIndex) * 0.25) + 400000
    dblY1 = dblMeans(lngIndex) + 3500000
    dblStDev = dblStDevs(lngIndex)
    
    dblYConfidence = 1.96 * (dblStDev / Sqr(UBound(dblStatVals) + 1))
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 - dblStDev
    pStDevPolygonPts.AddPoint pPoint
    
    Set pPoint = New Point
    pPoint.PutCoords dblX1, dblY1 - dblYConfidence
    pConfIntPolygonPts.AddPoint pPoint
  Next lngIndex
  
  pStDevPolygon.Close
  pStDevPolygon.SimplifyPreserveFromTo
  pConfIntPolygon.Close
  pConfIntPolygon.SimplifyPreserveFromTo
      
  Dim pTransform2D As ITransform2D
  Set pTransform2D = pMeanPolyline
  pTransform2D.Move dblXOffset, dblYOffset
  Set pTransform2D = pStDevPolygon
  pTransform2D.Move dblXOffset, dblYOffset
  Set pTransform2D = pConfIntPolygon
  pTransform2D.Move dblXOffset, dblYOffset
  
  Dim booFoundTwoPoints As Boolean
  
  ' MAKE POLYLINES
  For lngIndex1 = 0 To UBound(varVals, 2)
    If UBound(varVals, 2) > 10000 Then
      If lngIndex1 Mod 10000 = 0 Then Debug.Print "Making Polylines...." & CStr(lngIndex1) & " of " & CStr(UBound(varVals, 2))
    Else
      If lngIndex1 Mod 1000 = 0 Then Debug.Print "Making Polylines...." & CStr(lngIndex1) & " of " & CStr(UBound(varVals, 2))
    End If
    DoEvents
'    If lngIndex1 > 8 Then Exit For
    
    Set pPolyline = New Polyline
    Set pGeom = pPolyline
    Set pGeom.SpatialReference = pSpRef
    
    booFoundTwoPoints = False
'    strRow = ""
    For lngIndex2 = 0 To UBound(varVals, 1) - 1
      
      If lngIndex2 = 1 Then
        DoEvents
      End If
      
'      strRow = strRow & IIf(IsNull(varVals(lngIndex2, lngIndex1)), " Null, ", _
'          Format(varVals(lngIndex2, lngIndex1), "0.0") & ", ")
      
      If Not IsNull(varVals(lngIndex2, lngIndex1)) Then
        dblX1 = (CDbl(lngIndex2) * 0.25) + 400000
        dblY1 = varVals(lngIndex2, lngIndex1) + 3500000
        If Not IsNull(varVals(lngIndex2 + 1, lngIndex1)) Then
          dblX2 = (CDbl(lngIndex2) * 0.25) + 0.25 + 400000
          dblY2 = varVals(lngIndex2 + 1, lngIndex1) + 3500000
          booFoundTwoPoints = True
        End If
'        Debug.Print CStr(lngIndex2) & "] dblX1 = " & CStr(dblX1) & ", dblX2 = " & CStr(dblX2)
        
        Set pPoint = New Point
        pPoint.PutCoords dblX1, dblY1
        pPolyline.AddPoint pPoint
      End If
'      Debug.Print "Row " & CStr(lngIndex1) & ", Column " & CStr(lngIndex2) & ": " & CStr(dblOutput(lngIndex2, lngIndex1))
'      Debug.Print "     Segment:  [" & Format(dblX1, "0.00") & "," & Format(dblY1, "0.00") & "] to [" & _
'            Format(dblX2, "0.00") & "," & Format(dblY2, "0.00") & "]"
      DoEvents
    Next lngIndex2
    
'    Debug.Print strRow
    If booFoundTwoPoints Then
      Set pPoint = New Point
      pPoint.PutCoords dblX2, dblY2
      pPolyline.AddPoint pPoint
      pPolylineArray.Add pPolyline
    End If
    
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolyline
  Next lngIndex1
  
  ' COMMENT OUT BELOW IF ONLY GENERATING LINES AND LINE STATISTICS
  Debug.Print "Assembling output array..."
  
  ' Check each line until find intersection or end of line
  
  Dim booDoneArray() As Boolean
  ReDim booDoneArray(UBound(dblOutput, 1), UBound(dblOutput, 2))
  Dim lngArrayX As Long
  Dim lngArrayY As Long
  
  ' WORK THROUGH EACH LINE
  For lngIndex1 = 0 To UBound(varVals, 2)
    DoEvents
    lngIndexCounter = lngIndexCounter + 1
    If lngIndexCounter Mod 200 = 0 Then
      Debug.Print "Analyzing Array...." & Format(lngIndexCounter, "#,##0") & " of " & Format(UBound(varVals, 2), "#,##0")
      DoEvents
    End If
    
    ResetDoneArray booDoneArray
    
    ' WORK THROUGH VALUE ALONG TRANSECT
    For lngIndex2 = 0 To UBound(varVals, 1) - 1
    
      ' CONFIRM THAT VALUE AND NEXT VALUE ARE NOT NULL
      If Not IsNull(varVals(lngIndex2, lngIndex1)) And Not IsNull(varVals(lngIndex2 + 1, lngIndex1)) Then
        
        ' GET X- AND Y- VALUES OF START AND END OF LINE
        
        dblX1 = (CDbl(lngIndex2) * 0.25) ' + 400000
        dblY1 = varVals(lngIndex2, lngIndex1) ' + 3500000
        dblX2 = (CDbl(lngIndex2) * 0.25) + 0.25 ' + 400000
        dblY2 = varVals(lngIndex2 + 1, lngIndex1) ' + 3500000
        
        ' WORK THROUGH RECTANGLE CONTAINING THESE VALUES
        For dblIndexY = dblY1 To dblY2 Step 1    ' 0 To UBound(dblOutput, 2)
          For dblIndexX = dblX1 To dblX2 Step 1  ' 0 To UBound(dblOutput, 1)
            
            ' ONLY CHECK AGAINST RECTANGLE IF CELL CURRENTLY NOT MARKED YET FOR THIS LINE
            lngArrayX = Int(dblIndexX)
            lngArrayY = Int(dblIndexY)
            
            If Not booDoneArray(lngArrayX, lngArrayY) Then
              dblRectMinX = (dblIndexX * dblIncrement) ' + 400000
              dblRectMaxX = (dblIndexX * dblIncrement) + dblIncrement ' + 400000
              dblRectMinY = (dblIndexY * dblIncrement) ' + 3500000
              dblRectMaxY = (dblIndexY * dblIncrement) + dblIncrement ' + 3500000
        
          
              booIntersects = MyGeometricOperations.SegmentIntersectsRectangle(dblX1, dblY1, dblX2, dblY2, dblRectMinX, _
                    dblRectMaxX, dblRectMinY, dblRectMaxY, booFullyContained)
    
              If booIntersects Or booFullyContained Then
                booDoneArray(lngArrayX, lngArrayY) = True
                dblOutput(lngArrayX, lngArrayY) = dblOutput(lngArrayX, lngArrayY) + 1
                If dblOutput(lngArrayX, lngArrayY) > dblGraphMax Then dblGraphMax = dblOutput(lngArrayX, lngArrayY)
              End If
            End If
          Next dblIndexX     ' MOVE ON TO NEXT COLUMN IN OUTPUT ARRAY
        Next dblIndexY       ' MOVE ON TO NEXT ROW IN OUTPUT ARRAY
      End If
    Next lngIndex2 ' MOVE ON TO NEXT SEGMENT IN LINE
  Next lngIndex1   ' MOVE ON TO NEXT LINE
  
  FillGraphArray2_ExcludeNulls = dblOutput
  
ClearMemory:
  Erase dblOutput
  Set pSpRef = Nothing
  Set pPolyline = Nothing
  Set pPoint = Nothing
  Set pGeom = Nothing
  Erase dblMeans
  Erase dblStDevs
  Erase dblStatVals
  Set pMeanPolylinePts = Nothing
  Set pStDevPolygonPts = Nothing
  Set pConfIntPolygonPts = Nothing
  Set pTransform2D = Nothing


End Function

Public Sub ResetDoneArray(booDoneArray() As Boolean)

  Dim lngIndex As Long
  Dim lngIndex2 As Long
  For lngIndex = 0 To UBound(booDoneArray, 1)
    For lngIndex2 = 1 To UBound(booDoneArray, 2)
      booDoneArray(lngIndex, lngIndex2) = False
    Next lngIndex2
  Next lngIndex

End Sub




Public Function FillGraphArray(dblVals() As Double, dblGraphMax As Double, pMxDoc As IMxDocument, _
      pPolylineArray As esriSystem.IArray) As Double()
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim dblOutput() As Double
  Dim booFullyContained As Boolean
  Dim booIntersects As Boolean
  ReDim dblOutput(120, 60)
  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblRectMinX As Double
  Dim dblRectMinY As Double
  Dim dblRectMaxX As Double
  Dim dblRectMaxY As Double
  dblGraphMax = 0
  Set pPolylineArray = New esriSystem.Array
  
  Dim dblIncrement As Double
  dblIncrement = 8 / UBound(dblOutput, 1)
  
  Dim dblIndexX As Double
  Dim dblIndexY As Double
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  
  Dim pPolyline As IPointCollection
  Dim pPoint As IPoint
  Dim pGeom As IGeometry
  
  For lngIndex1 = 0 To UBound(dblVals, 2)
    If lngIndex1 Mod 100 = 0 Then Debug.Print ".." & CStr(lngIndex1) & " of " & CStr(UBound(dblVals, 2))
    DoEvents
'    If lngIndex1 > 8 Then Exit For
    
    Set pPolyline = New Polyline
    Set pGeom = pPolyline
    Set pGeom.SpatialReference = pSpRef
    
    For lngIndex2 = 0 To UBound(dblVals, 1) - 1
    
      If lngIndex2 = 1 Then
        DoEvents
      End If
      
      dblX1 = (CDbl(lngIndex2) * 0.25) + 400000
      dblY1 = dblVals(lngIndex2, lngIndex1) + 3500000
      dblX2 = (CDbl(lngIndex2) * 0.25) + 0.25 + 400000
      dblY2 = dblVals(lngIndex2 + 1, lngIndex1) + 3500000
      
'      Debug.Print CStr(lngIndex2) & "] dblX1 = " & CStr(dblX1) & ", dblX2 = " & CStr(dblX2)
      
      Set pPoint = New Point
      pPoint.PutCoords dblX1, dblY1
      pPolyline.AddPoint pPoint
      
      ' COMMENT OUT BELOW IF ONLY GENERATING LINES
      
      For dblIndexY = 0 To UBound(dblOutput, 2)
        For dblIndexX = 0 To UBound(dblOutput, 1)
          dblRectMinX = (dblIndexX * dblIncrement) + 400000
          dblRectMaxX = (dblIndexX * dblIncrement) + dblIncrement + 400000
          dblRectMinY = (dblIndexY * dblIncrement) + 3500000
          dblRectMaxY = (dblIndexY * dblIncrement) + dblIncrement + 3500000
          booIntersects = MyGeometricOperations.SegmentIntersectsRectangle(dblX1, dblY1, dblX2, dblY2, dblRectMinX, _
                dblRectMaxX, dblRectMinY, dblRectMaxY, booFullyContained)
          'booIntersects = MyGeometricOperations.SegmentIntersectsRectangle(dblX1, dblY1, dblX2, dblY2, -4 + (dblIndexX / 10), _
                -4 + ((dblIndexX + 1) / 10), 4 - (dblIndexY + 1 / 10), 4 - ((dblIndexY) / 10), booFullyContained)
          If booIntersects Or booFullyContained Then
            dblOutput(dblIndexX, dblIndexY) = dblOutput(dblIndexX, dblIndexY) + 1
            If dblOutput(dblIndexX, dblIndexY) > dblGraphMax Then dblGraphMax = dblOutput(dblIndexX, dblIndexY)
          End If
        Next dblIndexX
      Next dblIndexY
      
'      Debug.Print "Row " & CStr(lngIndex1) & ", Column " & CStr(lngIndex2) & ": " & CStr(dblOutput(lngIndex2, lngIndex1))
'      Debug.Print "     Segment:  [" & Format(dblX1, "0.00") & "," & Format(dblY1, "0.00") & "] to [" & _
'            Format(dblX2, "0.00") & "," & Format(dblY2, "0.00") & "]"
      
    Next lngIndex2
    
    Set pPoint = New Point
    pPoint.PutCoords dblX2, dblY2
    pPolyline.AddPoint pPoint
    pPolylineArray.Add pPolyline
    
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolyline
  Next lngIndex1
  
  FillGraphArray = dblOutput
 
ClearMemory:
  Erase dblOutput
  Set pSpRef = Nothing
  Set pPolyline = Nothing
  Set pPoint = Nothing
  Set pGeom = Nothing
 
End Function

Public Function RescaleTo4High(dblVals() As Double, dblMin As Double, dblMax As Double) As Double()
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim dblCentralVal As Double
  Dim dblCumulative As Double
  Dim dblReturn() As Double
  Dim dblNewVal As Double
  Dim dblRange As Double
  dblRange = dblMax - dblMin
  
  ReDim dblReturn(UBound(dblVals, 1), UBound(dblVals, 2))
  
  For lngIndex = 0 To UBound(dblVals, 2)
    For lngIndex2 = 0 To UBound(dblVals, 1)
      dblVal = dblVals(lngIndex2, lngIndex)
      dblNewVal = ((dblVal - dblMin) / dblRange) * 4
      
      dblReturn(lngIndex2, lngIndex) = dblNewVal
    Next lngIndex2
  Next lngIndex
  
  RescaleTo4High = dblReturn
  
ClearMemory:
  Erase dblReturn

End Function

Public Function RescaleTo4High_ExcludeNulls(varVals() As Variant, dblMin As Double, _
      dblMax As Double) As Variant()
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim varVal As Variant
  Dim dblCentralVal As Double
  Dim dblCumulative As Double
  Dim varReturn() As Variant
  Dim dblNewVal As Double
  Dim dblRange As Double
  dblRange = dblMax - dblMin
  
  ReDim varReturn(UBound(varVals, 1), UBound(varVals, 2))
  
  For lngIndex = 0 To UBound(varVals, 2)
    For lngIndex2 = 0 To UBound(varVals, 1)
      varVal = varVals(lngIndex2, lngIndex)
      If IsNull(varVal) Then
        varReturn(lngIndex2, lngIndex) = Null
      Else
        dblNewVal = ((CDbl(varVal) - dblMin) / dblRange) * 4
        varReturn(lngIndex2, lngIndex) = dblNewVal
      End If
    Next lngIndex2
  Next lngIndex
  
  RescaleTo4High_ExcludeNulls = varReturn
  
ClearMemory:
  Erase varReturn

End Function

Public Function ShiftToMean(dblVals() As Double, dblMean As Double, dblMin As Double, dblMax As Double, booIsFirst As Boolean) As Double()
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim dblCentralVal As Double
  Dim dblCumulative As Double
  Dim dblShift As Double
  Dim dblReturn() As Double
  Dim dblNewVal As Double
  ReDim dblReturn(UBound(dblVals, 1), UBound(dblVals, 2))
  
  If booIsFirst Then
    dblMin = dblVals(0, 0)
    dblMax = dblMin
  End If
  
  For lngIndex = 0 To UBound(dblVals, 2)
    dblCentralVal = dblVals(16, lngIndex)
    dblShift = dblMean - dblCentralVal
    For lngIndex2 = 0 To UBound(dblVals, 1)
      dblVal = dblVals(lngIndex2, lngIndex)
      dblNewVal = dblVal + dblShift
      If dblNewVal < dblMin Then dblMin = dblNewVal
      If dblNewVal > dblMax Then dblMax = dblNewVal
      dblReturn(lngIndex2, lngIndex) = dblNewVal
    Next lngIndex2
  Next lngIndex
  
  ShiftToMean = dblReturn
  
ClearMemory:
  Erase dblReturn

End Function
Public Function ShiftToMean_ExcludeNulls(varVals() As Variant, dblMean As Double, _
    dblMin As Double, dblMax As Double, booIsFirst As Boolean) As Variant()
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim varVal As Variant
  Dim dblCentralVal As Double
  Dim dblCumulative As Double
  Dim dblShift As Double
  Dim varReturn() As Variant
  Dim dblNewVal As Double
  ReDim varReturn(UBound(varVals, 1), UBound(varVals, 2))
  
  Dim booFoundVal As Boolean
  booFoundVal = False
  
'  If booIsFirst Then
'    For lngIndex = 0 To UBound(varVals, 2)
'      For lngIndex2 = 0 To UBound(varVals, 1)
'        If Not IsNull(varVals(lngIndex2, lngIndex)) Then
'          booFoundVal = True
'          dblMin = CDbl(varVals(lngIndex2, lngIndex))
'          dblMax = dblMin
'          Exit For
'        End If
'      Next lngIndex2
'      If booFoundVal Then Exit For
'    Next lngIndex
'  End If
  
  For lngIndex = 0 To UBound(varVals, 2)
    varVal = varVals(16, lngIndex)
    If IsNull(varVal) Then   ' IF WE CAN'T PLACE THE CENTRAL CELL, THEN IGNORE ENTIRE LINE
      For lngIndex2 = 0 To UBound(varVals, 1)
        varReturn(lngIndex2, lngIndex) = Null
      Next lngIndex2
    Else
      dblCentralVal = CDbl(varVal)
      dblShift = dblMean - dblCentralVal
      For lngIndex2 = 0 To UBound(varVals, 1)
        varVal = varVals(lngIndex2, lngIndex)
        If IsNumeric(varVal) Then
          
          dblNewVal = CDbl(varVal) + dblShift
          
          If booFoundVal Then
            If dblNewVal < dblMin Then dblMin = dblNewVal
            If dblNewVal > dblMax Then dblMax = dblNewVal
          Else
            dblMin = dblNewVal
            dblMax = dblNewVal
          End If
          varReturn(lngIndex2, lngIndex) = dblNewVal
          booFoundVal = True
        Else
          varReturn(lngIndex2, lngIndex) = Null
        End If
      Next lngIndex2
    End If
  Next lngIndex
  
  ShiftToMean_ExcludeNulls = varReturn
  
ClearMemory:
  Erase varReturn

End Function
Public Function ReturnMeanValue(dblVals() As Double) As Double
  Dim lngIndex As Long
  Dim dblCentralVal As Double
  Dim dblCumulative As Double
  
  For lngIndex = 0 To UBound(dblVals, 2)
    dblCentralVal = dblVals(16, lngIndex)
    dblCumulative = dblCumulative + dblCentralVal
'    Debug.Print CStr(lngIndex + 1) & "] Value = " & CStr(dblVal) & ", Cumulative = " & CStr(dblCumulative)
  Next lngIndex
'  Debug.Print "Mean = [" & CStr(dblCumulative) & " / " & CStr((CDbl(UBound(dblCaveVals, 2) + 1))) & "] = " & _
'    CStr(dblCumulative / (CDbl(UBound(dblCaveVals, 2) + 1)))
  ReturnMeanValue = dblCumulative / (CDbl(UBound(dblVals, 2) + 1))
End Function

Public Function ReturnMeanValue_ExcludeNulls(varVals() As Variant) As Variant
  Dim lngIndex As Long
  Dim dblCumulative As Double
  Dim varCentralVal As Variant
  Dim dblCounter As Double
  dblCounter = 0
  
  Dim dblCheckMin As Double
  Dim dblCheckMax As Double
  Dim booFound As Boolean
  
  For lngIndex = 0 To UBound(varVals, 2)
    varCentralVal = varVals(16, lngIndex)
    If Not IsNull(varCentralVal) Then
      dblCumulative = dblCumulative + CDbl(varCentralVal)
      dblCounter = dblCounter + 1
      
      If booFound Then
        If CDbl(varCentralVal) > dblCheckMax Then dblCheckMax = CDbl(varCentralVal)
        If CDbl(varCentralVal) < dblCheckMin Then dblCheckMin = CDbl(varCentralVal)
      Else
        dblCheckMax = CDbl(varCentralVal)
        dblCheckMin = CDbl(varCentralVal)
        booFound = True
      End If
      
    End If
'    Debug.Print CStr(lngIndex + 1) & "] Value = " & CStr(dblVal) & ", Cumulative = " & CStr(dblCumulative)
  Next lngIndex
'  Debug.Print "Mean = [" & CStr(dblCumulative) & " / " & CStr((CDbl(UBound(dblCaveVals, 2) + 1))) & "] = " & _
'    CStr(dblCumulative / (CDbl(UBound(dblCaveVals, 2) + 1)))
  ReturnMeanValue_ExcludeNulls = dblCumulative / dblCounter
  
End Function

Public Sub AnalyzeLines()
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim pCSVFiles As esriSystem.IStringArray
  Set pCSVFiles = MyGeneralOperations.ReturnFilesFromNestedFolders2( _
      "D:\arcGIS_stuff\consultation\Jut_Wynne\Analysis_Files\October_22_2015_Outputs", ".csv")
      
  Dim lngIndex As Long
  Dim strDayThermal As String
  Dim strDayCurve As String
  Dim strDaySlope As String
  Dim strDayTPI As String
  Dim strNightThermal As String
  Dim strNightCurve As String
  Dim strNightSlope As String
  Dim strNightTPI As String
  Dim strFilename As String
  Dim pTransform2D As ITransform2D
  
  For lngIndex = 0 To pCSVFiles.Count - 1
    strFilename = pCSVFiles.Element(lngIndex)
    If InStr(1, strFilename, "_Day_", vbTextCompare) > 0 Then
      If InStr(1, strFilename, "_Thermal_", vbTextCompare) > 0 Then
        strDayThermal = strFilename
      ElseIf InStr(1, strFilename, "_Slope_", vbTextCompare) > 0 Then
        strDaySlope = strFilename
      ElseIf InStr(1, strFilename, "_Curve_", vbTextCompare) > 0 Then
        strDayCurve = strFilename
      ElseIf InStr(1, strFilename, "_TPI_", vbTextCompare) > 0 Then
        strDayTPI = strFilename
      End If
    Else
      If InStr(1, strFilename, "_Thermal_", vbTextCompare) > 0 Then
        strNightThermal = strFilename
      ElseIf InStr(1, strFilename, "_Slope_", vbTextCompare) > 0 Then
        strNightSlope = strFilename
      ElseIf InStr(1, strFilename, "_Curve_", vbTextCompare) > 0 Then
        strNightCurve = strFilename
      ElseIf InStr(1, strFilename, "_TPI_", vbTextCompare) > 0 Then
        strNightTPI = strFilename
      End If
    End If
  Next lngIndex
  
  Debug.Print "-----------------------------------------"
  Debug.Print "Day Thermal = " & aml_func_mod.ReturnFilename2(strDayThermal)
  Debug.Print "Day Slope = " & aml_func_mod.ReturnFilename2(strDaySlope)
  Debug.Print "Day Curve = " & aml_func_mod.ReturnFilename2(strDayCurve)
  Debug.Print "Day TPI = " & aml_func_mod.ReturnFilename2(strDayTPI)
  Debug.Print "Night Thermal = " & aml_func_mod.ReturnFilename2(strNightThermal)
  Debug.Print "Night Slope = " & aml_func_mod.ReturnFilename2(strNightSlope)
  Debug.Print "Night Curve = " & aml_func_mod.ReturnFilename2(strNightCurve)
  Debug.Print "Night TPI = " & aml_func_mod.ReturnFilename2(strNightTPI)
  
  Dim strText As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String
  Dim dblCaveVals() As Double
  Dim dblRandomVals() As Double
  Dim dblMaxVal As Double
  Dim dblMinVal As Double
  Dim dblCentralVal As Double
  Dim lngCentralIndex As Long
  lngCentralIndex = 23
  Dim lngCaveOrRandomIndex As Long
  lngCaveOrRandomIndex = 5
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim lngCount As Long
  Dim lngInterval As Long
  Dim lngCounter As Long
  Dim strCaveRandom As String
  Dim booFirst As Boolean
  Dim lngCaveCounter As Long
  Dim lngRandomCounter As Long
  Dim dblCaveMean As Double
  Dim dblRandomMean As Double
  Dim dblRescaleCave() As Double
  Dim dblRescaleRandom() As Double
  Dim booIsFirst As Boolean
  Dim dblGraphMax As Double
  Dim dblGraphArray() As Double
  
  Dim pPolylineArray As esriSystem.IArray
  Set pPolylineArray = New esriSystem.Array
  Dim pGeomArray As esriSystem.IArray
  Set pGeomArray = New esriSystem.Array
  Dim pPolylineValArray As esriSystem.IVariantArray
  Set pPolylineValArray = New esriSystem.varArray
  Dim pPolylineSubArray As esriSystem.IVariantArray
  Dim pPolylineFieldArray As esriSystem.IVariantArray
  Set pPolylineFieldArray = New esriSystem.varArray
  Dim pPolylineField As iField
  Dim pPolylineFieldEdit As IFieldEdit
  Dim pPolyline As IPolyline
  
  Dim pGraphMap As IMap
  Set pGraphMap = MyGeneralOperations.ReturnMapByName("Graphs", pMxDoc)
  
  Set pPolylineField = New Field
  Set pPolylineFieldEdit = pPolylineField
  With pPolylineFieldEdit
    .Name = "Name"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pPolylineFieldArray.Add pPolylineField
  
  Set pPolylineField = New Field
  Set pPolylineFieldEdit = pPolylineField
  With pPolylineFieldEdit
    .Name = "Cave_Or_Random"
    .Type = esriFieldTypeString
    .length = 50
  End With
  pPolylineFieldArray.Add pPolylineField
    
  
'  ' DAY THERMAL
'  Call RunDatasetAnalysis(strDayThermal, "Daytime Thermal", 0, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
'       pPolylineValArray, pGeomArray)
'
'  ' DAY Slope
'  Call RunDatasetAnalysis(strDaySlope, "Daytime Slope", 10, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
'       pPolylineValArray, pGeomArray)
'
'  ' DAY curve
'  Call RunDatasetAnalysis(strDayCurve, "Daytime Curvature", 20, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
'       pPolylineValArray, pGeomArray)
'
'  ' DAY TPI
'  Call RunDatasetAnalysis(strDayTPI, "Daytime TPI", 30, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
'       pPolylineValArray, pGeomArray)
'
'
'  ' NIGHT THERMAL
'  Call RunDatasetAnalysis(strNightThermal, "Night Thermal", 40, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
'       pPolylineValArray, pGeomArray)
'
'  ' NIGHT Slope
'  Call RunDatasetAnalysis(strNightSlope, "Night Slope", 50, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
'       pPolylineValArray, pGeomArray)
  
  ' NIGHT curve
  Call RunDatasetAnalysis(strNightCurve, "Night Curvature", 60, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
       pPolylineValArray, pGeomArray)
  
  ' NIGHT TPI
  Call RunDatasetAnalysis(strNightTPI, "Night TPI", 70, pMxDoc, lngCaveOrRandomIndex, pPolylineArray, _
       pPolylineValArray, pGeomArray)
  
  
  Dim pFClass As IFeatureClass
  Set pFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pGeomArray, pPolylineValArray, pPolylineFieldArray)
  Dim pFLayer As IFeatureLayer
  Set pFLayer = New FeatureLayer
  Set pFLayer.FeatureClass = pFClass
  pFLayer.Name = "Polylines"
  pFLayer.Visible = True
  pGraphMap.AddLayer pFLayer
  pMxDoc.UpdateContents
  pMxDoc.ActiveView.Refresh
  
  Debug.Print "Done..."
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pCSVFiles = Nothing
  Set pTransform2D = Nothing
  Erase strLines
  Erase strLineSplit
  Erase dblCaveVals
  Erase dblRandomVals
  Erase dblRescaleCave
  Erase dblRescaleRandom
  Erase dblGraphArray
  Set pPolylineArray = Nothing
  Set pGeomArray = Nothing
  Set pPolylineValArray = Nothing
  Set pPolylineSubArray = Nothing
  Set pPolylineFieldArray = Nothing
  Set pPolylineField = Nothing
  Set pPolylineFieldEdit = Nothing
  Set pPolyline = Nothing
  Set pGraphMap = Nothing
  Set pFClass = Nothing
  Set pFLayer = Nothing


End Sub
Public Sub RunDatasetAnalysis(strFilename As String, strName As String, dblXOffset As Double, _
    pMxDoc As IMxDocument, lngCaveOrRandomIndex As Long, pPolylineArray As esriSystem.IArray, _
    pPolylineValArray As esriSystem.IVariantArray, pGeomArray As IArray)
  
  ' DAY Slope
  Dim strText As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String
  Dim lngCount As Long
  Dim lngInterval As Long
  Dim lngCounter As Long
  Dim booFirst As Boolean
  Dim lngCaveCounter As Long
  Dim lngRandomCounter As Long
  Dim booIsFirst As Boolean
  Dim lngIndex As Long
  Dim strCaveRandom As String
  Dim dblCaveVals() As Double
  Dim dblMinVal As Double
  Dim dblMaxVal As Double
  Dim dblCaveMean As Double
  Dim dblRandomMean As Double
  Dim dblRescaleCave() As Double
  Dim dblRescaleRandom() As Double
  Dim dblGraphMax As Double
  Dim dblGraphArray() As Double
  Dim lngIndex2 As Long
  Dim dblVal As Double
  Dim dblRandomVals() As Double
  
  Dim pGraphMap As IMap
  Set pGraphMap = MyGeneralOperations.ReturnMapByName("Graphs", pMxDoc)
  
  strText = MyGeneralOperations.ReadTextFile(strFilename)
  strLines = Split(strText, vbCrLf)
  strLine = strLines(0)
  strLineSplit = Split(strLine, ",")
'  For lngIndex = 0 To UBound(strLineSplit)
'    Debug.Print "(" & CStr(lngIndex) & ") = " & strLineSplit(lngIndex)
'  Next lngIndex
  
  lngCount = UBound(strLines)
  lngInterval = lngCount / 10
  lngCounter = 0
  booFirst = True
  lngCaveCounter = -1
  lngRandomCounter = -1
  booIsFirst = True
  
  Debug.Print strName & ":"
  For lngIndex = 1 To UBound(strLines)
    lngCounter = lngCounter + 1
    If lngCounter >= lngInterval Then
      lngCounter = 0
      Debug.Print "  --> " & Format(lngIndex, "#,##0"); " of " & Format(lngCount, "#,##0") & " [" & Format(lngIndex * 100 / lngCount, "0") & "%]"
      DoEvents
    End If
    
    strLine = strLines(lngIndex)
    If strLine <> "" Then
      strLineSplit = Split(strLine, ",")
      strCaveRandom = strLineSplit(lngCaveOrRandomIndex)
      If strCaveRandom = "At_Cave" Then
        lngCaveCounter = lngCaveCounter + 1
        ReDim Preserve dblCaveVals(32, lngCaveCounter)
        
        For lngIndex2 = 7 To 39
          dblVal = CDbl(strLineSplit(lngIndex2))
          If booFirst Then
            dblMaxVal = dblVal
            dblMinVal = dblVal
            booFirst = False
          Else
            If dblVal > dblMaxVal Then dblMaxVal = dblVal
            If dblVal < dblMinVal Then dblMinVal = dblVal
          End If
          dblCaveVals(lngIndex2 - 7, lngCaveCounter) = dblVal
        Next lngIndex2
      Else
        lngRandomCounter = lngRandomCounter + 1
        ReDim Preserve dblRandomVals(32, lngRandomCounter)
        
        For lngIndex2 = 7 To 39
          dblVal = CDbl(strLineSplit(lngIndex2))
          If booFirst Then
            dblMaxVal = dblVal
            dblMinVal = dblVal
            booFirst = False
          Else
            If dblVal > dblMaxVal Then dblMaxVal = dblVal
            If dblVal < dblMinVal Then dblMinVal = dblVal
          End If
          dblRandomVals(lngIndex2 - 7, lngRandomCounter) = dblVal
        Next lngIndex2
      End If
    End If
  Next lngIndex
  
  dblCaveMean = ReturnMeanValue(dblCaveVals)
  dblRandomMean = ReturnMeanValue(dblRandomVals)
  dblRescaleCave = ShiftToMean(dblCaveVals, dblCaveMean, dblMinVal, dblMaxVal, True)
  dblRescaleRandom = ShiftToMean(dblRandomVals, dblRandomMean, dblMinVal, dblMaxVal, False)
  dblRescaleCave = RescaleTo4High(dblRescaleCave, dblMinVal, dblMaxVal)
  dblRescaleRandom = RescaleTo4High(dblRandomVals, dblMinVal, dblMaxVal)
  dblGraphArray = FillGraphArray(dblRescaleCave, dblGraphMax, pMxDoc, pPolylineArray)
  AddToPolylineArray pPolylineArray, strName, "At Cave", pPolylineValArray, pGeomArray, dblXOffset, 6
  ConvertToPointArray dblGraphArray, dblXOffset, 6, dblGraphMax, strName, "At Cave", pGraphMap
  dblGraphArray = FillGraphArray(dblRescaleRandom, dblGraphMax, pMxDoc, pPolylineArray)
  AddToPolylineArray pPolylineArray, strName, "Random", pPolylineValArray, pGeomArray, dblXOffset, 0
  ConvertToPointArray dblGraphArray, dblXOffset, 0, dblGraphMax, strName, "Random", pGraphMap
  
ClearMemory:
  Erase strLines
  Erase strLineSplit
  Erase dblCaveVals
  Erase dblRescaleCave
  Erase dblRescaleRandom
  Erase dblGraphArray
  Erase dblRandomVals
  Set pGraphMap = Nothing
End Sub





