Attribute VB_Name = "Map_Module_Export_GeoTIFFs"
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
           "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
           ByVal lpFile As String, ByVal lpParameters As String, _
           ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
           
Public Sub TestExport()

  Dim strNameRoot As String
  Dim strOutputDir As String
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  strOutputDir = "D:\arcGIS_stuff\consultation\Jut_Wynne\Image_Files\"
  strNameRoot = aml_func_mod.ReturnFilename2(MyGeneralOperations.MakeUniquedBASEName(strOutputDir & "test"))
  
'  ExportActiveView2 strNameRoot, strOutputDir

  Dim booRestart As Boolean
  Dim lngXIntervals As Long
  Dim lngYIntervals As Long

  Dim pTopoLayer As ILayer
  Dim pImageLayer As ILayer
  Dim pNAIP2012 As ILayer
  Dim pNAIP2014 As ILayer
  
  Set pTopoLayer = MyGeneralOperations.ReturnLayerByName("USA_Topo_Maps", pMxDoc.FocusMap)
  Set pImageLayer = MyGeneralOperations.ReturnLayerByName("World_Imagery", pMxDoc.FocusMap)
  Set pNAIP2012 = MyGeneralOperations.ReturnLayerByName("Base_Remote_Sensing\NAIP_2012_4Band", pMxDoc.FocusMap)
  Set pNAIP2014 = MyGeneralOperations.ReturnLayerByName("Base_Remote_Sensing\NAIP_2014_County_Mosaics", pMxDoc.FocusMap)
  
  Dim pFLayer As IFeatureLayer
  Dim pGeoDataset As IGeoDataset
  Dim pPolygon As IPolygon
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pEnvPoly As IPolygon
  
  Dim pAOIBoundary As IPolygon
  Dim pRelOp As IRelationalOperator

  Set pFLayer = MyGeneralOperations.ReturnLayerByName("AOI", pMxDoc.FocusMap)
  Set pGeoDataset = pFLayer.FeatureClass
  Set pFClass = pFLayer.FeatureClass
  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Set pAOIBoundary = pFeature.ShapeCopy
  Set pRelOp = pAOIBoundary
  
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Static dblYIndex As Double
  Static dblXIndex As Double
  Static lngRow As Long
  Static lngCol As Long
  Dim pExtent As IEnvelope
  Dim pNewExtent As IEnvelope

  Dim dblHeight As Double
  Dim dblWidth As Double
  Dim dblXInterval As Double
  Dim dblYInterval As Double
  Dim lngTotalCounter As Long
  Dim lngTotalCount As Long
  Dim lngCountAll As Long
  Dim lngCountIntersect As Long
      
  Set pExtent = pGeoDataset.Extent
  dblHeight = pExtent.Height
  dblWidth = pExtent.Width
  
  ' TOPOS
  pTopoLayer.Visible = True
  pImageLayer.Visible = False
  pNAIP2012.Visible = False
  pNAIP2014.Visible = False
  pMxDoc.ActiveView.Refresh
  
  lngXIntervals = 20
  lngYIntervals = 10
  booRestart = False

  dblXInterval = dblWidth / lngXIntervals
  dblYInterval = dblHeight / lngYIntervals

  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  
  lngTotalCount = lngXIntervals * lngYIntervals
  lngTotalCounter = 0
  
  
'  For lngRow = lngYIntervals To 1 Step -1
'    For lngCol = 1 To lngXIntervals
'      lngTotalCounter = lngTotalCounter + 1
'      Debug.Print "Testing " & CStr(lngTotalCounter) & " of " & CStr(lngTotalCount) & " [Row " & CStr(lngRow) & _
'          " of " & CStr(lngXIntervals) & ", Column " & CStr(lngCol) & " of " & CStr(lngYIntervals) & "]..."
'
'      dblXIndex = pExtent.XMin + (CDbl(lngCol - 1) * dblXInterval)
'      dblYIndex = pExtent.YMin + (CDbl(lngRow - 1) * dblYInterval)
'      Set pNewExtent = New Envelope
'      Set pNewExtent.SpatialReference = pExtent.SpatialReference
'      pNewExtent.PutCoords dblXIndex, dblYIndex, dblXIndex + dblXInterval, dblYIndex + dblYInterval
'      pNewExtent.Expand 1.05, 1.05, True
'      Set pEnvPoly = MyGeometricOperations.EnvelopeToPolygon(pNewExtent)
'      If Not pRelOp.Disjoint(pEnvPoly) Then
''        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pEnvPoly, "Delete_Me"
'
'        strNameRoot = aml_func_mod.ReturnFilename2(MyGeneralOperations.MakeUniquedBASEName(strOutputDir & _
'            "USA_Topo_Maps\Topo_Row_" & Format(lngRow, "000") & "_Col_" & Format(lngCol, "000")))
'        pMxDoc.ActiveView.Extent = pNewExtent
'        pMxDoc.ActiveView.Refresh
'        ExportActiveView2 strNameRoot, strOutputDir & "USA_Topo_Maps\"
'      End If
'
'    Next lngCol
'  Next lngRow

'  ' IMAGERY
'  pTopoLayer.Visible = False
'  pImageLayer.Visible = True
'  pNAIP2012.Visible = False
'  pNAIP2014.Visible = False
'  pMxDoc.ActiveView.Refresh
'
'  lngXIntervals = 20
'  lngYIntervals = 10
'  booRestart = False
'
'  dblXInterval = dblWidth / lngXIntervals
'  dblYInterval = dblHeight / lngYIntervals
'
'  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
'
'  lngTotalCount = lngXIntervals * lngYIntervals
'  lngTotalCounter = 0
'
'  For lngRow = lngYIntervals To 1 Step -1
'    For lngCol = 1 To lngXIntervals
'      lngTotalCounter = lngTotalCounter + 1
'      Debug.Print "Testing " & CStr(lngTotalCounter) & " of " & CStr(lngTotalCount) & " [Row " & CStr(lngRow) & _
'          " of " & CStr(lngXIntervals) & ", Column " & CStr(lngCol) & " of " & CStr(lngYIntervals) & "]..."
'
'      dblXIndex = pExtent.XMin + (CDbl(lngCol - 1) * dblXInterval)
'      dblYIndex = pExtent.YMin + (CDbl(lngRow - 1) * dblYInterval)
'      Set pNewExtent = New Envelope
'      Set pNewExtent.SpatialReference = pExtent.SpatialReference
'      pNewExtent.PutCoords dblXIndex, dblYIndex, dblXIndex + dblXInterval, dblYIndex + dblYInterval
'      pNewExtent.Expand 1.05, 1.05, True
'      Set pEnvPoly = MyGeometricOperations.EnvelopeToPolygon(pNewExtent)
'
'      lngCountAll = lngCountAll + 1
'      If Not pRelOp.Disjoint(pEnvPoly) Then
'        lngCountIntersect = lngCountIntersect + 1
''        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pEnvPoly, "Delete_Me"
'
'        strNameRoot = aml_func_mod.ReturnFilename2(MyGeneralOperations.MakeUniquedBASEName(strOutputDir & _
'            "World_Imagery\Imagery_Row_" & Format(lngRow, "000") & "_Col_" & Format(lngCol, "000")))
'        pMxDoc.ActiveView.Extent = pNewExtent
'        pMxDoc.ActiveView.Refresh
'        ExportActiveView2 strNameRoot, strOutputDir & "World_Imagery\"
'      End If
'
'    Next lngCol
'  Next lngRow
'
'  Debug.Print lngCountIntersect
'  Debug.Print lngCountAll
'  DoEvents
'  ' NAIP_2012
'  pTopoLayer.Visible = False
'  pImageLayer.Visible = False
'  pNAIP2012.Visible = True
'  pNAIP2014.Visible = False
'  pMxDoc.ActiveView.Refresh
'
'  lngXIntervals = 20
'  lngYIntervals = 10
'  booRestart = False
'
'  dblXInterval = dblWidth / lngXIntervals
'  dblYInterval = dblHeight / lngYIntervals
'
'  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
'
'  lngTotalCount = lngXIntervals * lngYIntervals
'  lngTotalCounter = 0
'
'  For lngRow = lngYIntervals To 1 Step -1
'    For lngCol = 1 To lngXIntervals
'      lngTotalCounter = lngTotalCounter + 1
'      Debug.Print "Testing " & CStr(lngTotalCounter) & " of " & CStr(lngTotalCount) & " [Row " & CStr(lngRow) & _
'          " of " & CStr(lngXIntervals) & ", Column " & CStr(lngCol) & " of " & CStr(lngYIntervals) & "]..."
'
'      dblXIndex = pExtent.XMin + (CDbl(lngCol - 1) * dblXInterval)
'      dblYIndex = pExtent.YMin + (CDbl(lngRow - 1) * dblYInterval)
'      Set pNewExtent = New Envelope
'      Set pNewExtent.SpatialReference = pExtent.SpatialReference
'      pNewExtent.PutCoords dblXIndex, dblYIndex, dblXIndex + dblXInterval, dblYIndex + dblYInterval
'      pNewExtent.Expand 1.05, 1.05, True
'      Set pEnvPoly = MyGeometricOperations.EnvelopeToPolygon(pNewExtent)
'      If Not pRelOp.Disjoint(pEnvPoly) Then
''        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pEnvPoly, "Delete_Me"
'
'        strNameRoot = aml_func_mod.ReturnFilename2(MyGeneralOperations.MakeUniquedBASEName(strOutputDir & _
'            "NAIP_2012_Imagery\Imagery_Row_" & Format(lngRow, "000") & "_Col_" & Format(lngCol, "000")))
'        pMxDoc.ActiveView.Extent = pNewExtent
'        pMxDoc.ActiveView.Refresh
'        ExportActiveView2 strNameRoot, strOutputDir & "NAIP_2012_Imagery\"
'      End If
'
'    Next lngCol
'  Next lngRow


  ' NAIP_2014
  pTopoLayer.Visible = False
  pImageLayer.Visible = False
  pNAIP2012.Visible = False
  pNAIP2014.Visible = True
  pMxDoc.ActiveView.Refresh
  
  lngXIntervals = 20
  lngYIntervals = 10
  booRestart = False

  dblXInterval = dblWidth / lngXIntervals
  dblYInterval = dblHeight / lngYIntervals

  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  
  lngTotalCount = lngXIntervals * lngYIntervals
  lngTotalCounter = 0
    
  For lngRow = lngYIntervals To 1 Step -1
    For lngCol = 1 To lngXIntervals
      lngTotalCounter = lngTotalCounter + 1
      Debug.Print "Testing " & CStr(lngTotalCounter) & " of " & CStr(lngTotalCount) & " [Row " & CStr(lngRow) & _
          " of " & CStr(lngYIntervals) & ", Column " & CStr(lngCol) & " of " & CStr(lngXIntervals) & "]..."
      DoEvents
      dblXIndex = pExtent.XMin + (CDbl(lngCol - 1) * dblXInterval)
      dblYIndex = pExtent.YMin + (CDbl(lngRow - 1) * dblYInterval)
      Set pNewExtent = New Envelope
      Set pNewExtent.SpatialReference = pExtent.SpatialReference
      pNewExtent.PutCoords dblXIndex, dblYIndex, dblXIndex + dblXInterval, dblYIndex + dblYInterval
      pNewExtent.Expand 1.05, 1.05, True
      Set pEnvPoly = MyGeometricOperations.EnvelopeToPolygon(pNewExtent)
      If Not pRelOp.Disjoint(pEnvPoly) Then
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pEnvPoly, "Delete_Me"

        strNameRoot = aml_func_mod.ReturnFilename2(MyGeneralOperations.MakeUniquedBASEName(strOutputDir & _
            "NAIP_2014_Imagery\Imagery_Row_" & Format(lngRow, "000") & "_Col_" & Format(lngCol, "000")))
        pMxDoc.ActiveView.Extent = pNewExtent
        pMxDoc.ActiveView.Refresh
        ExportActiveView2 strNameRoot, strOutputDir & "NAIP_2014_Imagery\"
      End If

    Next lngCol
  Next lngRow
  
  Debug.Print "Done..."
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFLayer = Nothing
  Set pGeoDataset = Nothing
  Set pPolygon = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pAOIBoundary = Nothing
  Set pRelOp = Nothing
  Set pExtent = Nothing
  Set pNewExtent = Nothing
  Set pEnvPoly = Nothing


End Sub

Public Sub ExportActiveView2(strNameRoot As String, strOutputDir As String)
    
    Dim pMxDoc As IMxDocument
    Dim pActiveView As IActiveView
    Dim pExport As IExport
    Dim iPrevOutputImageQuality As Long
    Dim pOutputRasterSettings As IOutputRasterSettings
    Dim pPixelBoundsEnv As IEnvelope
    Dim exportRECT As tagRECT
    Dim DisplayBounds As tagRECT
    Dim pDisplayTransformation As IDisplayTransformation
    Dim pPageLayout As IPageLayout
    Dim pMapExtEnv As IEnvelope
    Dim hdc As Long
    Dim tmpDC As Long
    Dim iOutputResolution As Long
    Dim iScreenResolution As Long
    Dim bContinue As Boolean
    Dim msg As String
    Dim pTrackCancel As ITrackCancel
    Dim pGraphicsExtentEnv As IEnvelope
    Dim bClipToGraphicsExtent As Boolean
    Dim pUnitConvertor As IUnitConverter
    
    Set pMxDoc = Application.Document
    Set pActiveView = pMxDoc.ActiveView
    Set pTrackCancel = New CancelTracker
    
    'Create an ExportPDF object and QI the pExport interface pointer onto it.
    ' To export to a format other than PDF, simply create a different CoClass here
    Set pExport = New ExportTIFF
    Dim pExportTiff As IExportTIFF
    Dim pWorldFileSettings As IWorldFileSettings
    Set pExportTiff = pExport
    Set pWorldFileSettings = pExport
    If TypeOf pActiveView Is IMap Then
      pExportTiff.GeoTiff = True
      pWorldFileSettings.MapExtent = pActiveView.Extent
    End If
    Set pExportTiff = Nothing
    Set pWorldFileSettings = Nothing
    'assign a resolution for the export in dpi
    iOutputResolution = 200
    'assign True or False to determin is export image will be clipped to the graphic extent of layout elements.
    'this value is ignored for data view exports
    bClipToGraphicsExtent = True
    
    
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    iPrevOutputImageQuality = pOutputRasterSettings.ResampleRatio
    ' Output Image Quality of the export.  The value here will only be used if the export
    '  object is a format that allows setting of Output Image Quality, i.e. a vector exporter.
    '  The value assigned to ResampleRatio should be in the range 1 to 5.
    '  1 corresponds to "Best", 5 corresponds to "Fast"
    If TypeOf pExport Is IExportImage Then
        'always set the output quality of the display to 1 for image export formats
        SetOutputQuality2 pActiveView, 1
    ElseIf TypeOf pExport Is IOutputRasterSettings Then
        ' for vector formats, assign a ResampleRatio to control drawing of raster layers at export time
        Set pOutputRasterSettings = pExport
        pOutputRasterSettings.ResampleRatio = 1
        Set pOutputRasterSettings = Nothing
    End If
    
    'assign the output path and filename.  We can use the Filter property of the export object to
    ' automatically assign the proper extension to the file.
'    strOutputDir = "D:\arcGIS_stuff\consultation\Springs_Stewardship_Institute\Grand_Canyon\Temp_Images\"
'    strNameRoot = Left(ThisDocument.Title, Len(ThisDocument.Title) - 4)

    If Right(strOutputDir, 1) <> "\" Then strOutputDir = strOutputDir & "\"
    
    pExport.ExportFileName = strOutputDir & strNameRoot & "." & Right(Split(pExport.Filter, "|")(1), _
                             Len(Split(pExport.Filter, "|")(1)) - 2)
    tmpDC = GetDC(0)
    iScreenResolution = GetDeviceCaps(tmpDC, 88) '88 is the win32 const for Logical pixels/inch in X)
    ReleaseDC 0, tmpDC
    pExport.Resolution = iOutputResolution
    
    If TypeOf pActiveView Is IPageLayout Then
        DisplayBounds = pActiveView.ExportFrame
        Set pMapExtEnv = pGraphicsExtentEnv
    Else
        Set pDisplayTransformation = pActiveView.ScreenDisplay.DisplayTransformation
        DisplayBounds.Left = 0
        DisplayBounds.Top = 0
        DisplayBounds.Right = pDisplayTransformation.DeviceFrame.Right
        DisplayBounds.bottom = pDisplayTransformation.DeviceFrame.bottom
        Set pMapExtEnv = New Envelope
        Set pMapExtEnv = pDisplayTransformation.FittedBounds
    End If
    
    Set pPixelBoundsEnv = New Envelope
    If bClipToGraphicsExtent And (TypeOf pActiveView Is IPageLayout) Then
        Set pGraphicsExtentEnv = GetGraphicsExtent(pActiveView)
        Set pPageLayout = pActiveView
        Set pUnitConvertor = New UnitConverter
        'assign the x and y values representing the clipped area to the PixelBounds envelope
        pPixelBoundsEnv.XMin = 0
        pPixelBoundsEnv.YMin = 0
        pPixelBoundsEnv.XMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                               - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution
        pPixelBoundsEnv.YMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                               - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution
        
        'assign the x and y values representing the clipped export extent to the exportRECT
        With exportRECT
            .bottom = Fix(pPixelBoundsEnv.YMax) + 1
            .Left = Fix(pPixelBoundsEnv.XMin)
            .Top = Fix(pPixelBoundsEnv.YMin)
            .Right = Fix(pPixelBoundsEnv.XMax) + 1
        End With
        
        Set pMapExtEnv = pGraphicsExtentEnv
    Else
        'The values in the exportRECT tagRECT correspond to the width
        ' and height to export, measured in pixels with an origin in the top left corner.
        With exportRECT
            .bottom = DisplayBounds.bottom * (iOutputResolution / iScreenResolution)
            .Left = DisplayBounds.Left * (iOutputResolution / iScreenResolution)
            .Top = DisplayBounds.Top * (iOutputResolution / iScreenResolution)
            .Right = DisplayBounds.Right * (iOutputResolution / iScreenResolution)
        End With
        'populate the PixelBounds envelope with the values from exportRECT.
        ' We need to do this because the exporter object requires an envelope object
        ' instead of a tagRECT structure.
        pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
    End If
    
    'Assign the envelope object to the exporter object's PixelBounds property.  The exporter object
    ' will use these dimensions when allocating memory for the export file.
    pExport.PixelBounds = pPixelBoundsEnv
    
    Set pExport.TrackCancel = pTrackCancel
    Set pExport.StepProgressor = Application.StatusBar.ProgressBar
    pTrackCancel.Reset
    pTrackCancel.CancelOnClick = False
    pTrackCancel.CancelOnKeyPress = True
    bContinue = pTrackCancel.Continue()
    
    hdc = pExport.StartExporting
    
    'Redraw the active view, rendering it to the exporter object device context instead of the app display.
    'We pass the following values:
    ' * hDC is the device context of the exporter object.
    ' * exportRECT is the tagRECT structure that describes the dimensions of the view that will be rendered.
    ' The values in exportRECT should match those held in the exporter object's PixelBounds property.
    ' * pMapExtEnv is an envelope defining the section of the original image to draw into the export object.
    ' * pTrackCancel is a reference to a CancelTracker object
    pActiveView.Output hdc, pExport.Resolution, exportRECT, pMapExtEnv, pTrackCancel
    
    bContinue = pTrackCancel.Continue()
    If bContinue Then
        msg = "Writing export file..."
        Application.StatusBar.Message(0) = msg
        pExport.FinishExporting
        pExport.Cleanup
    Else
        pExport.Cleanup
    End If
    pTrackCancel.CancelOnClick = False
    pTrackCancel.CancelOnKeyPress = True
    
    bContinue = pTrackCancel.Continue()
    If bContinue Then
        msg = "Finished exporting '" & pExport.ExportFileName & "'"
        Application.StatusBar.Message(0) = msg
    End If
    
    SetOutputQuality2 pActiveView, iPrevOutputImageQuality
    Set pTrackCancel = Nothing
    Set pMapExtEnv = Nothing
    Set pPixelBoundsEnv = Nothing
End Sub


Private Sub SetOutputQuality2(pActiveView As IActiveView, iResampleRatio As Long)
    Dim pMap As IMap
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pElement As IElement
    Dim pOutputRasterSettings As IOutputRasterSettings
    Dim pMapFrame As IMapFrame
    Dim pTmpActiveView As IActiveView
    
    If TypeOf pActiveView Is IMap Then
        Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
        pOutputRasterSettings.ResampleRatio = iResampleRatio
    ElseIf TypeOf pActiveView Is IPageLayout Then
        
        'assign ResampleRatio for PageLayout
        Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
        pOutputRasterSettings.ResampleRatio = iResampleRatio
        
        'and assign ResampleRatio to the Maps in the PageLayout
        Set pGraphicsContainer = pActiveView
        pGraphicsContainer.Reset
        Set pElement = pGraphicsContainer.Next
        Do While Not pElement Is Nothing
            If TypeOf pElement Is IMapFrame Then
                Set pMapFrame = pElement
                Set pTmpActiveView = pMapFrame.Map
                Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
                pOutputRasterSettings.ResampleRatio = iResampleRatio
            End If
            DoEvents
            Set pElement = pGraphicsContainer.Next
        Loop
        Set pMap = Nothing
        Set pMapFrame = Nothing
        Set pGraphicsContainer = Nothing
        Set pTmpActiveView = Nothing
    End If
    Set pOutputRasterSettings = Nothing
    
End Sub

Function GetGraphicsExtent2(pActiveView As IActiveView) As IEnvelope
    Dim pBounds As IEnvelope
    Dim pEnv As IEnvelope
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pPageLayout As IPageLayout
    Dim pDisplay As IDisplay
    Dim pElement As IElement
    
    Set pBounds = New Envelope
    Set pEnv = New Envelope
    Set pPageLayout = pActiveView
    Set pDisplay = pActiveView.ScreenDisplay
    Set pGraphicsContainer = pActiveView
    pGraphicsContainer.Reset
    
    Set pElement = pGraphicsContainer.Next
    Do While Not pElement Is Nothing
        pElement.QueryBounds pDisplay, pEnv
        pBounds.Union pEnv
        DoEvents
        Set pElement = pGraphicsContainer.Next
    Loop
    
    Set GetGraphicsExtent2 = pBounds
    
    Set pBounds = Nothing
    Set pEnv = Nothing
    Set pGraphicsContainer = Nothing
    Set pPageLayout = Nothing
    Set pDisplay = Nothing
    Set pElement = Nothing
    
End Function

