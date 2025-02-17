Attribute VB_Name = "Map_Module"
Option Explicit

Public Function ExportActiveView(strFilename As String, Optional booAsTIFF As Boolean, _
    Optional booPDF As Boolean = False) As Boolean

  strFilename = Replace(strFilename, " / ", "_")

  Dim pApp As IApplication
  Dim pMxDoc As IMxDocument

  Set pApp = Application
  Set pMxDoc = ThisDocument

  ExportActiveView = False

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
  Dim sNameRoot As String
  Dim sOutputDir As String
  Dim iOutputResolution As Long
  Dim iScreenResolution As Long
  Dim bContinue As Boolean
  Dim msg As String
  Dim pTrackCancel As ITrackCancel
  Dim pGraphicsExtentEnv As IEnvelope
  Dim bClipToGraphicsExtent As Boolean
  Dim pUnitConvertor As IUnitConverter

  Dim pExportPNG As IExportPNG
  Dim pExportTIFF As IExportTIFF
  Dim pExportPDF3 As IExportPDF3
  Dim pExportPDF2 As IExportPDF2

  Set pActiveView = pMxDoc.ActiveView
  Set pTrackCancel = New CancelTracker

  If booPDF Then
    strFilename = aml_func_mod.SetExtension(strFilename, "pdf")
    Set pExport = New ExportPDF
    Set pExportPDF3 = pExport
    Set pExportPDF2 = pExport
    pExportPDF3.JPEGCompressionQuality = 100
    pExportPDF2.ExportPDFLayersAndFeatureAttributes = esriExportPDFLayerOptionsNone
  Else
    If booAsTIFF Then
      strFilename = aml_func_mod.SetExtension(strFilename, "tif")
      Set pExport = New ExportTIFF
      Set pExportTIFF = pExport
      pExportTIFF.CompressionType = esriTIFFCompressionLZW
      pExportTIFF.JPEGOrDeflateQuality = 100
    Else
      strFilename = aml_func_mod.SetExtension(strFilename, "png")
      Set pExport = New ExportPNG
    End If
  End If
  iOutputResolution = 600
  bClipToGraphicsExtent = False

  Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
  iPrevOutputImageQuality = pOutputRasterSettings.ResampleRatio
  If TypeOf pExport Is IExportImage Then
    SetOutputQuality pActiveView, 1
  ElseIf TypeOf pExport Is IOutputRasterSettings Then
    Set pOutputRasterSettings = pExport
    pOutputRasterSettings.ResampleRatio = 1
    Set pOutputRasterSettings = Nothing
  End If

  pExport.ExportFileName = strFilename
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
    pPixelBoundsEnv.XMin = 0
    pPixelBoundsEnv.YMin = 0
    pPixelBoundsEnv.XMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                          - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution
    pPixelBoundsEnv.YMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                          - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution

    With exportRECT
      .bottom = Fix(pPixelBoundsEnv.YMax) + 1
      .Left = Fix(pPixelBoundsEnv.XMin)
      .Top = Fix(pPixelBoundsEnv.YMin)
      .Right = Fix(pPixelBoundsEnv.XMax) + 1
    End With

    Set pMapExtEnv = pGraphicsExtentEnv
  Else
    With exportRECT
      .bottom = DisplayBounds.bottom * (iOutputResolution / iScreenResolution)
      .Left = DisplayBounds.Left * (iOutputResolution / iScreenResolution)
      .Top = DisplayBounds.Top * (iOutputResolution / iScreenResolution)
      .Right = DisplayBounds.Right * (iOutputResolution / iScreenResolution)
    End With
    pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
  End If

  pExport.PixelBounds = pPixelBoundsEnv

  Set pExport.TrackCancel = pTrackCancel
  Set pExport.StepProgressor = pApp.StatusBar.ProgressBar
  pTrackCancel.Reset
  pTrackCancel.CancelOnClick = False
  pTrackCancel.CancelOnKeyPress = True
  bContinue = pTrackCancel.Continue()

  hdc = pExport.StartExporting

  pActiveView.Output hdc, pExport.Resolution, exportRECT, pMapExtEnv, pTrackCancel

  bContinue = pTrackCancel.Continue()
  If bContinue Then
    msg = "Writing export file..."
    pApp.StatusBar.Message(0) = msg
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
    pApp.StatusBar.Message(0) = msg
  End If

  SetOutputQuality pActiveView, iPrevOutputImageQuality
  Set pTrackCancel = Nothing
  Set pMapExtEnv = Nothing
  Set pPixelBoundsEnv = Nothing

  ExportActiveView = True

End Function

Private Sub SetOutputQuality(pActiveView As IActiveView, iResampleRatio As Long)

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

    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = iResampleRatio

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

Function GetGraphicsExtent(pActiveView As IActiveView) As IEnvelope

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

  Set GetGraphicsExtent = pBounds

  Set pBounds = Nothing
  Set pEnv = Nothing
  Set pGraphicsContainer = Nothing
  Set pPageLayout = Nothing
  Set pDisplay = Nothing
  Set pElement = Nothing

End Function


