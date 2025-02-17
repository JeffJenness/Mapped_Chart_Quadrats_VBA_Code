Attribute VB_Name = "GridFunctions"
Option Explicit

Public Function ReturnCellSize(pRaster As IRaster) As Double
  On Error GoTo erh
  Dim pRasLayer As IRasterLayer
  Set pRasLayer = New RasterLayer
  pRasLayer.CreateFromRaster pRaster

  Dim pRasterProps As IRasterProps
  Set pRasterProps = pRaster

  Dim lngNumRows As Long
  lngNumRows = pRasLayer.RowCount

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pRasterProps.Extent

  ReturnCellSize = pEnvelope.Height / lngNumRows

  Set pRasLayer = Nothing
  Set pRasterProps = Nothing
  Set pEnvelope = Nothing

  Exit Function
erh:
    MsgBox "Failed in ReturnCellSize: " & err.Description

End Function

Public Function ReturnPixelHeight(pRaster As IRaster) As Double
  Dim pRasLayer As IRasterLayer
  Set pRasLayer = New RasterLayer
  pRasLayer.CreateFromRaster pRaster

  Dim pRasterProps As IRasterProps
  Set pRasterProps = pRaster

  Dim lngNumRows As Long
  lngNumRows = pRasLayer.RowCount

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pRasterProps.Extent

  ReturnPixelHeight = pEnvelope.Height / lngNumRows

  Set pRasLayer = Nothing
  Set pRasterProps = Nothing
  Set pEnvelope = Nothing

  Exit Function

End Function

Public Function ReturnPixelWidth(pRaster As IRaster) As Double
  Dim pRasLayer As IRasterLayer
  Set pRasLayer = New RasterLayer
  pRasLayer.CreateFromRaster pRaster

  Dim pRasterProps As IRasterProps
  Set pRasterProps = pRaster

  Dim lngNumCols As Long
  lngNumCols = pRasLayer.ColumnCount

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pRasterProps.Extent

  ReturnPixelWidth = pEnvelope.Width / lngNumCols

  Set pRasLayer = Nothing
  Set pRasterProps = Nothing
  Set pEnvelope = Nothing

  Exit Function

End Function

Public Function CellValues(pPoints As IPointCollection, pRaster As IRaster) As esriSystem.IVariantArray

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSize As Double
    dblCellSize = ReturnCellSize(pRaster)

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3
    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords dWidth, dHeight

    Dim pOrigin As IPnt
    Set pOrigin = New Pnt
    pOrigin.SetCoords 0, 0

    Set pPB = pRaster.CreatePixelBlock(pPnt)
    pRaster.Read pOrigin, pPB

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim pPoint As IPoint
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim iX As Long, iY As Long

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Dim vCellValue As Variant

    For lngIndex = 0 To pPoints.PointCount - 1
      Set pPoint = pPoints.Point(lngIndex)
      If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
        pOutArray.Add Null
      Else

        dx = pPoint.x - X1
        dy = Y2 - pPoint.Y

        nX = dx / dblCellSize
        ny = dy / dblCellSize

        iX = Int(nX)
        iY = Int(ny)

        If (iX < 0) Then iX = 0
        If (iY < 0) Then iY = 0
        If (iX > pRP.Width - 1) Then
          iX = pRP.Width - 1
        End If
        If (iY > pRP.Height - 1) Then
          iY = pRP.Height - 1
        End If

        vCellValue = pPB.GetVal(0, iX, iY)
        Debug.Print "From CellValues function..." & vCellValue
        If IsEmpty(vCellValue) Then
          pOutArray.Add Null
        Else
          pOutArray.Add CDbl(vCellValue)
        End If
      End If
    Next lngIndex
    Set CellValues = pOutArray

  Set pRP = Nothing
  Set pExtent = Nothing
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing
  Set pPoint = Nothing
  Set pOutArray = Nothing

End Function

Private Function IsCellNaN(expression As Variant) As Boolean

  On Error Resume Next
  If Not IsNumeric(expression) Then
    IsCellNaN = False
    Exit Function
  End If
  If (CStr(expression) = "-1.#QNAN") Or (CStr(expression) = "1,#QNAN") Then ' can vary by locale
    IsCellNaN = True
  Else
    IsCellNaN = False
  End If

End Function

Public Function CellValue4CellInterp(pPoint As IPoint, pRaster As IRaster, _
    Optional lngBandIndex As Long = 0) As Variant

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSizeX As Double
    Dim dblCellSizeY As Double
    dblCellSizeX = ReturnPixelWidth(pRaster)
    dblCellSizeY = ReturnPixelHeight(pRaster)

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3

    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords 2, 2
    Set pPB = pRaster.CreatePixelBlock(pPnt)

    Dim pOrigin As IPnt

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = pRP.Width - 1
    lngMaxY = pRP.Height - 1

    Dim bytQuadrant As Byte       ' 1 FOR NE, 2 FOR NW, 3 FOR SW, 4 FOR SE
    Dim varInterpVal As Variant

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Dim vCellValueNE As Variant
    Dim vCellValueNW As Variant
    Dim vCellValueSE As Variant
    Dim vCellValueSW As Variant

    Dim booIsNull As Boolean
    Dim dblWestProp As Double
    Dim dblEastProp As Double

    If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
      pOutArray.Add Null
    Else

      dx = pPoint.x - X1
      dy = Y2 - pPoint.Y

      nX = dx / dblCellSizeX
      ny = dy / dblCellSizeY

      iX = Int(nX)
      iY = Int(ny)

      If (iX < 0) Then iX = 0
      If (iY < 0) Then iY = 0
      If (iX > lngMaxX) Then
        iX = lngMaxX
      End If
      If (iY > lngMaxY - 1) Then
        iY = lngMaxY - 1
      End If

      dblXRemainder = (nX - iX) * dblCellSizeX
      dblYRemainder = (ny - iY) * dblCellSizeY

      If dblYRemainder < dblHalfCellY Then                  ' ON NORTH SIDE OF CELL, SOUTH HALF OF PIXEL BLOCK
        dblPropY = (dblYRemainder + dblHalfCellY) / dblCellSizeY
        If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE OF CELL, WEST HALF OF PIXEL BLOCK
          bytQuadrant = 1                                   ' ON NORTHEAST CORNER OF CELL, SOUTHWEST CORNER OF PIXEL BLOCK
          dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
        Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
          bytQuadrant = 2                                   ' ON NORTHWEST CORNER OF CELL, SOUTHEAST CORNER OF PIXEL BLOCK
          dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
        End If
      Else                                                  ' ON SOUTH SIDE OF CELL, NORTH HALF OF PIXEL BLOCK
        dblPropY = 1 - ((dblYRemainder - dblHalfCellY) / dblCellSizeY)
        If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE, WEST HALF OF PIXEL BLOCK
          bytQuadrant = 4                                   ' ON SOUTHEAST CORNER OF CELL, NORTHWEST CORNER OF PIXEL BLOCK
          dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
        Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
          bytQuadrant = 3                                   ' ON SOUTHWEST CORNER OF CELL, NORTHEAST CORNER OF PIXEL BLOCK
          dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
        End If
      End If

      Set pOrigin = New Pnt

      booIsNull = False
      Select Case bytQuadrant
        Case 1              ' NORTHEAST                =================
          If iX = lngMaxX Or iY = 0 Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX, iY - 1
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
              dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
              varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
            End If
          End If
        Case 2              ' NORTHWEST                =================
          If iX = 0 Or iY = 0 Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX - 1, iY - 1
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
              dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
              varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
            End If
          End If
        Case 3              ' SOUTHWEST                =================
          If iX = 0 Or iY = lngMaxY Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX - 1, iY
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
              dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
              varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
            End If
          End If
        Case 4              ' SOUTHEAST                =================
          If iX = lngMaxX Or iY = lngMaxY Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX, iY
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
              dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
              varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
            End If

          End If
      End Select

      If booIsNull Then
        CellValue4CellInterp = Null
      Else
        CellValue4CellInterp = CDbl(varInterpVal)
      End If

        "   dblPropX = " & CStr(dblPropX) & vbCrLf & _
        "   dblPropY = " & CStr(dblPropY) & vbCrLf & _
        "   Quadrant = " & CStr(bytQuadrant) & vbCrLf & _
        "   vCellValueNW = " & CStr(vCellValueNW) & vbCrLf & _
        "   vCellValueNE = " & CStr(vCellValueNE) & vbCrLf & _
        "   vCellValueSW = " & CStr(vCellValueSW) & vbCrLf & _
        "   vCellValueSE = " & CStr(vCellValueSE) & vbCrLf & _
        "   dblWestProp = " & CStr(dblWestProp) & vbCrLf & _
        "   dblEastProp = " & CStr(dblEastProp) & vbCrLf & _
        "   Interpolated Value = " & CStr(varInterpVal)
    End If

ClearMemory:
  Set pRP = Nothing
  Set pExtent = Nothing
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing
  varInterpVal = Null
  Set pOutArray = Nothing
  vCellValueNE = Null
  vCellValueNW = Null
  vCellValueSE = Null
  vCellValueSW = Null

End Function


