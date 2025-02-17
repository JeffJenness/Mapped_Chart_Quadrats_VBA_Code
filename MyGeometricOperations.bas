Attribute VB_Name = "MyGeometricOperations"
Option Explicit

Public Function DegToRad(dblDeg As Double) As Double

  DegToRad = dblDeg * dblPI / 180

End Function

Public Function EnvelopeToPolygon(pEnv As IEnvelope) As IPolygon

  Dim pPtColl As IPointCollection

  Set pPtColl = New Polygon
  With pPtColl
      .AddPoint pEnv.LowerLeft
      .AddPoint pEnv.UpperLeft
      .AddPoint pEnv.UpperRight
      .AddPoint pEnv.LowerRight
      .AddPoint pEnv.LowerLeft
  End With

  Dim pPolygon As IPolygon
  Set pPolygon = pPtColl
  Set pPolygon.SpatialReference = pEnv.SpatialReference
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pPolygon
  pTopoOp.Simplify

  Set EnvelopeToPolygon = pPtColl

End Function

Public Sub Graphic_MakeFromGeometry(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String)

  Dim pMxDocument As esriArcMapUI.IMxDocument
  Dim pActiveView As esriCarto.IActiveView

  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMxDoc.FocusMap

  Dim pElement As IElement
  Dim pPolygonElement As IPolygonElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties

  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone

  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType

  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
    Case 4, 11:
      Set pElement = New PolygonElement
    Case 5:
      Set pElement = New RectangleElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select

  pElement.Geometry = pNewGeometry
  Set pGraphicElement = pElement
  Set pSpatialReference = pGeometry.SpatialReference
  Set pGraphicElement.SpatialReference = pSpatialReference
  Set pElementProperties = pElement
  pElementProperties.Name = strName

  pGContainer.AddElement pElement, 0

  pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

End Sub

Public Function PointAdd(pPointA As IPoint, pPointB As IPoint) As IPoint

  Set PointAdd = New Point
  PointAdd.PutCoords pPointA.x + pPointB.x, pPointA.Y + pPointB.Y

End Function

Public Function PointSubtract(pPointA As IPoint, pPointB As IPoint) As IPoint

  Set PointSubtract = New Point
  PointSubtract.PutCoords pPointA.x - pPointB.x, pPointA.Y - pPointB.Y

End Function

Public Function AsRadians(theDegrees As Double) As Double

  AsRadians = dblPI * (theDegrees / 180)

End Function

Public Sub CalcPointLine(ptOrigin As IPoint, theLength As Double, dblAzimuth As Double, ptEndPoint As IPoint, _
    Optional pLine As IPolyline)

Dim theAzimuth As Double
theAzimuth = dblAzimuth

Set ptEndPoint = New Point

Do While theAzimuth < 0
  theAzimuth = theAzimuth + 360
Loop
Do While theAzimuth > 360
  theAzimuth = theAzimuth - 360
Loop
Dim NorthSouthDistance As Double
Dim EastWestDistance As Double
Dim EastWest As Integer
Dim NorthSouth As Integer

If theAzimuth = 0 Or theAzimuth = 360 Then
  NorthSouthDistance = theLength
  NorthSouth = 1
  EastWestDistance = 0
  EastWest = 1
ElseIf (theAzimuth = 180) Then
  NorthSouthDistance = theLength
  NorthSouth = -1
  EastWestDistance = 0
  EastWest = 1
ElseIf (theAzimuth = 90) Then
  NorthSouthDistance = 0
  NorthSouth = 1
  EastWestDistance = theLength
  EastWest = 1
ElseIf (theAzimuth = 270) Then
  NorthSouthDistance = 0
  NorthSouth = 1
  EastWestDistance = theLength
  EastWest = -1
ElseIf ((theAzimuth > 0) And (theAzimuth < 90)) Then
  NorthSouthDistance = Cos(AsRadians(theAzimuth)) * theLength
  NorthSouth = 1
  EastWestDistance = Sin(AsRadians(theAzimuth)) * theLength
  EastWest = 1
ElseIf ((theAzimuth > 90) And (theAzimuth < 180)) Then
  NorthSouthDistance = (Sin(AsRadians(theAzimuth - 90))) * theLength
  NorthSouth = -1
  EastWestDistance = (Cos(AsRadians(theAzimuth - 90))) * theLength
  EastWest = 1
ElseIf ((theAzimuth > 180) And (theAzimuth < 270)) Then
  NorthSouthDistance = (Cos(AsRadians(theAzimuth - 180))) * theLength
  NorthSouth = -1
  EastWestDistance = (Sin(AsRadians(theAzimuth - 180))) * theLength
  EastWest = -1
ElseIf ((theAzimuth > 270) And (theAzimuth < 360)) Then
  NorthSouthDistance = (Sin(AsRadians(theAzimuth - 270))) * theLength
  NorthSouth = 1
  EastWestDistance = (Cos(AsRadians(theAzimuth - 270))) * theLength
  EastWest = -1
End If

Dim theMovementNorth As Double
Dim theMovementWest As Double

theMovementNorth = NorthSouthDistance * NorthSouth
theMovementWest = EastWestDistance * EastWest

Dim startX As Double
Dim startY As Double

ptOrigin.QueryCoords startX, startY
ptEndPoint.PutCoords startX + theMovementWest, startY + theMovementNorth

Set ptEndPoint.SpatialReference = ptOrigin.SpatialReference

If Not pLine Is Nothing Then
  Dim pPointColl As IPointCollection
  pLine.SetEmpty
  Set pPointColl = pLine
  pPointColl.AddPoint ptOrigin
  pPointColl.AddPoint ptEndPoint
  Set pLine.SpatialReference = ptOrigin.SpatialReference
End If

End Sub

Public Function CreateCircleAroundPoint(pOrigin As IPoint, dblRadius As Double, lngPtCount As Long)

  Dim dblInterval As Double
  dblInterval = 360 / lngPtCount
  Dim dblIndex As Double

  Dim pCircle As IPointCollection
  Set pCircle = New Polygon
  Dim pGeom As IGeometry
  Set pGeom = pCircle
  Set pGeom.SpatialReference = pOrigin.SpatialReference

  Dim pNewPoint As IPoint
  Dim pTopoOp As ITopologicalOperator

  dblIndex = 0
  Do Until dblIndex >= 360

    CalcPointLine pOrigin, dblRadius, dblIndex, pNewPoint
    pCircle.AddPoint pNewPoint

    dblIndex = dblIndex + dblInterval
  Loop

  Dim pFinal As IPolygon
  Set pFinal = pCircle
  pFinal.Close
  Set pTopoOp = pFinal
  pTopoOp.Simplify

  Set CreateCircleAroundPoint = pFinal

End Function

Public Function SquaredDistanceBetweenSegments( _
    dblSeg1Start() As Double, _
    dblSeg1End() As Double, _
    dblSeg2Start() As Double, _
    dblSeg2End() As Double, _
    dblClosePointOnSeg1() As Double, _
    dblClosePointOnSeg2() As Double) As Double

  Dim dblVectorU() As Double    ' VECTOR OF (SEGMENT 1 END POINT) - (SEGMENT 1 START POINT)
  Dim dblVectorV() As Double    ' VECTOR OF (SEGMENT 2 END POINT) - (SEGMENT 2 START POINT)
  Dim dblVectorW() As Double    ' VECTOR OF (SEGMENT 1 START POINT) - (SEGMENT 2 START POINT)

  Dim dblA As Double     ' DOT PRODUCT OF (VectorU * VectorU)
  Dim dblB As Double     ' DOT PRODUCT OF (VectorU * VectorV)
  Dim dblC As Double     ' DOT PRODUCT OF (VectorV * VectorV)
  Dim dblD As Double     ' DOT PRODUCT OF (VectorU * VectorW)
  Dim dblE As Double     ' DOT PRODUCT OF (VectorV * VectorW)
  Dim lngIndex As Long
  Dim lngUpper As Long
  Dim dblDenominator As Double
  Dim dblsc As Double
  Dim dblsN As Double
  Dim dblSD As Double
  Dim dblTC As Double
  Dim dbltN As Double
  Dim dbltD As Double

  Dim dblSmallNum As Double
  dblSmallNum = 0.000000000001

  lngUpper = UBound(dblSeg1Start)
  ReDim dblVectorU(lngUpper)
  ReDim dblVectorV(lngUpper)
  ReDim dblVectorW(lngUpper)
  ReDim dblClosePointOnSeg1(lngUpper)
  ReDim dblClosePointOnSeg2(lngUpper)

  dblA = 0
  dblB = 0
  dblC = 0
  dblD = 0
  dblE = 0

  For lngIndex = 0 To lngUpper
    dblVectorU(lngIndex) = (dblSeg1End(lngIndex) - dblSeg1Start(lngIndex))
    dblVectorV(lngIndex) = (dblSeg2End(lngIndex) - dblSeg2Start(lngIndex))
    dblVectorW(lngIndex) = (dblSeg1Start(lngIndex) - dblSeg2Start(lngIndex))
  Next lngIndex

  For lngIndex = 0 To lngUpper
    dblA = dblA + (dblVectorU(lngIndex) * dblVectorU(lngIndex))
    dblB = dblB + (dblVectorU(lngIndex) * dblVectorV(lngIndex))
    dblC = dblC + (dblVectorV(lngIndex) * dblVectorV(lngIndex))
    dblD = dblD + (dblVectorU(lngIndex) * dblVectorW(lngIndex))
    dblE = dblE + (dblVectorV(lngIndex) * dblVectorW(lngIndex))
  Next lngIndex

  dblDenominator = (dblA * dblC) - (dblB * dblB)
  dblsc = dblDenominator
  dblsN = dblDenominator
  dblSD = dblDenominator
  dblTC = dblDenominator
  dbltN = dblDenominator
  dbltD = dblDenominator

  If dblDenominator < dblSmallNum Then
    dblsN = 0
    dblSD = 1
    dbltN = dblE
    dbltD = dblC

  Else
    dblsN = (dblB * dblE) - (dblC * dblD)
    dbltN = (dblA * dblE) - (dblB * dblD)

    If dblsN < 0 Then
      dblsN = 0
      dbltN = dblE
      dbltD = dblC

    ElseIf dblsN > dblSD Then
      dblsN = dblSD
      dbltN = dblE + dblB
      dbltD = dblC
    End If
  End If

  If dbltN < 0 Then
    dbltN = 0

    If -dblD < 0 Then
      dblsN = 0

    ElseIf -dblD > dblA Then
      dblsN = dblSD

    Else
      dblsN = -dblD
      dblSD = dblA

    End If

  ElseIf dbltN > dbltD Then
    dbltN = dbltD

    If ((-dblD + dblB) < 0) Then
      dblsN = 0

    ElseIf ((-dblD + dblB) > dblA) Then
      dblsN = dblSD

    Else
      dblsN = -dblD + dblB
      dblSD = dblA

    End If
  End If

  If Abs(dblsN) < dblSmallNum Then
    dblsc = 0
  Else
    dblsc = dblsN / dblSD
  End If

  If Abs(dbltN) < dblSmallNum Then
    dblTC = 0
  Else
    dblTC = dbltN / dbltD
  End If

  Dim dblP() As Double
  ReDim dblP(lngUpper)
  Dim dblDistance As Double
  dblDistance = 0
  For lngIndex = 0 To lngUpper
          (dbltc * (dblVectorV(lngIndex))))
    dblClosePointOnSeg1(lngIndex) = dblSeg1Start(lngIndex) + dblsc * (dblVectorU(lngIndex))
    dblClosePointOnSeg2(lngIndex) = dblSeg2Start(lngIndex) + dblTC * (dblVectorV(lngIndex))
    dblDistance = dblDistance + ((dblClosePointOnSeg1(lngIndex) - dblClosePointOnSeg2(lngIndex)) ^ 2)
  Next lngIndex

  SquaredDistanceBetweenSegments = dblDistance

End Function

Public Function UnionGeometries3(pGeomArray As esriSystem.IVariantArray, _
    Optional lngMaxNumberToUnion As Long = -999) As IGeometry

  Dim pTopoOp As ITopologicalOperator
  Dim pGeom As IGeometry
  Dim pGeometryCollection As IGeometryCollection

  Set pGeometryCollection = New GeometryBag

  Dim pSpRef As ISpatialReference
  Dim pTempGeom As IGeometry
  Dim pNewGeom As IGeometry
  Dim lngIndex As Long
  Dim booFoundGeometry As Boolean

  Do Until lngIndex = pGeomArray.Count Or Not pSpRef Is Nothing
    Set pGeom = pGeomArray.Element(0)
    If Not pGeom Is Nothing Then
      Set pSpRef = pGeom.SpatialReference
      booFoundGeometry = True
    End If
    lngIndex = lngIndex + 1
  Loop

  Dim lngGeomType As esriGeometryType
  lngGeomType = pGeom.GeometryType

  If Not booFoundGeometry Then
    Set UnionGeometries3 = Nothing
  Else
    For lngIndex = 0 To pGeomArray.Count - 1
      Set pGeom = pGeomArray.Element(lngIndex)

      If Not pGeom Is Nothing Then
        If Not pGeom.IsEmpty Then
          pGeometryCollection.AddGeometry pGeom

          If lngMaxNumberToUnion > 1 Then
            If pGeometryCollection.GeometryCount >= lngMaxNumberToUnion Then

              If lngGeomType = esriGeometryPoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryMultipoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryPolyline Then
                Set pTempGeom = New Polyline
              ElseIf lngGeomType = esriGeometryPolygon Then
                Set pTempGeom = New Polygon
              End If

              Set pTopoOp = pTempGeom
              pTopoOp.ConstructUnion pGeometryCollection
              pTopoOp.Simplify

              Set pTempGeom.SpatialReference = pSpRef
              Set pGeometryCollection = New GeometryBag
              pGeometryCollection.AddGeometry pTempGeom

            End If
          End If
        End If
      End If

    Next lngIndex

    If pGeometryCollection.GeometryCount = 1 Then
      Set pNewGeom = pGeometryCollection.Geometry(0)
    Else
      If lngGeomType = esriGeometryPoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryMultipoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryPolyline Then
        Set pNewGeom = New Polyline
      ElseIf lngGeomType = esriGeometryPolygon Then
        Set pNewGeom = New Polygon
      End If

      Set pTopoOp = pNewGeom
      pTopoOp.ConstructUnion pGeometryCollection
      pTopoOp.Simplify

      Set pNewGeom.SpatialReference = pSpRef
    End If

    Set UnionGeometries3 = pNewGeom
  End If

  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

  GoTo ClearMemory
ClearMemory:

  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

End Function

Public Function UnionGeometries4(pGeomArray As esriSystem.IVariantArray, _
    Optional lngMaxNumberToUnion As Long = -999) As IGeometry

  Dim pTopoOp As ITopologicalOperator
  Dim pGeom As IGeometry
  Dim pGeometryCollection As IGeometryCollection

  Set pGeometryCollection = New GeometryBag

  Dim pSpRef As ISpatialReference
  Dim pTempGeom As IGeometry
  Dim pNewGeom As IGeometry
  Dim lngIndex As Long
  Dim booFoundGeometry As Boolean

  Do Until lngIndex = pGeomArray.Count Or Not pSpRef Is Nothing
    Set pGeom = pGeomArray.Element(0)
    If Not pGeom Is Nothing Then
      Set pSpRef = pGeom.SpatialReference
      booFoundGeometry = True
    End If
    lngIndex = lngIndex + 1
  Loop

  Dim lngGeomType As esriGeometryType
  lngGeomType = pGeom.GeometryType

  Dim pTempPoly As IPolygon
  Dim pNewPoly As IPolygon
  Dim pSimplifyTopoOp As ITopologicalOperator4

  If Not booFoundGeometry Then
    Set UnionGeometries4 = Nothing
  Else
    For lngIndex = 0 To pGeomArray.Count - 1
      Set pGeom = pGeomArray.Element(lngIndex)

      If Not pGeom Is Nothing Then
        If Not pGeom.IsEmpty Then
          Set pSimplifyTopoOp = pGeom
          pSimplifyTopoOp.IsKnownSimple = False
          pSimplifyTopoOp.Simplify

          pGeometryCollection.AddGeometry pGeom

          If lngMaxNumberToUnion > 1 Then
            If pGeometryCollection.GeometryCount >= lngMaxNumberToUnion Then

              If lngGeomType = esriGeometryPoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryMultipoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryPolyline Then
                Set pTempGeom = New Polyline
              ElseIf lngGeomType = esriGeometryPolygon Then
                Set pTempGeom = New Polygon
              End If

              Set pTopoOp = pTempGeom
              pTopoOp.ConstructUnion pGeometryCollection
              pTopoOp.Simplify

              Set pTempGeom.SpatialReference = pSpRef
              Set pGeometryCollection = New GeometryBag
              pGeometryCollection.AddGeometry pTempGeom

            End If
          End If
        End If
      End If

    Next lngIndex

    If pGeometryCollection.GeometryCount = 1 Then
      Set pNewGeom = pGeometryCollection.Geometry(0)

    Else
      If lngGeomType = esriGeometryPoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryMultipoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryPolyline Then
        Set pNewGeom = New Polyline
      ElseIf lngGeomType = esriGeometryPolygon Then
        Set pNewGeom = New Polygon
      End If

      Set pTopoOp = pNewGeom
      pTopoOp.ConstructUnion pGeometryCollection
      pTopoOp.Simplify

      Set pNewGeom.SpatialReference = pSpRef
    End If

    Set UnionGeometries4 = pNewGeom
  End If

  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

  GoTo ClearMemory
ClearMemory:

  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

End Function

Public Function ReturnPolygonRingsAsDoubleArray(pPolygon As IPolygon) As Variant()

  Dim varReturn() As Variant
  Dim pGeomColl As IGeometryCollection
  Dim pRing As IRing
  Dim lngIndex As Long
  Dim lngCounter As Long
  Dim lngRingCounter As Long
  Dim pPtColl As IPointCollection
  Dim pTestPoint1 As IPoint
  Dim lngIndex2 As Long
  Dim lngPointCount As Long
  Dim pPoint As IPoint
  Dim dblArray() As Double
  Dim booAdd As Boolean
  Dim pClone As IClone
  Dim pWorkPolygon As IPolygon

  If pPolygon.IsEmpty Then
    ReDim varReturn(0)
    varReturn(0) = Null
  Else

    Set pPoint = New Point
    Set pClone = pPolygon
    Set pWorkPolygon = pClone.Clone
    pWorkPolygon.SimplifyPreserveFromTo

    lngRingCounter = -1
    Set pGeomColl = pWorkPolygon
    For lngIndex = 0 To pGeomColl.GeometryCount - 1
      Set pRing = pGeomColl.Geometry(lngIndex)
      If Not pRing.IsEmpty Then
        If Not pRing.IsClosed Then pRing.Close
        Set pPtColl = pRing
        lngCounter = -1
        For lngIndex2 = 0 To pPtColl.PointCount - 1
          pPtColl.QueryPoint lngIndex2, pPoint
          booAdd = True
          If lngCounter > -1 Then
            If pPoint.x = dblArray(0, lngCounter) And pPoint.Y = dblArray(1, lngCounter) Then booAdd = False
          End If

          If booAdd Then
            lngCounter = lngCounter + 1
            ReDim Preserve dblArray(1, lngCounter)
            dblArray(0, lngCounter) = pPoint.x
            dblArray(1, lngCounter) = pPoint.Y
          End If
        Next lngIndex2

        If lngCounter > -1 And ((dblArray(0, 0) <> dblArray(0, lngCounter)) Or (dblArray(0, 0) <> dblArray(0, lngCounter))) Then
          lngCounter = lngCounter + 1
          ReDim Preserve dblArray(1, lngCounter)
          dblArray(0, lngCounter) = dblArray(0, 0)
          dblArray(1, lngCounter) = dblArray(1, 0)
        End If

        lngRingCounter = lngRingCounter + 1
        ReDim Preserve varReturn(lngRingCounter)
        varReturn(lngRingCounter) = dblArray
      End If
    Next lngIndex
  End If

  ReturnPolygonRingsAsDoubleArray = varReturn

End Function

Public Function PointInPoly_Winding(dblPointX As Double, dblPointY As Double, varPolygon() As Variant, _
    Optional lngWindCounter As Long) As Boolean

  Dim lngRingIndex As Long
  Dim dblRing() As Double
  Dim lngPointIndex As Long

  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblSegProportion As Double

  lngWindCounter = 0

  For lngRingIndex = 0 To UBound(varPolygon)
    dblRing = varPolygon(lngRingIndex)
    dblX1 = dblRing(0, 0)
    dblY1 = dblRing(1, 0)

    For lngPointIndex = 0 To UBound(dblRing, 2) - 1
      dblX2 = dblRing(0, lngPointIndex + 1)
      dblY2 = dblRing(1, lngPointIndex + 1)

      If dblY1 <= dblPointY Then
        If dblY2 > dblPointY Then  ' If an upward Crossing
                (dblX1 * dblPointY) + (dblY1 * dblPointX)) > 0 Then lngWindCounter = lngWindCounter + 1
          If CalcCheckClockwiseNumbers2(dblX1, dblY1, dblX2, dblY2, dblPointX, dblPointY) > 0 Then lngWindCounter = lngWindCounter + 1
        End If
      Else  ' AUTOMATICALLY WE KNOW dblY1 > dblPointY; no test needed
        If dblY2 <= dblPointY Then  ' If a downward crossing
                (dblX1 * dblPointY) + (dblY1 * dblPointX)) < 0 Then lngWindCounter = lngWindCounter - 1
          If CalcCheckClockwiseNumbers2(dblX1, dblY1, dblX2, dblY2, dblPointX, dblPointY) < 0 Then lngWindCounter = lngWindCounter - 1
        End If
      End If
      dblX1 = dblX2
      dblY1 = dblY2
    Next lngPointIndex
  Next lngRingIndex

  PointInPoly_Winding = lngWindCounter <> 0

End Function

Public Function CalcCheckClockwiseNumbers2(dblPX As Double, dblPY As Double, dblQX As Double, _
    dblQY As Double, dblRX As Double, dblRY As Double, Optional dblDistance As Double) As Double

  CalcCheckClockwiseNumbers2 = (dblQX * (dblRY - dblPY)) + (dblQY * (dblPX - dblRX)) - (dblPX * dblRY) _
        + (dblPY * dblRX)

End Function

Public Function DistanceBetweenPolygons(booUsingPolygons As Boolean, varSourceObjects_PolysOrDoubleArrays As Variant, _
     Optional dblCloseX1 As Double, Optional dblCloseY1 As Double, _
     Optional dblCloseX2 As Double, Optional dblCloseY2 As Double) As Double

  Dim pPolygon1 As IPolygon
  Dim dblPolyArray1() As Double

  Dim pPolygon2 As IPolygon
  Dim dblPolyArray2() As Double

  Dim varArrays1() As Variant
  Dim varArrays2() As Variant

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngPointIndex1 As Long
  Dim lngPointIndex2 As Long

  If booUsingPolygons Then
    Set pPolygon1 = varSourceObjects_PolysOrDoubleArrays(0)
    Set pPolygon2 = varSourceObjects_PolysOrDoubleArrays(1)

    varArrays1 = MyGeometricOperations.ReturnPolygonRingsAsDoubleArray(pPolygon1)
    varArrays2 = MyGeometricOperations.ReturnPolygonRingsAsDoubleArray(pPolygon2)

  Else
    varArrays1 = varSourceObjects_PolysOrDoubleArrays(0)
    varArrays2 = varSourceObjects_PolysOrDoubleArrays(1)
  End If

  Dim dblMinDist As Double
  Dim dblTestDist As Double
  Dim dblSeg1Start(1) As Double
  Dim dblSeg1End(1) As Double
  Dim dblSeg2Start(1) As Double
  Dim dblSeg2End(1) As Double
  Dim dblClose1() As Double
  Dim dblClose2() As Double

  dblMinDist = 2 ^ 100

  For lngIndex1 = 0 To UBound(varArrays1)
    dblPolyArray1 = varArrays1(lngIndex1)

    For lngIndex2 = 0 To UBound(varArrays2)
      dblPolyArray2 = varArrays2(lngIndex2)

      For lngPointIndex1 = 0 To UBound(dblPolyArray1, 2) - 1
        For lngPointIndex2 = 0 To UBound(dblPolyArray2, 2) - 1

          dblSeg1Start(0) = dblPolyArray1(0, lngPointIndex1)
          dblSeg1Start(1) = dblPolyArray1(1, lngPointIndex1)
          dblSeg1End(0) = dblPolyArray1(0, lngPointIndex1 + 1)
          dblSeg1End(1) = dblPolyArray1(1, lngPointIndex1 + 1)

          dblSeg2Start(0) = dblPolyArray2(0, lngPointIndex2)
          dblSeg2Start(1) = dblPolyArray2(1, lngPointIndex2)
          dblSeg2End(0) = dblPolyArray2(0, lngPointIndex2 + 1)
          dblSeg2End(1) = dblPolyArray2(1, lngPointIndex2 + 1)

          dblTestDist = SquaredDistanceBetweenSegments(dblSeg1Start, dblSeg1End, dblSeg2Start, dblSeg2End, dblClose1, dblClose2)

          If dblTestDist < dblMinDist Then
            dblMinDist = dblTestDist
            dblCloseX1 = dblClose1(0)
            dblCloseY1 = dblClose1(1)
            dblCloseX2 = dblClose2(0)
            dblCloseY2 = dblClose2(1)
          End If
        Next lngPointIndex2
      Next lngPointIndex1
    Next lngIndex2
  Next lngIndex1

  DistanceBetweenPolygons = Sqr(dblMinDist)

ClearMemory:
  Set pPolygon1 = Nothing
  Erase dblPolyArray1
  Set pPolygon2 = Nothing
  Erase dblPolyArray2
  Erase varArrays1
  Erase varArrays2
  Erase dblSeg1Start
  Erase dblSeg1End
  Erase dblSeg2Start
  Erase dblSeg2End
  Erase dblClose1
  Erase dblClose2

End Function


