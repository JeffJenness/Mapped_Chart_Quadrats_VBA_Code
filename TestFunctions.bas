Attribute VB_Name = "TestFunctions"
Option Explicit

Public Sub TestClipSet()

  Debug.Print "-------------------------------"
  
  Dim pMxDoc As IMxDocument
  Dim pFLayer As IFeatureLayer
  Dim pFClass1 As IFeatureClass
  Dim pFClass2 As IFeatureClass
  Dim lngOIDArray1() As Long
  Dim lngOIDArray2() As Long
  
  Set pMxDoc = ThisDocument
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Intersect_set_All_1", pMxDoc.FocusMap)
'  Set pFClass1 = pFLayer.FeatureClass
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Intersect_set_All_2", pMxDoc.FocusMap)
'  Set pFClass2 = pFLayer.FeatureClass
  
'  ReDim lngOIDArray1(3)
'  lngOIDArray1(0) = 1
'  lngOIDArray1(1) = 2
'  lngOIDArray1(2) = 3
'  lngOIDArray1(3) = 4
'  ReDim lngOIDArray2(3)
'  lngOIDArray2(0) = 1
'  lngOIDArray2(1) = 2
'  lngOIDArray2(2) = 3
'  lngOIDArray2(3) = 4
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Intersect_Should_Combine", pMxDoc.FocusMap)
  Set pFClass1 = pFLayer.FeatureClass
  
  ReDim lngOIDArray1(1)
  lngOIDArray1(0) = 1
  lngOIDArray1(1) = 2
  ClipSetOfPolygons pFClass1, lngOIDArray1, pFClass2, lngOIDArray2, pMxDoc
  
  
  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pFLayer = Nothing
  Set pFClass1 = Nothing
  Set pFClass2 = Nothing
  Erase lngOIDArray1
  Erase lngOIDArray2



End Sub

Public Sub ClipSetOfPolygons(pFClass1 As IFeatureClass, lngOIDArray1() As Long, _
    pFClass2 As IFeatureClass, lngOIDArray2() As Long, Optional pMxDoc As IMxDocument)
        
  Dim pBuffCon As IBufferConstruction
  Dim pTransform2D As ITransform2D
  Dim pCentroid As IPoint
  Dim pBuffer As IPolygon
  Dim pTopoOp As ITopologicalOperator3
  
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim pPoly1 As IPolygon
  Dim pPoly2 As IPolygon
  Dim pArea1 As IArea
  Dim pArea2 As IArea
  Dim dblDist As Double
  Dim pEnv As IEnvelope
  Dim dblCloseX1 As Double
  Dim dblCloseY1 As Double
  Dim dblCloseX2 As Double
  Dim dblCloseY2 As Double
  Dim lngCount1 As Long
  Dim lngCount2 As Long
  
  If IsDimmed(lngOIDArray1) Then
    lngCount1 = UBound(lngOIDArray1) + 1
  Else
    lngCount1 = 0
  End If
  
  If IsDimmed(lngOIDArray2) Then
    lngCount2 = UBound(lngOIDArray2) + 1
  Else
    lngCount2 = 0
  End If
  
  If lngCount1 + lngCount2 <= 0 Then Exit Sub
  
  Set pBuffCon = New BufferConstruction
  
  Dim lngCounter As Long
  Dim varPolyArray() As Variant
  Dim lngFClassReferences() As Long
  Dim lngOIDReferences() As Long
  ReDim varPolyArray(lngCount1 + lngCount2 - 1)
  ReDim lngFClassReferences(lngCount1 + lngCount2 - 1)
  ReDim lngOIDReferences(lngCount1 + lngCount2 - 1)
  
  lngCounter = -1
  If lngCount1 > 0 Then
    For lngIndex = 0 To UBound(lngOIDArray1)
      lngCounter = lngCounter + 1
      Set varPolyArray(lngCounter) = pFClass1.GetFeature(lngOIDArray1(lngIndex)).ShapeCopy
      lngFClassReferences(lngCounter) = 1
      lngOIDReferences(lngCounter) = lngOIDArray1(lngIndex)
    Next lngIndex
  End If
  
  If lngCount2 > 0 Then
    For lngIndex = 0 To UBound(lngOIDArray2)
      lngCounter = lngCounter + 1
      Set varPolyArray(lngCounter) = pFClass2.GetFeature(lngOIDArray2(lngIndex)).ShapeCopy
      lngFClassReferences(lngCounter) = 2
      lngOIDReferences(lngCounter) = lngOIDArray2(lngIndex)
    Next lngIndex
  End If
  
  If Not pMxDoc Is Nothing Then
    For lngIndex = 0 To UBound(varPolyArray)
      Set pPoly1 = varPolyArray(lngIndex)
      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoly1, "Delete_me", Nothing
    Next lngIndex
  End If
  
  Dim lngPoly1Source As Long
  Dim lngPoly2Source As Long
  Dim lngPoly1OID As Long
  Dim lngPoly2OID As Long
  Dim pUpdate As IFeatureCursor
  Dim pFeature As IFeature
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix1 As String
  Dim strSuffix1 As String
  Dim strPrefix2 As String
  Dim strSuffix2 As String
  Dim booUpdateFirstPoly As Boolean
  Dim pMergeArray As esriSystem.IVariantArray
  Dim pMergePoly As IPolygon
  
  MyGeneralOperations.ReturnQuerySpecialCharacters pFClass1, strPrefix1, strSuffix1
  MyGeneralOperations.ReturnQuerySpecialCharacters pFClass2, strPrefix2, strSuffix2
  
  Set pQueryFilt = New QueryFilter
  
  For lngIndex = 0 To UBound(varPolyArray) - 1
    Set pPoly1 = varPolyArray(lngIndex)
    Set pArea1 = pPoly1
    lngPoly1Source = lngFClassReferences(lngIndex)
    lngPoly1OID = lngOIDReferences(lngIndex)
        
    For lngIndex2 = lngIndex + 1 To UBound(varPolyArray)
      Set pPoly2 = varPolyArray(lngIndex2)
      Set pArea2 = pPoly2
      lngPoly2Source = lngFClassReferences(lngIndex2)
      lngPoly2OID = lngOIDReferences(lngIndex2)
      
      If lngIndex = 1 And lngIndex2 = 3 Then
        DoEvents
      End If
      
      ' IS IS POSSIBLE POLYGON MIGHT BE EMPTY AT THIS POINT
      If Not pPoly1.IsEmpty And Not pPoly2.IsEmpty Then
      
        dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2), _
            dblCloseX1, dblCloseY1, dblCloseX2, dblCloseY2)
            
        If dblDist < 0.75 Then
          Set pEnv = pPoly1.Envelope
          pEnv.Union pPoly2.Envelope
          Set pCentroid = MyGeneralOperations.Get_Element_Or_Envelope_Point(pEnv, ENUM_Center_Center)
          
          ' SCALE UP
          Set pTransform2D = pPoly1
          With pTransform2D
            .Scale pCentroid, 1000, 1000
          End With
          Set pTransform2D = pPoly2
          With pTransform2D
            .Scale pCentroid, 1000, 1000
          End With
          
          booUpdateFirstPoly = pArea1.Area > pArea2.Area
            
          ' CHECK TO SEE IF EITHER PROPORTION AREA OVERLAP OR PROPORTION LINE OVERLAP IS > 40%
          If CheckShouldCombine(pPoly1, pPoly2) Then
            Debug.Print "Should Combine..."
            Set pMergeArray = New esriSystem.varArray
            pMergeArray.Add pPoly1
            pMergeArray.Add pPoly2
            Set pMergePoly = MyGeometricOperations.UnionGeometries3(pMergeArray, 5)
            Set pMergePoly.SpatialReference = pPoly1.SpatialReference
            
            Set pPoly1 = pMergePoly
            Set pTransform2D = pPoly1
            With pTransform2D
              .Scale pCentroid, 0.001, 0.001
            End With
          
            pPoly2.SetEmpty
            
            ' Both Poly1 and Poly2 have been edited
            Set varPolyArray(lngIndex) = pPoly1  ' NOW THE MERGED VERSION
            Set varPolyArray(lngIndex2) = pPoly2
            
            ' REPLACE POLY 1
            If lngPoly1Source = 1 Then
              ' if Poly1 comes from Feature Class 1
              pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly1OID, "0")
              Set pUpdate = pFClass1.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly1
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            Else
              ' if Poly1 comes from Feature Class 2
              pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly1OID, "0")
              Set pUpdate = pFClass2.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly1
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            End If
            
            ' REPLACE POLY 2
            If lngPoly2Source = 1 Then
              ' if Poly2 comes from Feature Class 1
              pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly2OID, "0")
              Set pUpdate = pFClass1.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly2
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            Else
              ' if Poly2 comes from Feature Class 2
              pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly2OID, "0")
              Set pUpdate = pFClass2.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly2
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            End If
            
          Else
          
            ' CLIP
            If booUpdateFirstPoly Then
              Set pBuffer = pBuffCon.Buffer(pPoly2, 0.75)
              Set pTopoOp = pPoly1
              Set pPoly1 = pTopoOp.Difference(pBuffer)
            Else
              Set pBuffer = pBuffCon.Buffer(pPoly1, 0.75)
              Set pTopoOp = pPoly2
              Set pPoly2 = pTopoOp.Difference(pBuffer)
            End If
          
            ' SCALE BACK DOWN
            Set pTransform2D = pPoly1
            With pTransform2D
              .Scale pCentroid, 0.001, 0.001
            End With
            Set pTransform2D = pPoly2
            With pTransform2D
              .Scale pCentroid, 0.001, 0.001
            End With
          
            ' UPDATE FEATURE CLASS
            If booUpdateFirstPoly Then
              ' if Poly1 is the one that got edited
              Set varPolyArray(lngIndex) = pPoly1
              
              If lngPoly1Source = 1 Then
                ' if Poly1 comes from Feature Class 1
                pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly1OID, "0")
                Set pUpdate = pFClass1.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly1
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              Else
                ' if Poly1 comes from Feature Class 2
                pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly1OID, "0")
                Set pUpdate = pFClass2.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly1
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              End If
            Else
              ' if Poly2 is the one that got edited
              Set varPolyArray(lngIndex2) = pPoly2
              
              If lngPoly2Source = 1 Then
                ' if Poly2 comes from Feature Class 1
                pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly2OID, "0")
                Set pUpdate = pFClass1.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly2
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              Else
                ' if Poly2 comes from Feature Class 2
                pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly2OID, "0")
                Set pUpdate = pFClass2.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly2
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              End If
            End If
          End If
        End If
      End If
    Next lngIndex2
  Next lngIndex

  If Not pMxDoc Is Nothing Then
    For lngIndex = 0 To UBound(varPolyArray)
      Set pPoly1 = varPolyArray(lngIndex)
      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoly1, "Delete_me", Nothing
    Next lngIndex
  End If
  
  DoEvents
  
ClearMemory:
  Set pBuffCon = Nothing
  Set pTransform2D = Nothing
  Set pCentroid = Nothing
  Set pBuffer = Nothing
  Set pTopoOp = Nothing
  Set pPoly1 = Nothing
  Set pPoly2 = Nothing
  Set pArea1 = Nothing
  Set pArea2 = Nothing
  Set pEnv = Nothing
  Erase varPolyArray
  Erase lngFClassReferences
  Erase lngOIDReferences
  Set pUpdate = Nothing
  Set pFeature = Nothing
  Set pQueryFilt = Nothing



End Sub

Public Function CheckShouldCombine(pPoly1 As IPolygon, pPoly2 As IPolygon) As Boolean

  Dim pIntersect As IPolygon
  Dim pTopoOp As ITopologicalOperator
  Dim pArea1 As IArea
  Dim pArea2 As IArea
  Dim pArea3 As IArea
  Dim pLineIntersect As IPolyline
  
  CheckShouldCombine = False
  
  Set pTopoOp = pPoly1
  Set pIntersect = pTopoOp.Intersect(pPoly2, pPoly1.Dimension)
  Set pArea1 = pPoly1
  Set pArea2 = pPoly2
  Set pArea3 = pIntersect
  
  If pArea3.Area / pArea1.Area > 0.4 Or pArea3.Area / pArea2.Area > 0.4 Then
    CheckShouldCombine = True
  Else
    Set pLineIntersect = pTopoOp.Intersect(pPoly2, esriGeometry1Dimension)
    If pLineIntersect.length / pPoly1.length > 0.4 Or pLineIntersect.length / pPoly2.length > 0.4 Then
      CheckShouldCombine = True
    Else
      Set pTopoOp = pPoly2
      Set pLineIntersect = pTopoOp.Intersect(pPoly1, esriGeometry1Dimension)
      If pLineIntersect.length / pPoly1.length > 0.4 Or pLineIntersect.length / pPoly2.length > 0.4 Then
        CheckShouldCombine = True
      End If
    End If
  End If
      
  
ClearMemory:
  Set pIntersect = Nothing
  Set pTopoOp = Nothing
  Set pArea1 = Nothing
  Set pArea2 = Nothing
  Set pArea3 = Nothing
  Set pLineIntersect = Nothing



End Function


Public Sub TestClipFunction()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim strFolder As String
  strFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_August_4_2019\Combined_by_Quadrat.gdb"
  
  Dim strFClassName As String
  strFClassName = "Cover_All"
  
  Dim lngIndex As Long
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pFClass As IFeatureClass
  Dim pFeature As IFeature
  Dim pPolygon1 As IPolygon
  Dim pPolygon2 As IPolygon
  Dim pPolygon3 As IPolygon
  Dim pPolygon4 As IPolygon
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Dim pSpRefRes As ISpatialReferenceResolution
  Dim dblDist2 As Double
  Dim pFill As ISimpleFillSymbol
  Set pFill = New SimpleFillSymbol
  Dim pLine As ISimpleLineSymbol
  Set pLine = New SimpleLineSymbol
  Dim pColor As IRgbColor
  Set pColor = New RgbColor
  pColor.RGB = RGB(127, 0, 0)
  pLine.Color = pColor
  pLine.Style = esriSLSSolid
  pLine.Width = 2
  pFill.Style = esriSFSHollow
  pFill.Outline = pLine
  Dim dblCloseX1 As Double
  Dim dblCloseY1 As Double
  Dim dblCloseX2 As Double
  Dim dblCloseY2 As Double
  
  Dim pRelOp As IRelationalOperator
  Dim pProxOp As IProximityOperator
  Dim pTopoOp As ITopologicalOperator3
  Dim pBuffer As IPolygon
  
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFolder, 0)
  Debug.Print "--------------------"
  Debug.Print strFClassName
  Set pFClass = pWS.OpenFeatureClass(strFClassName)
  
  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference
  Set pSpRefRes = pSpRef
  
  Debug.Print pSpRef.Name
  Debug.Print "Resolution [True] = " & Format(pSpRefRes.XYResolution(True), "0.000000000000")
  Debug.Print "Resolution [False] = " & Format(pSpRefRes.XYResolution(False), "0.000000000000")
  
  ' does not intersect
'  Set pFeature = pFClass.GetFeature(15189)
'  Set pPolygon = pFeature.ShapeCopy
'  Set pFeature = pFClass.GetFeature(15190)
'  Set pPolygon2 = pFeature.ShapeCopyz
  
  ' does intersect
  Set pFeature = pFClass.GetFeature(27976)
  Set pPolygon1 = pFeature.ShapeCopy
  Set pFeature = pFClass.GetFeature(27974)
  Set pPolygon2 = pFeature.ShapeCopy
  Set pFeature = pFClass.GetFeature(27975)
  Set pPolygon3 = pFeature.ShapeCopy
  Set pFeature = pFClass.GetFeature(27972)
  Set pPolygon4 = pFeature.ShapeCopy
  
  Dim pBuffCon As IBufferConstruction
  Set pBuffCon = New BufferConstruction
  Dim lngOIDArray() As Long
  ReDim lngOIDArray(3)
  lngOIDArray(0) = 27976
  lngOIDArray(1) = 27974
  lngOIDArray(2) = 27975
  lngOIDArray(3) = 27972
  
  Call ClipSetOfPolygons(pFClass, lngOIDArray, pMxDoc)
  
  
'  Dim pTransform2D As ITransform2D
'  Dim pCentroid As IPoint
'
'
'  Dim lngIndex1 As Long
'  Dim lngIndex2 As Long
'  Dim pPoly1 As IPolygon
'  Dim pPoly2 As IPolygon
'  Dim pArea1 As IArea
'  Dim pArea2 As IArea
'  Dim dblDist As Double
'  Dim pEnv As IEnvelope
'  Dim pEnv2 As IEnvelope
'
'  Dim varPolyArray() As Variant
'  ReDim varPolyArray(UBound(lngOIDArray))
'
'  For lngIndex = 0 To UBound(lngOIDArray)
'    Set varPolyArray(lngIndex) = pFClass.GetFeature(lngOIDArray(lngIndex)).ShapeCopy
'  Next lngIndex
'
'  For lngIndex = 0 To UBound(lngOIDArray) - 1
''    Set pPoly1 = pFClass.GetFeature(lngOIDArray(lngIndex)).ShapeCopy
'    Set pPoly1 = varPolyArray(lngIndex)
'    Set pArea1 = pPoly1
'
'    For lngIndex2 = lngIndex + 1 To UBound(lngOIDArray)
''      Set pPoly2 = pFClass.GetFeature(lngOIDArray(lngIndex2)).ShapeCopy
'      Set pPoly2 = varPolyArray(lngIndex2)
'      Set pArea2 = pPoly2
'
'      If lngIndex = 1 And lngIndex2 = 3 Then
'        DoEvents
'      End If
'
'      dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2), _
'          dblCloseX1, dblCloseY1, dblCloseX2, dblCloseY2)
'      If dblDist < 0.75 Then
'        Set pEnv = pPoly1.Envelope
'        pEnv.Union pPoly2.Envelope
'        Set pCentroid = MyGeneralOperations.Get_Element_Or_Envelope_Point(pEnv, ENUM_Center_Center)
'        Set pTransform2D = pPoly1
'        With pTransform2D
'          .Scale pCentroid, 1000, 1000
'        End With
'        Set pTransform2D = pPoly2
'        With pTransform2D
'          .Scale pCentroid, 1000, 1000
'        End With
'
'        If pArea1.Area > pArea2.Area Then
'          Set pBuffer = pBuffCon.Buffer(pPoly2, 0.75)
'          Set pTopoOp = pPoly1
'          Set pPoly1 = pTopoOp.Difference(pBuffer)
'
'          ' REPLACE PPOLY1
'          Set varPolyArray(lngIndex) = pPoly1
'        Else
'          Set pBuffer = pBuffCon.Buffer(pPoly1, 0.75)
'          Set pTopoOp = pPoly2
'          Set pPoly2 = pTopoOp.Difference(pBuffer)
'
'          ' REPLACE PPOLY2
'          Set varPolyArray(lngIndex2) = pPoly2
'        End If
'
'        Set pTransform2D = pPoly1
'        With pTransform2D
'          .Scale pCentroid, 0.001, 0.001
'        End With
'        Set pTransform2D = pPoly2
'        With pTransform2D
'          .Scale pCentroid, 0.001, 0.001
'        End With
'
''        Set pPoly2 = pTransform2D.Scale(pCentroid, 100, 100)
'      End If
'    Next lngIndex2
'  Next lngIndex
'
''  Set pTopoOp = pPolygon1
'
'  For lngIndex = 0 To UBound(lngOIDArray)
'    Set pPoly1 = varPolyArray(lngIndex)
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoly1, "Delete_me", pFill
'  Next lngIndex
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pFClass = Nothing
  Set pFeature = Nothing
  Set pPolygon1 = Nothing
  Set pPolygon2 = Nothing
  Set pPolygon3 = Nothing
  Set pPolygon4 = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pSpRefRes = Nothing
  Set pFill = Nothing
  Set pLine = Nothing
  Set pColor = Nothing
  Set pRelOp = Nothing
  Set pProxOp = Nothing
  Set pTopoOp = Nothing
  Set pBuffer = Nothing




End Sub

Public Sub ManuallyCheckIntersect()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim strFolder As String
  strFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_August_4_2019\Combined_by_Quadrat.gdb"
  
  Dim strFClassName As String
  strFClassName = "Cover_All"
  
  Dim lngIndex As Long
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pFClass As IFeatureClass
  Dim pFeature As IFeature
  Dim pPolygon As IPolygon
  Dim pPolygon2 As IPolygon
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Dim pSpRefRes As ISpatialReferenceResolution
  Dim dblDist2 As Double
  
  Dim pRelOp As IRelationalOperator
  Dim pProxOp As IProximityOperator
  Dim pTopoOp As ITopologicalOperator3
  
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFolder, 0)
  Debug.Print "--------------------"
  Debug.Print strFClassName
  Set pFClass = pWS.OpenFeatureClass(strFClassName)
  
  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference
  Set pSpRefRes = pSpRef
  
  Debug.Print pSpRef.Name
  Debug.Print "Resolution [True] = " & Format(pSpRefRes.XYResolution(True), "0.000000000000")
  Debug.Print "Resolution [False] = " & Format(pSpRefRes.XYResolution(False), "0.000000000000")
  
  ' does not intersect
'  Set pFeature = pFClass.GetFeature(15189)
'  Set pPolygon = pFeature.ShapeCopy
'  Set pFeature = pFClass.GetFeature(15190)
'  Set pPolygon2 = pFeature.ShapeCopyz
  
  ' does intersect
  Set pFeature = pFClass.GetFeature(27804)
  Set pPolygon = pFeature.ShapeCopy
  Set pFeature = pFClass.GetFeature(27840)
  Set pPolygon2 = pFeature.ShapeCopy
  
  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me"
  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon2, "Delete_Me"
  
  Set pRelOp = pPolygon
  Debug.Print "Is Disjoint = " & CStr(pRelOp.Disjoint(pPolygon2))
  Debug.Print "Touches = " & CStr(pRelOp.Touches(pPolygon2))
  
  Set pProxOp = pPolygon
  Debug.Print "Prox Op Distance = " & CStr(pProxOp.ReturnDistance(pPolygon2))
  
  Dim dblCloseX1 As Double
  Dim dblCloseY1 As Double
  Dim dblCloseX2 As Double
  Dim dblCloseY2 As Double
  
  dblDist2 = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPolygon, pPolygon2), _
      dblCloseX1, dblCloseY1, dblCloseX2, dblCloseY2)
  
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Set pPoint1 = New Point
  Set pPoint2 = New Point
  Set pPoint1.SpatialReference = pPolygon.SpatialReference
  Set pPoint2.SpatialReference = pPolygon.SpatialReference
  pPoint1.PutCoords dblCloseX1, dblCloseY1
  pPoint2.PutCoords dblCloseX2, dblCloseY2
  
  Debug.Print "Manual Distance = " & Format(dblDist2 * 1000#, "0.000") & " mm"
  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint1, "Delete_Me"
  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint2, "Delete_Me"
    
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pFClass = Nothing
  Set pFeature = Nothing
  Set pPolygon = Nothing
  Set pPolygon2 = Nothing



End Sub
Public Sub CheckEmptyFeatures()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim strFolder As String
  strFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\contemporary data - Original\Q67"
  
  Dim strFClassName As String
  strFClassName = "Q67_2012_D"
  
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pFClass As IFeatureClass
  Dim pFeature As IFeature
  Dim pPoint As IPoint
  
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFolder, 0)
  Debug.Print "--------------------"
  Debug.Print strFClassName
  Set pFClass = pWS.OpenFeatureClass(strFClassName)
  Dim lngIndex As Long
  For lngIndex = 0 To pFClass.FeatureCount(Nothing) - 1
    Set pFeature = pFClass.GetFeature(lngIndex)
    Set pPoint = pFeature.ShapeCopy
    If pPoint.IsEmpty Then
      Debug.Print "FID value " & Format(pFeature.OID, "0") & " empty" ' = " & CStr(pPoint.IsEmpty)
    End If
  Next lngIndex
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pFClass = Nothing
  Set pFeature = Nothing
  Set pPoint = Nothing


End Sub

Public Sub CompareFClassCounts()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim str2017_v1 As String
  Dim str2017_v2 As String
  
  Dim pFolders_v1 As esriSystem.IStringArray
  Dim pFolders_v2 As esriSystem.IStringArray
  
  str2017_v1 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_as_of_Aug_24_2018\" & _
      "2017 Hill Wild Bill digitized 1m2 quadrats"
  str2017_v2 = "D:\arcGIS_stuff\consultation\Margaret_Moore\New_Data_May_29_2019\Hill-WildBill_2017"
  
  Set pFolders_v1 = MyGeneralOperations.ReturnFoldersFromNestedFolders(str2017_v1, "")
  Set pFolders_v2 = MyGeneralOperations.ReturnFoldersFromNestedFolders(str2017_v2, "")

  Dim strPath As String
  Dim lngIndex As Long
  
  Debug.Print "pFolders_v1.Count = " & Format(pFolders_v1.Count, "0")
  Debug.Print "pFolders_v2.Count = " & Format(pFolders_v2.Count, "0")
  
  Dim pDone1 As New Collection
  Dim pDone2 As New Collection
  Dim strNames1() As String
  Dim strNames2() As String
  Dim lngNameIndex As Long
  
  Dim pDatasets As IEnumDataset
  Dim pDataset As IDataset
  Dim strName As String
  
  lngNameIndex = -1
  For lngIndex = 0 To pFolders_v1.Count - 1
    strPath = pFolders_v1.Element(lngIndex)
    Set pDatasets = ReturnDatasets(strPath)
    pDatasets.Reset
    Set pDataset = pDatasets.Next
    Do Until pDataset Is Nothing
      strName = pDataset.BrowseName
      If Not MyGeneralOperations.CheckCollectionForKey(pDone1, strName) Then
        pDone1.Add pDataset, strName
        lngNameIndex = lngNameIndex + 1
        ReDim Preserve strNames1(lngNameIndex)
        strNames1(lngNameIndex) = strName
        ' Debug.Print strName
      End If
      Set pDataset = pDatasets.Next
    Loop
  Next lngIndex
  Debug.Print "Set 1 has " & Format(pDone1.Count) & " feature classes..."
  
  lngNameIndex = -1
  For lngIndex = 0 To pFolders_v2.Count - 1
    strPath = pFolders_v2.Element(lngIndex)
    Set pDatasets = ReturnDatasets(strPath)
    pDatasets.Reset
    Set pDataset = pDatasets.Next
    Do Until pDataset Is Nothing
      strName = pDataset.BrowseName
      If Not MyGeneralOperations.CheckCollectionForKey(pDone2, strName) Then
        pDone2.Add pDataset, strName
        lngNameIndex = lngNameIndex + 1
        ReDim Preserve strNames2(lngNameIndex)
        strNames2(lngNameIndex) = strName
        ' Debug.Print strName
      End If
      Set pDataset = pDatasets.Next
    Loop
  Next lngIndex
  Debug.Print "Set 2 has " & Format(pDone2.Count) & " feature classes..."
  
  Dim strMissingFrom1 As String
  Dim strMissingFrom2 As String
  Dim pGeoDataset As IGeoDataset
  Dim pDataset2 As IDataset
  Dim pGeoDataset2 As IGeoDataset
  Dim pFClass As IFeatureClass
  Dim pFClass2 As IFeatureClass
  Dim lngCount1 As Long
  Dim lngCount2 As Long
  Dim pSpRef As ISpatialReference
  
  lngNameIndex = 0
  For lngIndex = 0 To UBound(strNames2)
    strName = strNames2(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDone1, strName) Then
      lngNameIndex = lngNameIndex + 1
      strMissingFrom1 = strMissingFrom1 & Format(lngNameIndex, "0") & "] " & strName & " Missing" & vbCrLf
    Else
      Set pDataset = pDone1.Item(strName)
      Set pDataset2 = pDone2.Item(strName)
      Set pGeoDataset = pDataset
      Set pGeoDataset2 = pDataset2
      Set pFClass = pDataset
      Set pFClass2 = pDataset2
      lngCount1 = pFClass.FeatureCount(Nothing)
      lngCount2 = pFClass2.FeatureCount(Nothing)
      If lngCount1 <> lngCount2 Then
        strMissingFrom1 = strMissingFrom1 & "--> " & Format(lngNameIndex, "0") & "] " & strName & " Counts: " & _
            Format(lngCount1, "#,##0") & ", " & Format(lngCount2, "#,##0") & vbCrLf
      End If
      If pGeoDataset.SpatialReference.FactoryCode <> pGeoDataset2.SpatialReference.FactoryCode Then
        strMissingFrom1 = strMissingFrom1 & "--> " & Format(lngNameIndex, "0") & "] " & strName & " Spatial References: " & _
            pGeoDataset.SpatialReference.Name & ", " & pGeoDataset2.SpatialReference.Name & vbCrLf
      End If
                  
    End If
  Next lngIndex
  Debug.Print "strMissingFrom1:" & vbCrLf & strMissingFrom1
      
  lngNameIndex = 0
  For lngIndex = 0 To UBound(strNames1)
    strName = strNames1(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pDone2, strName) Then
      lngNameIndex = lngNameIndex + 1
      strMissingFrom2 = strMissingFrom2 & Format(lngNameIndex, "0") & "] " & strName & " Missing" & vbCrLf
    Else
      Set pDataset = pDone1.Item(strName)
      Set pDataset2 = pDone2.Item(strName)
      Set pGeoDataset = pDataset
      Set pGeoDataset2 = pDataset2
      Set pFClass = pDataset
      Set pFClass2 = pDataset2
      lngCount1 = pFClass.FeatureCount(Nothing)
      lngCount2 = pFClass2.FeatureCount(Nothing)
      If lngCount1 <> lngCount2 Then
        strMissingFrom1 = strMissingFrom1 & "--> " & Format(lngNameIndex, "0") & "] " & strName & " Counts: " & _
            Format(lngCount1, "#,##0") & ", " & Format(lngCount2, "#,##0") & vbCrLf
      End If
      If pGeoDataset.SpatialReference.FactoryCode <> pGeoDataset2.SpatialReference.FactoryCode Then
        strMissingFrom1 = strMissingFrom1 & "--> " & Format(lngNameIndex, "0") & "] " & strName & " Spatial References: " & _
            pGeoDataset.SpatialReference.Name & ", " & pGeoDataset2.SpatialReference.Name & vbCrLf
      End If
    End If
  Next lngIndex
  Debug.Print "strMissingFrom2:" & vbCrLf & strMissingFrom2
  
  Debug.Print "Done..."
  
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders_v1 = Nothing
  Set pFolders_v2 = Nothing


  
End Sub

Public Function ReturnDatasets(strPath As String) As IEnumDataset
  On Error GoTo ErrHandler
  
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strPath, 0)
  
  Set ReturnDatasets = pWS.Datasets(esriDTFeatureClass)
  GoTo ClearMemory
  
ErrHandler:
  Set ReturnDatasets = Nothing
  
ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing

End Function

Public Sub testNoData()
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Dim pLayers As esriSystem.IVariantArray
  Set pLayers = MyGeneralOperations.ReturnLayersByType(pMxDoc.FocusMap, ENUM_jenRasterLayers)
  Dim pRLayer As IRasterLayer
  Dim pRaster As IRaster
  Dim pRastProps As IRasterProps
  Dim varVal As Variant
  Dim pRastBandColl As IRasterBandCollection
  Dim pRastBand As IRasterBand
  Dim lngIndex As Long
  For lngIndex = 0 To pLayers.Count - 1
    Set pRLayer = pLayers.Element(lngIndex)
    Set pRaster = pRLayer.Raster
    Set pRastBandColl = pRaster
    Set pRastBand = pRastBandColl.Item(0)
    Set pRastProps = pRastBand
    varVal = pRastProps.NoDataValue
    Debug.Print CStr(lngIndex + 1) & "] " & pRLayer.Name & ": " & CStr(varVal)
  Next lngIndex
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pLayers = Nothing
  Set pRLayer = Nothing
  Set pRaster = Nothing
  Set pRastProps = Nothing
  varVal = Null
  Set pRastBandColl = Nothing
  Set pRastBand = Nothing


End Sub
