Attribute VB_Name = "TestFunctions"
Option Explicit

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

      If Not pPoly1.IsEmpty And Not pPoly2.IsEmpty Then

        dblDist = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPoly1, pPoly2), _
            dblCloseX1, dblCloseY1, dblCloseX2, dblCloseY2)

        If dblDist < 0.75 Then
          Set pEnv = pPoly1.Envelope
          pEnv.Union pPoly2.Envelope
          Set pCentroid = MyGeneralOperations.Get_Element_Or_Envelope_Point(pEnv, ENUM_Center_Center)

          Set pTransform2D = pPoly1
          With pTransform2D
            .Scale pCentroid, 1000, 1000
          End With
          Set pTransform2D = pPoly2
          With pTransform2D
            .Scale pCentroid, 1000, 1000
          End With

          booUpdateFirstPoly = pArea1.Area > pArea2.Area

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

            Set varPolyArray(lngIndex) = pPoly1  ' NOW THE MERGED VERSION
            Set varPolyArray(lngIndex2) = pPoly2

            If lngPoly1Source = 1 Then
              pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly1OID, "0")
              Set pUpdate = pFClass1.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly1
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            Else
              pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly1OID, "0")
              Set pUpdate = pFClass2.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly1
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            End If

            If lngPoly2Source = 1 Then
              pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly2OID, "0")
              Set pUpdate = pFClass1.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly2
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            Else
              pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly2OID, "0")
              Set pUpdate = pFClass2.Update(pQueryFilt, False)
              Set pFeature = pUpdate.NextFeature
              Set pFeature.Shape = pPoly2
              pUpdate.UpdateFeature pFeature
              pUpdate.Flush
            End If

          Else

            If booUpdateFirstPoly Then
              Set pBuffer = pBuffCon.Buffer(pPoly2, 0.75)
              Set pTopoOp = pPoly1
              Set pPoly1 = pTopoOp.Difference(pBuffer)
            Else
              Set pBuffer = pBuffCon.Buffer(pPoly1, 0.75)
              Set pTopoOp = pPoly2
              Set pPoly2 = pTopoOp.Difference(pBuffer)
            End If

            Set pTransform2D = pPoly1
            With pTransform2D
              .Scale pCentroid, 0.001, 0.001
            End With
            Set pTransform2D = pPoly2
            With pTransform2D
              .Scale pCentroid, 0.001, 0.001
            End With

            If booUpdateFirstPoly Then
              Set varPolyArray(lngIndex) = pPoly1

              If lngPoly1Source = 1 Then
                pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly1OID, "0")
                Set pUpdate = pFClass1.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly1
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              Else
                pQueryFilt.WhereClause = strPrefix2 & pFClass2.OIDFieldName & strSuffix2 & " = " & Format(lngPoly1OID, "0")
                Set pUpdate = pFClass2.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly1
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              End If
            Else
              Set varPolyArray(lngIndex2) = pPoly2

              If lngPoly2Source = 1 Then
                pQueryFilt.WhereClause = strPrefix1 & pFClass1.OIDFieldName & strSuffix1 & " = " & Format(lngPoly2OID, "0")
                Set pUpdate = pFClass1.Update(pQueryFilt, False)
                Set pFeature = pUpdate.NextFeature
                Set pFeature.Shape = pPoly2
                pUpdate.UpdateFeature pFeature
                pUpdate.Flush
              Else
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
  Dim pRelOp As IRelationalOperator

  CheckShouldCombine = False

  Set pTopoOp = pPoly1
  Set pIntersect = pTopoOp.Intersect(pPoly2, pPoly1.Dimension)
  Set pArea1 = pPoly1
  Set pArea2 = pPoly2
  Set pArea3 = pIntersect

  If pArea3.Area / pArea1.Area > 0.4 Or pArea3.Area / pArea2.Area > 0.4 Then
    CheckShouldCombine = True
  Else
    Set pRelOp = pPoly1
    If Not pRelOp.Disjoint(pPoly2) Then
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
  End If

ClearMemory:
  Set pIntersect = Nothing
  Set pTopoOp = Nothing
  Set pArea1 = Nothing
  Set pArea2 = Nothing
  Set pArea3 = Nothing
  Set pLineIntersect = Nothing

End Function

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


