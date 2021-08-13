Attribute VB_Name = "QuickSort"
Option Explicit
Option Compare Binary
   
Public Enum JenVariableTypes
  enum_TypeString = 1
  enum_TypeDouble = 2
  enum_TypeLong = 4
  enum_TypeDate = 8
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
' see http://vbnet.mvps.org/index.html?code/sort/qsvariations.htm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODIFIED JAN. 4, 2006 BY JEFF JENNESS, TO SIMPLIFY IMPLEMENTATION IN ARCGIS
'-----------------------------------------------------------------------------

Public Sub MultiSort(varArray() As Variant, varTypes() As Variant, lngCaseSensitive As VbCompareMethod)

  ' SORTS IN COLUMN ORDER
  ' ACCEPTS 2-DIMENSIONAL VARIANT ARRAY CONTAINING STRINGS, SINGLES, DOUBLES AND/OR DATES
  ' lngCaseSensitive ONLY CONSIDERED IF COLUMN CONTAINS STRINGS
  
  Dim lngColIndex As Long
  Dim lngIndex2 As Long
  Dim lngType As JenVariableTypes
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim lngIndex As Long
'  Dim varTest As Variant
  Dim lngTestIndex As Long
  Dim varRow() As Variant
      
  lngType = varTypes(0)
  QuickSort.VariantAscending_TwoDimensional varArray, 0, UBound(varArray, 2), 0, _
      UBound(varArray, 1), lngType, lngCaseSensitive
    
  If UBound(varArray, 1) > 0 And UBound(varArray, 2) > 0 Then
    For lngColIndex = 1 To UBound(varArray, 1)
      lngType = varTypes(lngColIndex)
'      varTest = varArray(lngColIndex - 1, 0)
      varRow = ReturnRow(varArray, 0)
      lngStart = 0
      For lngIndex = 1 To UBound(varArray, 2)
        If CheckIfRowDifferent(varArray, varRow, lngColIndex - 1, lngIndex, varTypes, lngCaseSensitive) Then
        
          QuickSort.VariantAscending_TwoDimensional varArray, lngStart, lngIndex - 1, lngColIndex, _
              UBound(varArray, 1), lngType, lngCaseSensitive
              
          lngStart = lngIndex
'          varTest = varArray(lngColIndex - 1, lngIndex)
          varRow = ReturnRow(varArray, lngIndex)
        End If
      Next lngIndex
      
      QuickSort.VariantAscending_TwoDimensional varArray, lngStart, UBound(varArray, 2), lngColIndex, _
          UBound(varArray, 1), lngType, lngCaseSensitive
              
    Next lngColIndex
  End If

End Sub

Public Function CheckIfRowDifferent(varArray() As Variant, varTest() As Variant, _
    lngEndIndex As Long, lngRowToTest As Long, varTypes() As Variant, _
    lngCaseSensitive As VbCompareMethod) As Boolean
  
  CheckIfRowDifferent = False
  Dim lngIndex As Long
  Dim lngType As JenVariableTypes
  For lngIndex = lngEndIndex To 0 Step -1
    lngType = varTypes(lngIndex)
    If lngType = enum_TypeString Then
      If StrComp(CStr(varArray(lngIndex, lngRowToTest)), CStr(varTest(lngIndex)), lngType) <> 0 Then
        CheckIfRowDifferent = True
        Exit For
      End If
    Else
      If varArray(lngIndex, lngRowToTest) <> varTest(lngIndex) Then
        CheckIfRowDifferent = True
        Exit For
      End If
    End If
  Next lngIndex

End Function

Public Function ReturnRow(varArray() As Variant, lngRow As Long) As Variant()

  ' assumes a 2-dimensional array
  Dim lngIndex As Long
  Dim varReturn() As Variant
  ReDim varReturn(UBound(varArray, 1))
  
  For lngIndex = 0 To UBound(varArray, 1)
    varReturn(lngIndex) = varArray(lngIndex, lngRow)
  Next lngIndex
  ReturnRow = varReturn
  
  Erase varReturn

End Function


Public Sub VariantAscending_TwoDimensional(narray() As Variant, inLow As Long, inHi As Long, _
    lngSortColumn As Long, lngMaxColumNumber As Long, lngJenVarType As JenVariableTypes, _
    lngCaseSensitive As VbCompareMethod)
  
  ' SORTS BY lngSortColumn, 0-BASED
  ' ASSUMES INDEX 1 IS COLUMN
  ' ASSUMES INDEX 2 IS ROW
  ' EXAMPLE: SORT ON SECOND COLUMN (0-BASED, SO SORT COLUMN = "1")
  '  QuickSort.DoubleAscending_TwoDimensional dblArray, 0, UBound(dblArray, 2), 1, UBound(dblArray, 1)
  
  Dim pivot As Variant
  Dim tmpSwap As Variant
  Dim tmpLow As Long
  Dim tmpHi  As Long
  Dim lngIndex As Long
   
  tmpLow = inLow
  tmpHi = inHi
  Select Case lngJenVarType
    Case enum_TypeString
      pivot = CStr(narray(lngSortColumn, (inLow + inHi) / 2))
    Case enum_TypeDouble
      pivot = CDbl(narray(lngSortColumn, (inLow + inHi) / 2))
    Case enum_TypeLong
      pivot = CLng(narray(lngSortColumn, (inLow + inHi) / 2))
    Case enum_TypeDate
      pivot = DateToJulian(CDate(narray(lngSortColumn, (inLow + inHi) / 2)))
  End Select
    
  While (tmpLow <= tmpHi)
       
    Select Case lngJenVarType
      Case enum_TypeString
        While ((StrComp(CStr(narray(lngSortColumn, tmpLow)), CStr(pivot), lngCaseSensitive) < 0) And tmpLow < inHi)
           tmpLow = tmpLow + 1
        Wend
        While ((StrComp(CStr(pivot), CStr(narray(lngSortColumn, tmpHi)), lngCaseSensitive) < 0) And tmpHi > inLow)
           tmpHi = tmpHi - 1
        Wend
      
'        While (CStr(narray(lngSortColumn, tmpLow)) < pivot And tmpLow < inHi)
'           tmpLow = tmpLow + 1
'        Wend
'        While (pivot < CStr(narray(lngSortColumn, tmpHi)) And tmpHi > inLow)
'           tmpHi = tmpHi - 1
'        Wend
      Case enum_TypeDouble
        While (CDbl(narray(lngSortColumn, tmpLow)) < pivot And tmpLow < inHi)
           tmpLow = tmpLow + 1
        Wend
        While (pivot < CDbl(narray(lngSortColumn, tmpHi)) And tmpHi > inLow)
           tmpHi = tmpHi - 1
        Wend
      Case enum_TypeLong
        While (CLng(narray(lngSortColumn, tmpLow)) < pivot And tmpLow < inHi)
           tmpLow = tmpLow + 1
        Wend
        While (pivot < CLng(narray(lngSortColumn, tmpHi)) And tmpHi > inLow)
           tmpHi = tmpHi - 1
        Wend
      Case enum_TypeDate
        While (DateToJulian(CDate(narray(lngSortColumn, tmpLow))) < pivot And tmpLow < inHi)
           tmpLow = tmpLow + 1
        Wend
        While (pivot < DateToJulian(CDate(narray(lngSortColumn, tmpHi))) And tmpHi > inLow)
           tmpHi = tmpHi - 1
        Wend
    End Select
    
    If (tmpLow <= tmpHi) Then
      For lngIndex = 0 To lngMaxColumNumber
        tmpSwap = narray(lngIndex, tmpLow)
        narray(lngIndex, tmpLow) = narray(lngIndex, tmpHi)
        narray(lngIndex, tmpHi) = tmpSwap
      Next lngIndex
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then VariantAscending_TwoDimensional narray(), inLow, tmpHi, _
      lngSortColumn, lngMaxColumNumber, lngJenVarType, lngCaseSensitive
  If (tmpLow < inHi) Then VariantAscending_TwoDimensional narray(), tmpLow, inHi, _
      lngSortColumn, lngMaxColumNumber, lngJenVarType, lngCaseSensitive

End Sub

Public Sub AngleSort(pAngleInDegrees As esriSystem.IDoubleArray, Optional booSortClockwise As Boolean = True, _
    Optional dblCentralAngle As Double = -999)
    
  Dim dblAngles() As Double
  Dim lngIndex As Long
  ReDim dblAngles(pAngleInDegrees.Count - 1)
  
  For lngIndex = 0 To pAngleInDegrees.Count - 1
    dblAngles(lngIndex) = pAngleInDegrees.Element(lngIndex)
  Next lngIndex
  
  DoubleAscending dblAngles, 0, UBound(dblAngles)
  
  
'  Debug.Print
'  Debug.Print "Sorted ..."
'  Dim dblDebugGap As Double
'  Dim strDebugGap As String
'  For lngIndex = 0 To UBound(dblAngles)
'    If lngIndex = 0 Then
'      dblDebugGap = dblAngles(0) + 360 - dblAngles(UBound(dblAngles))
'    Else
'      dblDebugGap = dblAngles(lngIndex) - dblAngles(lngIndex - 1)
'    End If
'    strDebugGap = CStr(Format(dblDebugGap, "0"))
'    Debug.Print CStr(lngIndex + 1) & "]  " & CStr(Format(dblAngles(lngIndex), "0.00")) & "       Gap = " & strDebugGap
'  Next lngIndex
  
  Dim pTempArray As esriSystem.IDoubleArray
  Set pTempArray = New esriSystem.DoubleArray
  
  Dim lngSplitIndex As Long
  
  If dblCentralAngle <> -999 Then
    
    Dim dblSplitAngle As Double
    dblSplitAngle = dblCentralAngle - 180
    If dblSplitAngle < 0 Then
      dblSplitAngle = dblSplitAngle + 360
    End If
    
    lngSplitIndex = 0
    Do While lngSplitIndex <= UBound(dblAngles)
      If dblAngles(lngSplitIndex) > dblSplitAngle Then
        Exit Do
      End If
      lngSplitIndex = lngSplitIndex + 1
    Loop
    
    If lngSplitIndex <= UBound(dblAngles) Then
      For lngIndex = lngSplitIndex To UBound(dblAngles)
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
    
    If lngSplitIndex > 0 Then
      For lngIndex = 0 To lngSplitIndex - 1
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
    
  Else
  
    Dim dblLargestGap As Double
    dblLargestGap = dblAngles(0) + 360 - dblAngles(UBound(dblAngles))
    lngSplitIndex = 0
    Dim dblTempGap As Double
    
    For lngIndex = 0 To UBound(dblAngles) - 1
      dblTempGap = dblAngles(lngIndex + 1) - dblAngles(lngIndex)
      If dblTempGap > dblLargestGap Then
        dblLargestGap = dblTempGap
        lngSplitIndex = lngIndex + 1
      End If
    Next lngIndex
    
    
    If lngSplitIndex <= UBound(dblAngles) Then
      For lngIndex = lngSplitIndex To UBound(dblAngles)
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
    
    If lngSplitIndex > 0 Then
      For lngIndex = 0 To lngSplitIndex - 1
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
  
  End If
  
'  Debug.Print
'  Debug.Print "Sorted and Wrapped..."
'  For lngIndex = 0 To UBound(dblAngles)
'    If lngIndex = 0 Then
'      dblDebugGap = pTempArray.Element(0) + 360 - pTempArray.Element(UBound(dblAngles))
'    Else
'      dblDebugGap = pTempArray.Element(lngIndex) - pTempArray.Element(lngIndex - 1)
'    End If
'    If dblDebugGap > 360 Then
'      dblDebugGap = dblDebugGap - 360
'    ElseIf dblDebugGap < 0 Then
'      dblDebugGap = dblDebugGap + 360
'    End If
'    strDebugGap = CStr(Format(dblDebugGap, "0"))
'    Debug.Print CStr(lngIndex + 1) & "]  " & CStr(Format(pTempArray.Element(lngIndex), "0.00")) & "       Gap = " & strDebugGap
'  Next lngIndex
  
  pAngleInDegrees.RemoveAll
  If booSortClockwise Then
    For lngIndex = 0 To pTempArray.Count - 1
      pAngleInDegrees.Add pTempArray.Element(lngIndex)
    Next lngIndex
  Else
    For lngIndex = pTempArray.Count - 1 To 0 Step -1
      pAngleInDegrees.Add pTempArray.Element(lngIndex)
    Next lngIndex
  End If

End Sub
Public Sub ByteAscending(narray() As Byte, inLow As Long, inHi As Long)

   Dim pivot As Byte
   Dim tmpSwap As Byte
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then ByteAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then ByteAscending narray(), tmpLow, inHi

End Sub


Public Sub ByteDescending(narray() As Byte, inLow As Long, inHi As Long)

   Dim pivot As Byte
   Dim tmpSwap As Byte
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then ByteDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then ByteDescending narray(), tmpLow, inHi

End Sub
Public Sub LongAscending(narray() As Long, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then LongAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then LongAscending narray(), tmpLow, inHi

End Sub


Public Sub LongDescending(narray() As Long, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then LongDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then LongDescending narray(), tmpLow, inHi

End Sub




Public Sub SingleAscending(narray() As Single, inLow As Long, inHi As Long)

   Dim pivot As Single
   Dim tmpSwap As Single
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then SingleAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then SingleAscending narray(), tmpLow, inHi

End Sub


Public Sub SingleDescending(narray() As Single, inLow As Long, inHi As Long)

   Dim pivot As Single
   Dim tmpSwap As Single
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then SingleDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then SingleDescending narray(), tmpLow, inHi

End Sub
Public Sub DoubleAscendingWithObjects(narray() As Double, varObjArray() As Variant, inLow As Long, inHi As Long)

  Dim pivot As Double
  Dim tmpSwap As Double
  Dim tmpSizeSwap As Variant
  Dim tmpLow As Long
  Dim tmpHi  As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = narray((inLow + inHi) / 2)
  While (tmpLow <= tmpHi)
       
    While (narray(tmpLow) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < narray(tmpHi) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
       tmpSwap = narray(tmpLow)
       narray(tmpLow) = narray(tmpHi)
       narray(tmpHi) = tmpSwap
       
       tmpSizeSwap = varObjArray(tmpLow)
       varObjArray(tmpLow) = varObjArray(tmpHi)
       varObjArray(tmpHi) = tmpSizeSwap
       
       tmpLow = tmpLow + 1
       tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DoubleAscendingWithObjects narray(), varObjArray(), inLow, tmpHi
  If (tmpLow < inHi) Then DoubleAscendingWithObjects narray(), varObjArray(), tmpLow, inHi

End Sub

Public Sub StringsAscendingWithObjects(sarray() As String, varObjArray() As Variant, inLow As Long, inHi As Long)

   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpSizeSwap As Variant
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
       
         tmpSizeSwap = varObjArray(tmpLow)
         varObjArray(tmpLow) = varObjArray(tmpHi)
         varObjArray(tmpHi) = tmpSizeSwap
        
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then StringsAscendingWithObjects sarray(), varObjArray(), inLow, tmpHi
   If (tmpLow < inHi) Then StringsAscendingWithObjects sarray(), varObjArray(), tmpLow, inHi

End Sub


Public Sub DoubleAscendingWithSizes(narray() As Double, nSizeArray() As Double, inLow As Long, inHi As Long)

  Dim pivot As Double
  Dim tmpSwap As Double
  Dim tmpSizeSwap As Double
  Dim tmpLow As Long
  Dim tmpHi  As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = narray((inLow + inHi) / 2)
  While (tmpLow <= tmpHi)
       
    While (narray(tmpLow) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < narray(tmpHi) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
       tmpSwap = narray(tmpLow)
       narray(tmpLow) = narray(tmpHi)
       narray(tmpHi) = tmpSwap
       
       tmpSizeSwap = nSizeArray(tmpLow)
       nSizeArray(tmpLow) = nSizeArray(tmpHi)
       nSizeArray(tmpHi) = tmpSizeSwap
       
       tmpLow = tmpLow + 1
       tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DoubleAscendingWithSizes narray(), nSizeArray(), inLow, tmpHi
  If (tmpLow < inHi) Then DoubleAscendingWithSizes narray(), nSizeArray(), tmpLow, inHi

End Sub

Public Sub DoubleAscending_TwoDimensional(narray() As Double, inLow As Long, inHi As Long, _
    lngSortColumn As Long, lngMaxColumNumber As Long)
  
  ' SORTS BY lngSortColumn, 0-BASED
  ' ASSUMES INDEX 1 IS COLUMN
  ' ASSUMES INDEX 2 IS ROW
  ' EXAMPLE: SORT ON SECOND COLUMN (0-BASED, SO SORT COLUMN = "1")
  '  QuickSort.DoubleAscending_TwoDimensional dblArray, 0, UBound(dblArray, 2), 1, UBound(dblArray, 1)
  
  Dim pivot As Double
  Dim tmpSwap As Double
  Dim tmpSizeSwap As Double
  Dim tmpLow As Long
  Dim tmpHi  As Long
  Dim lngIndex As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = narray(lngSortColumn, (inLow + inHi) / 2)
  While (tmpLow <= tmpHi)
       
    While (narray(lngSortColumn, tmpLow) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < narray(lngSortColumn, tmpHi) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
      For lngIndex = 0 To lngMaxColumNumber
        tmpSwap = narray(lngIndex, tmpLow)
        narray(lngIndex, tmpLow) = narray(lngIndex, tmpHi)
        narray(lngIndex, tmpHi) = tmpSwap
      Next lngIndex
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DoubleAscending_TwoDimensional narray(), inLow, tmpHi, lngSortColumn, lngMaxColumNumber
  If (tmpLow < inHi) Then DoubleAscending_TwoDimensional narray(), tmpLow, inHi, lngSortColumn, lngMaxColumNumber

End Sub

Public Sub DoubleAscending(narray() As Double, inLow As Long, inHi As Long)

   Dim pivot As Double
   Dim tmpSwap As Double
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DoubleAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DoubleAscending narray(), tmpLow, inHi

End Sub


Public Sub DoubleDescending(narray() As Double, inLow As Long, inHi As Long)

   Dim pivot As Double
   Dim tmpSwap As Double
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DoubleDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DoubleDescending narray(), tmpLow, inHi

End Sub

Public Sub StringsAscending(sarray() As String, inLow As Long, inHi As Long)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then StringsAscending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then StringsAscending sarray(), tmpLow, inHi

End Sub
Public Sub StringAscending_TwoDimensional(narray() As String, inLow As Long, inHi As Long, _
    lngSortColumn As Long, lngMaxColumNumber As Long)
  
  ' SORTS BY lngSortColumn, 0-BASED
  ' ASSUMES INDEX 1 IS COLUMN
  ' ASSUMES INDEX 2 IS ROW
  ' EXAMPLE: SORT ON SECOND COLUMN (0-BASED, SO SORT COLUMN = "1")
  '  QuickSort.DoubleAscending_TwoDimensional dblArray, 0, UBound(dblArray, 2), 1, UBound(dblArray, 1)
  
  Dim pivot As String
  Dim tmpSwap As String
  Dim tmpSizeSwap As Double
  Dim tmpLow As Long
  Dim tmpHi  As Long
  Dim lngIndex As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = narray(lngSortColumn, (inLow + inHi) / 2)
  While (tmpLow <= tmpHi)
       
    While (narray(lngSortColumn, tmpLow) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < narray(lngSortColumn, tmpHi) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
      For lngIndex = 0 To lngMaxColumNumber
        tmpSwap = narray(lngIndex, tmpLow)
        narray(lngIndex, tmpLow) = narray(lngIndex, tmpHi)
        narray(lngIndex, tmpHi) = tmpSwap
      Next lngIndex
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then StringAscending_TwoDimensional narray(), inLow, tmpHi, lngSortColumn, lngMaxColumNumber
  If (tmpLow < inHi) Then StringAscending_TwoDimensional narray(), tmpLow, inHi, lngSortColumn, lngMaxColumNumber

End Sub

Public Sub StringsDescending(sarray() As String, inLow As Long, inHi As Long)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
      
      While (sarray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
    
      While (pivot > sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
  
   Wend
  
   If (inLow < tmpHi) Then StringsDescending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then StringsDescending sarray(), tmpLow, inHi

End Sub


Public Sub VariantAscending(sarray() As Variant, inLow As Long, inHi As Long)
  
   Dim pivot As Variant
   Dim tmpSwap As Variant
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then VariantAscending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then VariantAscending sarray(), tmpLow, inHi

End Sub


Public Sub VariantDescending(sarray() As Variant, inLow As Long, inHi As Long)
  
   Dim pivot As Variant
   Dim tmpSwap As Variant
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
      
      While (sarray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
    
      While (pivot > sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
  
   Wend
  
   If (inLow < tmpHi) Then VariantDescending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then VariantDescending sarray(), tmpLow, inHi

End Sub

Public Sub DatesDescending(narray() As Date, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = DateToJulian(narray((inLow + inHi) / 2))
   
   While (tmpLow <= tmpHi)
        
      While DateToJulian(narray(tmpLow)) > pivot And (tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > DateToJulian(narray(tmpHi))) And (tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DatesDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DatesDescending narray(), tmpLow, inHi

End Sub


Public Sub DatesAscending(narray() As Date, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = DateToJulian(narray((inLow + inHi) / 2))

   While (tmpLow <= tmpHi)
       
      While (DateToJulian(narray(tmpLow)) < pivot) And (tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < DateToJulian(narray(tmpHi))) And (tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
      
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
         
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DatesAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DatesAscending narray(), tmpLow, inHi

End Sub
Public Sub DatesAscendingWithObjects(narray() As Date, nPairedArray() As Variant, inLow As Long, inHi As Long)

  Dim pivot As Double
  Dim tmpSwap As Double
  Dim tmpSizePair As Variant
  Dim tmpLow As Long
  Dim tmpHi  As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = DateToJulian(narray((inLow + inHi) / 2))
  While (tmpLow <= tmpHi)
       
    While DateToJulian((narray(tmpLow)) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < DateToJulian(narray(tmpHi)) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
       tmpSwap = narray(tmpLow)
       narray(tmpLow) = narray(tmpHi)
       narray(tmpHi) = tmpSwap
       
       tmpSizePair = nPairedArray(tmpLow)
       nPairedArray(tmpLow) = nPairedArray(tmpHi)
       nPairedArray(tmpHi) = tmpSizePair
       
       tmpLow = tmpLow + 1
       tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DatesAscendingWithObjects narray(), nPairedArray(), inLow, tmpHi
  If (tmpLow < inHi) Then DatesAscendingWithObjects narray(), nPairedArray(), tmpLow, inHi

End Sub
Public Sub DatesAscending_TwoDimensional(narray() As Date, inLow As Long, inHi As Long, _
    lngSortColumn As Long, lngMaxColumNumber As Long)
  
  ' SORTS BY lngSortColumn, 0-BASED
  ' ASSUMES INDEX 1 IS COLUMN
  ' ASSUMES INDEX 2 IS ROW
  ' EXAMPLE: SORT ON SECOND COLUMN (0-BASED, SO SORT COLUMN = "1")
  '  QuickSort.DoubleAscending_TwoDimensional dblArray, 0, UBound(dblArray, 2), 1, UBound(dblArray, 1)
  
  Dim pivot As String
  Dim tmpSwap As String
  Dim tmpSizeSwap As Double
  Dim tmpLow As Long
  Dim tmpHi  As Long
  Dim lngIndex As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = DateToJulian(narray((inLow + inHi) / 2))
  
  While (tmpLow <= tmpHi)
       
    While (DateToJulian(narray(lngSortColumn, tmpLow)) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < DateToJulian(narray(lngSortColumn, tmpHi)) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
      For lngIndex = 0 To lngMaxColumNumber
        tmpSwap = narray(lngIndex, tmpLow)
        narray(lngIndex, tmpLow) = narray(lngIndex, tmpHi)
        narray(lngIndex, tmpHi) = tmpSwap
      Next lngIndex
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DatesAscending_TwoDimensional narray(), inLow, tmpHi, lngSortColumn, lngMaxColumNumber
  If (tmpLow < inHi) Then DatesAscending_TwoDimensional narray(), tmpLow, inHi, lngSortColumn, lngMaxColumNumber

End Sub

Private Function DateToJulian(MyDate As Date) As Long

  'Return a numeric value representing
  'the passed date
   DateToJulian = DateValue(MyDate)

End Function

Public Sub SortStringArray(pStringArray As esriSystem.IStringArray)

  Dim strArray() As String
  ReDim strArray(pStringArray.Count - 1)
  Dim lngIndex As Long
  For lngIndex = 0 To pStringArray.Count - 1
    strArray(lngIndex) = pStringArray.Element(lngIndex)
  Next lngIndex
  QuickSort.StringsAscending strArray, LBound(strArray), UBound(strArray)
  pStringArray.RemoveAll
  For lngIndex = 0 To UBound(strArray)
    pStringArray.Add strArray(lngIndex)
  Next lngIndex

End Sub
Public Sub SortDoubleArray(pDoubleArray As esriSystem.IDoubleArray)

  Dim dblArray() As Double
  ReDim dblArray(pDoubleArray.Count - 1)
  Dim lngIndex As Long
  For lngIndex = 0 To pDoubleArray.Count - 1
    dblArray(lngIndex) = pDoubleArray.Element(lngIndex)
  Next lngIndex
  QuickSort.DoubleAscending dblArray, LBound(dblArray), UBound(dblArray)
  pDoubleArray.RemoveAll
  For lngIndex = 0 To UBound(dblArray)
    pDoubleArray.Add dblArray(lngIndex)
  Next lngIndex

End Sub






