Attribute VB_Name = "QuickSort"
Option Explicit

Public Enum JenVariableTypes
  enum_TypeString = 1
  enum_TypeDouble = 2
  enum_TypeLong = 4
  enum_TypeDate = 8
End Enum

Public Sub MultiSort(varArray() As Variant, varTypes() As Variant, lngCaseSensitive As VbCompareMethod)

  Dim lngColIndex As Long
  Dim lngIndex2 As Long
  Dim lngType As JenVariableTypes
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim lngIndex As Long
  Dim lngTestIndex As Long
  Dim varRow() As Variant

  lngType = varTypes(0)
  QuickSort.VariantAscending_TwoDimensional varArray, 0, UBound(varArray, 2), 0, _
      UBound(varArray, 1), lngType, lngCaseSensitive

  If UBound(varArray, 1) > 0 And UBound(varArray, 2) > 0 Then
    For lngColIndex = 1 To UBound(varArray, 1)
      lngType = varTypes(lngColIndex)
      varRow = ReturnRow(varArray, 0)
      lngStart = 0
      For lngIndex = 1 To UBound(varArray, 2)
        If CheckIfRowDifferent(varArray, varRow, lngColIndex - 1, lngIndex, varTypes, lngCaseSensitive) Then

          QuickSort.VariantAscending_TwoDimensional varArray, lngStart, lngIndex - 1, lngColIndex, _
              UBound(varArray, 1), lngType, lngCaseSensitive

          lngStart = lngIndex
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

Public Sub DoubleAscending_TwoDimensional(narray() As Double, inLow As Long, inHi As Long, _
    lngSortColumn As Long, lngMaxColumNumber As Long)

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

Private Function DateToJulian(MyDate As Date) As Long

   DateToJulian = DateValue(MyDate)

End Function


