Attribute VB_Name = "Sort"
Private Sub test()
    
    Dim tempList
    Dim tempString As String
    
     Set tempList = CreateObject("system.collections.arraylist")
    
    Count = 346
    
    For i = 1 To Count
        tempString = Sheets(1).Cells(i, 1).Value
        tempList.Add tempString
    Next
    
    
    
        tempAry = tempList.toArray()
    
    Call QuickSort_top(tempAry, 0, UBound(tempAry))
    
    For i = 1 To Count
    
        Sheets(1).Cells(i, 2).Value = tempAry(i - 1)
    Next
    
    
    tempList.Clear
    
    
    
End Sub



Sub QuickSort_top(vArray As Variant, inLow As Long, inHi As Long)

  Dim tmpAry() As String
  
  
  tmpAry = vArray
  tempIndexAry = mkAry(UBound(vArray) + 1, 1)
  
 'Get max number
  tmpNum = 0
    
  For Each tempSubAry In vArray
    If getNum(tempSubAry)(0) > tmpNum Then
        tmpNum = getNum(tempSubAry)(0)
    End If
  Next
  
  'right(string(10,"0") & a,10)
  
  ReDim tmpAry(UBound(vArray))
  
  For i = 0 To UBound(vArray)
    
    splitStr = getStr(vArray(i))
    splitNum = getNum(vArray(i))(1)
    
    tmpAry(i) = splitStr & Right(String(tmpNum, "0") & splitNum, tmpNum)
   
  Next
  
  Call QuickSort(tmpAry, tempIndexAry, inLow, inHi)
  
  
  For i = 0 To UBound(tempIndexAry)
  
    tmpAry(i) = vArray(tempIndexAry(i) - 1)
    
  Next
  
  vArray = tmpAry
  
  
End Sub


Private Sub QuickSort(vArray As Variant, vIndex As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpSwap2 As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        tmpSwap2 = vIndex(tmpLow)
        
        vArray(tmpLow) = vArray(tmpHi)
        vIndex(tmpLow) = vIndex(tmpHi)
        
        vArray(tmpHi) = tmpSwap
        vIndex(tmpHi) = tmpSwap2
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, vIndex, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, vIndex, tmpLow, inHi
End Sub


Private Function getStr(tempString As Variant) As String
    Dim tempStr As String

    For i = 1 To Len(tempString)
    
        Select Case Asc(Mid(tempString, i, 1))
        Case 65 To 90
            tempStr = tempStr & Mid(tempString, i, 1)
        End Select
    
    Next
    
    getStr = tempStr
    
End Function

Private Function getNum(tempString As Variant) As Long()

    Dim tempNum As String
    Dim tempAry(1) As Long
    

    For i = 1 To Len(tempString)
    
        Select Case Asc(Mid(tempString, i, 1))
        Case 48 To 57
            tempNum = tempNum & Mid(tempString, i, 1)
        End Select
    
    Next
    tempAry(0) = Len("" & CLng(tempNum))
    tempAry(1) = CLng(tempNum)
    getNum = tempAry
    
End Function

Private Function mkAry(num As Integer, startNum As Integer) As Variant
    
    Dim tmpIndexAry() As Integer
    
     ReDim tmpIndexAry(num - 1)
     
     For i = 0 To UBound(tmpIndexAry)
        tmpIndexAry(i) = startNum + i
     Next
    
    
    mkAry = tmpIndexAry
    
End Function

