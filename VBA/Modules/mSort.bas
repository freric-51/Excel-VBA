Attribute VB_Name = "mSort"
Option Explicit
' This routine uses the "heap sort" algorithm to sort a VB collection.
' It returns the sorted collection.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function SortCollection(ByVal c As Collection) As Collection
   Dim n As Long: n = c.Count
   If n = 0 Then Set SortCollection = New Collection: Exit Function
   ReDim Index(0 To n - 1) As Long                    ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(i) = i + 1: Next         ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                     ' generate ordered heap
      Heapify c, Index, i, n
      Next
   For m = n To 2 Step -1                             ' sort the index array
      Exchange Index, 0, m - 1                        ' move highest element to top
      Heapify c, Index, 0, m - 1
      Next
   Dim c2 As New Collection
   For i = 0 To n - 1: c2.add c.Item(Index(i)): Next  ' fill output collection
   Set SortCollection = c2
   End Function

Private Sub Heapify(ByVal c As Collection, Index() As Long, ByVal i1 As Long, ByVal n As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   Dim nDiv2 As Long: nDiv2 = n \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim k As Long: k = 2 * i + 1
      If k + 1 < n Then
         If c.Item(Index(k)) < c.Item(Index(k + 1)) Then k = k + 1
         End If
      If c.Item(Index(i)) >= c.Item(Index(k)) Then Exit Do
      Exchange Index, i, k
      i = k
      Loop
   End Sub

Private Sub Exchange(Index() As Long, ByVal i As Long, ByVal j As Long)
   Dim temp As Long: temp = Index(i)
   Index(i) = Index(j)
   Index(j) = temp
End Sub
