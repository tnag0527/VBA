Attribute VB_Name = "m_QuickSort"
Option Explicit


'---------------------------------------------------------------------
'◇クイックソート
'---------------------------------------------------------------------
Sub quicksort(ByRef vArr, Optional ByVal iL As Long = 0, Optional ByVal iR As Long = -1, Optional ByRef vID, Optional bDebugPrint As Boolean = False)
    Dim s As String, i As Long, b As Boolean, v, w
   
  '初期化
    If (iR < 0) Then iR = UBound(vArr)
    '
    b = False
    b = b Or IsMissing(vID)
    b = b Or IsEmpty(vID)
    If b Then
        vID = vArr
        For i = LBound(vID) To UBound(vID)
            vID(i) = i
        Next
    End If
    '
    Dim l As Long:    l = iL    '左側位置フラグ
    Dim r As Long:    r = iR    '右側位置フラグ
    Dim p As Long:    p = l     '基準(Pivot)
   
 
    If bDebugPrint Then Debug.Print
    If bDebugPrint Then dpArray vArr, iL, iR, p, l, r
   
    '「v(l) <= v(p) <= v(r)」が正常な並び順
    'p = l からスタート
    Do While l < r
       
        'rで調べる
        Do While p < r
       
            If vArr(p) <= vArr(r) Then
              '⇒位置と大小関係が正常 (→rを1つ左へ)
                r = r - 1
            Else
              '⇒位置と大小関係が違う (→値交換)(→pをrに移動)(→r終了)
                Call swap(vArr, r, p)
                Call swap(vID, r, p)
                p = r
                If bDebugPrint Then dpArray vArr, iL, iR, p, l, r, False: Debug.Print " <R "
                Exit Do
            End If
        DoEvents
        Loop
       
       
        'lで調べる
        Do While l < p
         
            If vArr(l) <= vArr(p) Then
                '⇒位置と大小関係が正常 (→lを1つ右へ)
                l = l + 1
            Else
                '⇒位置と大小関係が違う (→値交換)(→pをlに移動)(→l終了)
                Call swap(vArr, l, p)
                Call swap(vID, l, p)
                p = l
                If bDebugPrint Then dpArray vArr, iL, iR, p, l, r, False: Debug.Print "  L>"
                Exit Do
            End If
        DoEvents
        Loop
       
    DoEvents
    Loop
   
    If (iL < p - 1) Then quicksort vArr, iL, p - 1, vID, bDebugPrint
    If (p + 1 < iR) Then quicksort vArr, p + 1, iR, vID, bDebugPrint
   
 
End Sub
'----  ----  ----  ----  ----  ----  ----  ----  ----  ----
Private Sub TEST_quicksort_1()
    Dim i As Long, v
   
    v = Array(11, 12, 13, 14, 15, 16, 17, 18, 19)
    v = Array(19, 11, 12, 13, 14, 15, 16, 17, 18)
    v = Array(14, 13, 19, 15, 11, 17, 18, 12, 16)
   
    dpArr v
    quicksort v, , , , True 'メイン
    dpArr v
 
End Sub
'----  ----  ----  ----  ----  ----  ----  ----  ----  ----
Private Sub TEST_quicksort_2()
    Dim i As Long, v, w, x, vID
   
    '2桁×61個のサンプルを生成
    ReDim v(60)
    Randomize
    For i = LBound(v) To UBound(v)
        v(i) = Int(Rnd * 89) + 10
    Next
   
    x = v           'データ保持
    vID = Empty     '初期化
   
    dpArr v
    quicksort v, , , vID, True
    dpArr v
    dpArr vID
   
    For i = LBound(vID) To UBound(vID)
        Debug.Print x(vID(i));
    Next
    Debug.Print ""
 
 
End Sub
'----  ----  ----  ----  ----  ----  ----  ----  ----  ----
Private Sub TEST_quicksort_3()
    Dim i As Long, j As Long, v, w, x, y, vID
   
    w = Array( _
          Array(18, 13, 19, 15, 11, 17, 14, 12, 16) _
        , Array(27, 21, 26, 29, 22, 24, 28, 23, 25) _
      )
   
    ReDim x(40)
    ReDim y(40)
    Randomize
    For i = LBound(x) To UBound(x)
        x(i) = Int(Rnd * 89) + 10
        y(i) = Int(Rnd * 89) + 10
    Next
    w = Array(x, y)
   
    v = w(1)
    vID = Empty
   
    dpArr v
    quicksort v, , , vID, True
    dpArr v
    dpArr vID
   
    For j = LBound(w) To UBound(w)
    For i = LBound(vID) To UBound(vID)
        Debug.Print w(j)(vID(i));
    Next
        Debug.Print
    Next
    Debug.Print ""
 
 
End Sub
 
 
 
 
'---------------------------------------------------------------------
Sub swap(ByRef vArr, a As Long, b As Long)
    Dim v
    v = vArr(a)
    vArr(a) = vArr(b)
    vArr(b) = v
End Sub
 
 
 
'---------------------------------------------------------------------
'◇
'---------------------------------------------------------------------
Sub dpArr(v)
    Dim i As Long
   
    If IsMissing(v) Then Exit Sub
    If IsEmpty(v) Then Exit Sub
   
    
    For i = LBound(v) To UBound(v)
        Debug.Print v(i);
    Next
    Debug.Print " dpArr"
End Sub
 
'---------------------------------------------------------------------
'◇
'---------------------------------------------------------------------
Sub dpArray(v, Optional vL As Long = 0, Optional vR As Long = -1, Optional p As Long = -1, Optional l As Long = -1, Optional r As Long = -1, Optional bEndCR As Boolean = True)
    Dim i As Long
 
    If (vR < 0) Then vR = UBound(v)
 
    For i = LBound(v) To UBound(v)
        If (vL <= i) And (i <= vR) Then
            If i = p Then
                Debug.Print "[" & v(i) & "]";
            ElseIf i = l Then
                Debug.Print " " & v(i) & ">";
            ElseIf i = r Then
                Debug.Print "<" & v(i) & " ";
            Else
                Debug.Print v(i);
            End If
        Else
            Debug.Print " .. ";
        End If
    Next
    If bEndCR Then Debug.Print
End Sub
 
 


