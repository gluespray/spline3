Attribute VB_Name = "ModSpline"
Option Explicit

Private Const RowDataStart As Long = 2 'データの開始行
Private Const deltax As Double = 0.1 '補間計算のステップ

Public Sub スプライン()
    'データ点読み込み
    Dim x As Variant, y As Variant
    x = GetValSht("A")
    y = GetValSht("B")
    
    
    '係数の計算
    Dim a3 As Variant, a2 As Variant, a1 As Variant, a0 As Variant
    CalcSpline3 x, y, a3, a2, a1, a0
    
    '補間値の計算
    Dim xy As Dictionary  '計算した点の座標xy(j)=array(x,y)
    Set xy = CalcInterPolation(x, a3, a2, a1, a0)
    
    
    
    '係数の出力
    Dim i As Long
    Dim ii As Long
    Cells(1, "D") = "a3"
    Cells(1, "E") = "a2"
    Cells(1, "F") = "a1"
    Cells(1, "G") = "a0"
    For i = 0 To UBound(a3)
        ii = i + RowDataStart
        Cells(ii, "D") = a3(i)
        Cells(ii, "E") = a2(i)
        Cells(ii, "F") = a1(i)
        Cells(ii, "G") = a0(i)
    Next

    '補間値の出力
    Cells(1, "I") = "x"
    Cells(1, "J") = "y"
    Dim keywd As Variant
    i = RowDataStart
    For Each keywd In xy.keys
        Cells(i, "I") = CDbl(xy(keywd)(0))
        Cells(i, "J") = CDbl(xy(keywd)(1))
        i = i + 1
    Next

    Set xy = Nothing
End Sub

'スプラインの係数計算
Private Sub CalcSpline3(x As Variant, y As Variant, _
    ByRef a3 As Variant, ByRef a2 As Variant, _
    ByRef a1 As Variant, ByRef a0 As Variant)

    '点の番号
    Dim n As Long
    n = UBound(x) '終了点n。'開始点0。ポイント数はn+1
    
    '行列でつかう係数の計算
    Dim h As Variant 'xの差分　xi+1 - xi
    h = Seth(x)
    Dim Ho As Variant '2*(hi + hi+1)
    Ho = SetHo(h)
    Dim k As Variant ' yとhで求まる定数
    k = Setk(h, y)
    '2階微分係数を解く
    'Hmat * g2 = k
    Dim Hmat As Variant
    Hmat = SetHmat(n, Ho, h)
    SolveEquation Hmat, k
    Dim g2 As Variant '2次微分係数
    g2 = Setg2(n, k)
    
    a3 = Seta3(n, g2, h)
    a2 = Seta2(n, g2)
    a1 = Seta1(n, g2, h, y)
    a0 = Seta0(n, y)
    
End Sub

'各係数の計算
Private Function Seta3(n As Long, g2 As Variant, h As Variant)
    Dim a3 As Variant
    ReDim a3(n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a3(i) = (g2(i + 1) - g2(i)) / (6 * h(i))
    Next
    Seta3 = a3
End Function
Private Function Seta2(n As Long, g2 As Variant)
    Dim a2 As Variant
    ReDim a2(n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a2(i) = g2(i) / 2
    Next
    Seta2 = a2
End Function
Private Function Seta1(n As Long, g2 As Variant, h As Variant, y As Variant)
    Dim a1 As Variant
    ReDim a1(n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a1(i) = (y(i + 1) - y(i)) / h(i) - h(i) * (2 * g2(i) + g2(i + 1)) / 6
    Next
    Seta1 = a1
End Function
Private Function Seta0(n As Long, y As Variant)
    Dim a0 As Variant
    ReDim a0(n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a0(i) = y(i)
    Next
    Seta0 = a0
End Function
Private Function Setg2(n As Long, k As Variant) As Variant
    Dim g2 As Variant
    ReDim g2(n)
    Dim i As Long
    For i = 0 To n - 1
        g2(i) = k(i)
    Next
    g2(n) = 0
    Setg2 = g2
End Function
Private Function SetHmat(n As Long, Ho As Variant, h As Variant) As Variant
    Dim Hmat As Variant
    ReDim Hmat(1 To n - 1, 1 To n - 1)
    Dim i As Long
    For i = 1 To n - 1
        Hmat(i, i) = Ho(i - 1)
    Next
    For i = 1 To n - 2
        Hmat(i, i + 1) = h(i)
        Hmat(i + 1, i) = h(i)
    Next
    SetHmat = Hmat
End Function
Private Function Seth(x As Variant) As Variant
    Dim h As Variant
    ReDim h(UBound(x) - 1)
    Dim i As Long
    For i = LBound(x) To UBound(x) - 1
        h(i) = x(i + 1) - x(i)
    Next
    Seth = h
End Function
Private Function SetHo(h As Variant) As Variant
    Dim Ho As Variant
    ReDim Ho(UBound(h) - 1)
    Dim i As Long
    For i = LBound(h) To UBound(h) - 1
        Ho(i) = 2 * (h(i + 1) + h(i))
    Next
    SetHo = Ho
End Function
Private Function Setk(h As Variant, y As Variant) As Variant
    Dim k As Variant
    ReDim k(UBound(y) - 1)
    Dim i As Long
    For i = LBound(y) + 1 To UBound(y) - 1
        k(i) = 6 * (y(i + 1) - y(i)) / h(i) - 6 * (y(i) - y(i - 1)) / h(i - 1)
    Next
    Setk = k
End Function

'補間の計算
Private Function CalcInterPolation(x As Variant, a3 As Variant, a2 As Variant, a1 As Variant, a0 As Variant) As Dictionary
    Dim xy As New Dictionary
    Dim i As Long
    Dim j As Long 'xyのカウンター
    Dim xx As Double, yy As Double '計算する点のx,y座標
    Dim xxx As Double, dx As Double
    Dim xmin As Double, xmax As Double '計算する区間の始点終点
    Dim n As Long '終点
    n = UBound(x)
    j = 0
    dx = deltax
    For i = 0 To n - 1 'データ点の番号、区間
        xmin = x(i)
        xmax = x(i + 1)
        
        For xx = xmin To xmax - dx Step dx
            xxx = xx - xmin
            yy = a3(i) * xxx ^ 3 + a2(i) * xxx ^ 2 + a1(i) * xxx + a0(i)
            xy(j) = Array(xx, yy)
            j = j + 1
        Next
    Next
    Set CalcInterPolation = xy
End Function

'ワークシートの値を取得して変数に入れて返す
Private Function GetValSht(col As String) As Variant
    Dim RowEnd As Long
    RowEnd = Cells(Rows.Count, col).End(xlUp).Row
    Dim n As Long
    n = RowEnd - RowDataStart + 1
    Dim ret As Variant
    ReDim ret(0 To n - 1)
    Dim i As Long
    
    For i = 0 To n - 1
        ret(i) = Cells(i + RowDataStart, col)
    Next
    GetValSht = ret
End Function



'SolveEquationが1からと0からに対応しているか確認
Private Sub testSolveEquation()

    Const i0 = 23, j0 = 3, j1 = 6

    Dim aaa As Variant, bbb As Variant
    Dim n As Long
    n = 3
    Dim ista As Long, iend As Long
    ista = 0
    iend = n - 1
    ReDim aaa(ista To iend, ista To iend), bbb(ista To iend)
    'ReDim aaa(0 To n - 1, 0 To n - 1), bbb(0 To n - 1)
    
    Dim i As Long, j As Long
    For i = ista To iend
        For j = ista To iend
            aaa(i, j) = Cells(i0 + i - ista, j0 + j - ista)
        Next
        bbb(i) = Cells(i0 + i - ista, j1)
    Next
    
    

    SolveEquation aaa, bbb
    
End Sub


'連立一次方程式を解く。
'aaa * x = bbb を解く。答えx は bbb　に入る。元のaaa,bbbは壊れる。
'Numerical Recipes in C による。
Function SolveEquation(aaa As Variant, bbb As Variant)
    Dim ddd As Double, nnn As Long, indx() As Long
    nnn = UBound(bbb) - LBound(bbb) + 1 'データ数
    ReDim indx(LBound(bbb) To UBound(bbb))

    LUdecomp aaa, indx, ddd
    LUbksubs aaa, indx, bbb
    
    Erase indx
End Function

'LU分解。Decompression
Function LUdecomp(aaa As Variant, indx() As Long, ddd As Double)
    Dim i As Long, imax As Long, j As Long, k As Long
    Dim big As Double, dum As Double, sum As Double, temp As Double
    Dim vv() As Double
    Dim strmatrix As String
    Dim nmin As Long, nmax As Long, nnn As Long
    
    nmin = LBound(aaa, 1)
    nmax = UBound(aaa, 1)
    nnn = nmax - nmin + 1
    Const TINY = 1E-100
    ReDim vv(nmin To nmax)
    ddd = 1#
    For i = nmin To nmax
        big = 0#
        For j = nmin To nmax
            temp = Abs(aaa(i, j))
            If temp > big Then big = temp
        Next j
        If big = 0# Then big = TINY
        vv(i) = 1# / big
    Next i
    For j = nmin To nmax
        For i = nmin To j - 1
            sum = aaa(i, j)
            For k = nmin To i - 1
                sum = sum - aaa(i, k) * aaa(k, j)
            Next k
            aaa(i, j) = sum
        Next i
        big = 0#
        For i = j To nmax
            sum = aaa(i, j)
            For k = nmin To j - 1
                sum = sum - aaa(i, k) * aaa(k, j)
            Next k
            aaa(i, j) = sum
            dum = vv(i) + Abs(sum)
            If dum >= big Then
                big = dum
                imax = i
            End If
        Next i
        If j <> imax Then
            For k = nmin To nmax
                dum = aaa(imax, k)
                aaa(imax, k) = aaa(j, k)
                aaa(j, k) = dum
            Next k
            ddd = -ddd
            vv(imax) = vv(j)
        End If
        indx(j) = imax
        If aaa(j, j) = 0# Then aaa(j, j) = TINY
        If j <> nnn Then
            dum = 1# / aaa(j, j)
            For i = j + 1 To nmax
                aaa(i, j) = aaa(i, j) * dum
            Next i
        End If
    Next j
    Erase vv
    
    Exit Function
    
error100:
    strmatrix = "ludcmp 行" & i & "が全てゼロなので計算できない" & vbCrLf
    For i = nmin To nmax
        For j = nmin To nmax
            strmatrix = strmatrix & " " & aaa(i, j)
        Next j
        strmatrix = strmatrix & vbCrLf
    Next i
    
    Exit Function
End Function

'/** LU分解 Back Substitution */
Function LUbksubs(aaa As Variant, indx() As Long, bbb As Variant)
    Dim i As Long, ii As Long, ip As Long, j As Long, sum As Double
    Dim nmin As Long, nmax As Long, nnn As Long
    
    nmin = LBound(aaa, 1)
    nmax = UBound(aaa, 1)
    nnn = nmax - nmin + 1
    
    ii = nmin - 1 'ii = 0
    For i = nmin To nmax
        ip = indx(i)
        sum = bbb(ip)
        bbb(ip) = bbb(i)
        If ii <> nmin - 1 Then 'If ii <>  0 Then
            For j = ii To i - 1
                sum = sum - aaa(i, j) * bbb(j)
            Next j
        ElseIf sum <> nmin - 1 Then 'ElseIf sum <> 0 Then
            ii = i
        End If
        bbb(i) = sum
    Next i
    For i = nmax To nmin Step -1
        sum = bbb(i)
        For j = i + 1 To nmax 'For j = i + 1 To nnn
            sum = sum - aaa(i, j) * bbb(j)
        Next j
        bbb(i) = sum / aaa(i, i)
    Next i
End Function

'2022/06/14 14:36:12
'2022/06/15 11:20:54 スプライン完成
'2022/06/15 11:23:17
'2022/06/15 11:44:18
'2022/06/15 12:03:49
'2022/06/17 12:22:21
