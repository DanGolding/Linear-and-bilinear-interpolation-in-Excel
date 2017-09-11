Attribute VB_Name = "Interpolation"
Option Explicit

Function Interp1(xRange As Variant, _
                 yRange As Variant, _
                 xVal As Double, _
                 Optional isSorted As Long = 1) As Double
    ' This function performs linear interpolation and was created by Dan Golding (https://github.com/DanGolding/Linear-and-bilinear-interpolation-in-Excel)
    ' It is a minor adaptation of: http://www.vbaexpress.com/forum/showthread.php?41522-Linear-Interpolation
    
    Dim yVal As Double
    Dim xBelow As Double, xAbove As Double
    Dim yBelow As Double, yAbove As Double
    Dim testVal As Double
    Dim High As Long, Med As Long, Low As Long
     
    Low = 1
    High = xRange.Cells.Count
     
    If isSorted <> 0 Then
        ' binary search sorted range
        Do
            Med = Int((Low + High) \ 2)
            If (xRange.Cells(Med).Value) < (xVal) Then
                Low = Med
            Else
                High = Med
            End If
        Loop Until Abs(High - Low) <= 1
    Else
        ' search every entry
        xBelow = -1E+205
        xAbove = 1E+205
         
        For Med = 1 To xRange.Cells.Count
            testVal = xRange.Cells(Med)
            If testVal < xVal Then
                If Abs(xVal - testVal) < Abs(xVal - xBelow) Then
                    Low = Med
                    xBelow = testVal
                End If
            Else
                If Abs(xVal - testVal) < Abs(xVal - xAbove) Then
                    High = Med
                    xAbove = testVal
                End If
            End If
        Next Med
    End If
     
    xBelow = xRange.Cells(Low): xAbove = xRange.Cells(High)
    yBelow = yRange.Cells(Low): yAbove = yRange.Cells(High)
     
    Interp1 = yBelow + (xVal - xBelow) * (yAbove - yBelow) / (xAbove - xBelow)
End Function

Public Function Interp2(xAxis As Range, yAxis As Range, zSurface As Range, xcoord As Double, ycoord As Double) As Double
' This function performs bilinear interpolation and was created by Dan Golding (https://github.com/DanGolding/Linear-and-bilinear-interpolation-in-Excel)
' It was adapted from http://www.quantcode.com/modules/mydownloads/singlefile.php?lid=416 which no longer appears to be online

    Dim xArr() As Variant
    xArr = xAxis.Value
    Dim yArr() As Variant
    yArr = yAxis.Value
    Dim zArr() As Variant
    zArr = zSurface.Value
    
    'first find 4 neighbouring points
    Dim nx As Long
    Dim ny As Long
    nx = UBound(xArr, 2)
    ny = UBound(yArr, 1)
    
    Dim lx As Single 'index of x coordinate of adjacent grid point to left of P
    Dim ux As Single 'index of x coordinate of adjacent grid point to right of P
    
    GetNeigbourIndices xArr, xcoord, lx, ux
    
    Dim ly As Single  'index of y coordinate of adjacent grid point below P
    Dim uy As Single  'index of y coordinate of adjacent grid point above P
    
    GetNeigbourIndices yArr, ycoord, ly, uy
    
    Dim fQ11, fQ21, fQ12, fQ22 As Double
    
    fQ11 = zArr(lx, ly)
    fQ21 = zArr(ux, ly)
    fQ12 = zArr(lx, uy)
    fQ22 = zArr(ux, uy)
    
    'if point exactly found on a node do not interpolate
    If ((lx = ux) And (ly = uy)) Then
        Interp2 = fQ11
        Exit Function
    End If
    
    Dim x, y, x1, x2, y1, y2 As Double
    
    x = xcoord
    y = ycoord
    
    x1 = xArr(lx, 1)
    x2 = xArr(ux, 1)
    y1 = yArr(ly, 1)
    y2 = yArr(uy, 1)
    
    'if xcoord lies exactly on an xAxis node do linear interpolation
    If (lx = ux) Then
        Interp2 = fQ11 + (fQ12 - fQ11) * (y - y1) / (y2 - y1)
        Exit Function
    End If
    
    'if ycoord lies exactly on an xAxis node do linear interpolation
    If (ly = uy) Then
        Interp2 = fQ11 + (fQ22 - fQ11) * (x - x1) / (x2 - x1)
        Exit Function
    End If
    
    Dim fxy As Double
    
    fxy = fQ11 * (x2 - x) * (y2 - y)
    fxy = fxy + fQ21 * (x - x1) * (y2 - y)
    fxy = fxy + fQ12 * (x2 - x) * (y - y1)
    fxy = fxy + fQ22 * (x - x1) * (y - y1)
    fxy = fxy / ((x2 - x1) * (y2 - y1))
    
    Interp2 = fxy
  
End Function

Public Sub GetNeigbourIndices(inArr As Variant, x As Double, ByRef lowerX As Single, ByRef upperX As Single)
' This function was created by Dan Golding (https://github.com/DanGolding/Linear-and-bilinear-interpolation-in-Excel)
' It is required for the Iterp2 function and was adapted from http://www.quantcode.com/modules/mydownloads/singlefile.php?lid=416 which no longer appears to be online

    Dim n As Long
    n = UBound(inArr, 1)
    
    If n = 1 Then
        'Transpose the arr
        inArr = Application.Transpose(inArr)
        n = UBound(inArr, 1)
    End If
    
    If x <= inArr(1, 1) Then
        lowerX = 1
        upperX = 1
    ElseIf x >= inArr(n, 1) Then
        lowerX = n
        upperX = n
    Else
        Dim I As Long
        For I = 2 To n
            If x < inArr(I, 1) Then
                lowerX = I - 1
                upperX = I
                Exit For
            ElseIf x = inArr(I, 1) Then
                lowerX = I
                upperX = I
                Exit For
            End If
        Next I
    End If
    
End Sub
