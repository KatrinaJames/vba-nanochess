Attribute VB_Name = "Chess"
Option Explicit

' Code adapted from https://nanochess.org/chess4.html

Public gg As Integer
Public l(68) As Integer
Public bb As Integer
Public i As Integer
Public y As Integer
Public u As Integer
Public ii(411) As Integer
Public b As Integer

Public Const x As Integer = 10
Public Const z As Integer = 15
Public Const mm As Double = 10000

Public pieces(0 To 14, 0 To 1) As Byte

Public Const SHEET_NAME As String = "Chess"


Public Function ToHex(ByVal str As String) As Byte()
    Dim arrBytes() As Byte
    arrBytes = str
    ToHex = arrBytes
End Function

Public Function FromHex(ByRef arrBytes() As Byte) As String
    Dim str As String
    str = arrBytes
    FromHex = str
End Function


Public Sub NewGame()

    Dim blankSpace(1) As Byte: blankSpace(0) = 0: blankSpace(1) = 0

    Dim blackRook(1) As Byte: blackRook(0) = 92: blackRook(1) = 38
    Dim blackKnight(1) As Byte: blackKnight(0) = 94: blackKnight(1) = 38
    Dim blackBishop(1) As Byte: blackBishop(0) = 93: blackBishop(1) = 38
    Dim blackQueen(1) As Byte: blackQueen(0) = 91: blackQueen(1) = 38
    Dim blackKing(1) As Byte: blackKing(0) = 90: blackKing(1) = 38
    Dim blackPawn(1) As Byte: blackPawn(0) = 95: blackPawn(1) = 38
    
    Dim whiteRook(1) As Byte: whiteRook(0) = 86: whiteRook(1) = 38
    Dim whiteKnight(1) As Byte: whiteKnight(0) = 88: whiteKnight(1) = 38
    Dim whiteBishop(1) As Byte: whiteBishop(0) = 87: whiteBishop(1) = 38
    Dim whiteQueen(1) As Byte: whiteQueen(0) = 85: whiteQueen(1) = 38
    Dim whiteKing(1) As Byte: whiteKing(0) = 84: whiteKing(1) = 38
    Dim whitePawn(1) As Byte: whitePawn(0) = 89: whitePawn(1) = 38
    
    pieces(0, 0) = blankSpace(0)
    pieces(0, 1) = blankSpace(1)
    
    pieces(1, 0) = blackPawn(0)
    pieces(1, 1) = blackPawn(1)
    pieces(2, 0) = blackKing(0)
    pieces(2, 1) = blackKing(1)
    pieces(3, 0) = blackKnight(0)
    pieces(3, 1) = blackKnight(1)
    pieces(4, 0) = blackBishop(0)
    pieces(4, 1) = blackBishop(1)
    pieces(5, 0) = blackRook(0)
    pieces(5, 1) = blackRook(1)
    pieces(6, 0) = blackQueen(0)
    pieces(6, 1) = blackQueen(1)
    
    pieces(7, 0) = blankSpace(0)
    pieces(7, 1) = blankSpace(1)
    pieces(8, 0) = blankSpace(0)
    pieces(8, 1) = blankSpace(1)
    
    pieces(9, 0) = whitePawn(0)
    pieces(9, 1) = whitePawn(1)
    pieces(10, 0) = whiteKing(0)
    pieces(10, 1) = whiteKing(1)
    pieces(11, 0) = whiteKnight(0)
    pieces(11, 1) = whiteKnight(1)
    pieces(12, 0) = whiteBishop(0)
    pieces(12, 1) = whiteBishop(1)
    pieces(13, 0) = whiteRook(0)
    pieces(13, 1) = whiteRook(1)
    pieces(14, 0) = whiteQueen(0)
    pieces(14, 1) = whiteQueen(1)
    
    gg = 120
    
    l(0) = 5: l(1) = 3: l(2) = 4: l(3) = 6: l(4) = 2: l(5) = 4: l(6) = 3: l(7) = 5: l(8) = 1: l(9) = 1: l(10) = 1: l(11) = 1: l(12) = 1: l(13) = 1: l(14) = 1: l(15) = 1
    l(16) = 9: l(17) = 9: l(18) = 9: l(19) = 9: l(20) = 9: l(21) = 9: l(22) = 9: l(23) = 9: l(24) = 13: l(25) = 11: l(26) = 12: l(27) = 14: l(28) = 10: l(29) = 12: l(30) = 11: l(31) = 13
    l(32) = 0: l(33) = 99: l(34) = 0: l(35) = 306: l(36) = 297: l(37) = 495: l(38) = 846
    l(39) = -1: l(40) = 0: l(41) = 1: l(42) = 2: l(43) = 2: l(44) = 1: l(45) = 0: l(46) = -1
    l(47) = -1: l(48) = 1: l(49) = -10: l(50) = 10: l(51) = -11: l(52) = -9: l(53) = 9: l(54) = 11
    l(55) = 10: l(56) = 20: l(57) = -9: l(58) = -11: l(59) = -10: l(60) = -20
    l(61) = -21: l(62) = -19: l(63) = -12: l(64) = -8: l(65) = 8: l(66) = 12: l(67) = 19: l(68) = 21
    
    u = 0
    y = 0
    i = 0
    bb = 1
    
    Do While bb <= 120
        If -CBool(bb Mod x) Then
            If -CBool((((bb / x) * 100) Mod (x * 100)) / 100 < 2 Or (bb Mod x) < 2) Then
                ii(bb - 1) = 7
            Else
                If -CBool(Int(bb / x) And 4) Then
                    ii(bb - 1) = 0
                Else
                    ii(bb - 1) = l(i) Or 16
                    i = i + 1
                End If
            End If
        Else
            ii(bb - 1) = 7
        End If
        bb = bb + 1
    Loop
    Call ww
    
End Sub

Public Sub yy(ByVal s As Integer)

    Dim dummy As Long
    Dim difficulty As Integer

    i = (ii(s) Xor y) And z
    If i > 8 Then
        b = s
        Call ww
    Else
        If -CBool(bb) And i < 9 Then
            b = s
            i = ii(bb) And z
            If ((i And 7) = 1 And (b < 29 Or b > 90)) Then
                Dim index As Integer: index = 0
                Dim cell As Range
                
                With ThisWorkbook
                    Dim optionsRange As Range: Set optionsRange = .Sheets("hidden").Range("OPTIONS")
                    Dim pawnPromotion As Range: Set pawnPromotion = .Sheets(SHEET_NAME).Range("UPGRADE")
                End With
                
                If pawnPromotion.value = "" Then
                    index = 1
                Else
                    For Each cell In optionsRange
                        index = index + 1
                        If cell.value = pawnPromotion.value Then
                            GoTo BREAK
                        End If
                    Next cell
                End If
BREAK:
                i = 14 - (index - 1) Xor y
            End If
            dummy = xx(0, 0, 0, 21, u, 1)
            
            With ThisWorkbook
                Dim difficultyRange As Range: Set difficultyRange = .Sheets(SHEET_NAME).Range("DIFFICULTY")
                Dim difficultiesRange As Range: Set difficultiesRange = .Sheets("hidden").Range("DIFFICULTIES")
            End With
            
            difficulty = 1
            If difficultyRange.value <> "" Then
                For Each cell In difficultiesRange
                    If cell.value = difficultyRange.value Then
                        GoTo END_FOR
                    Else
                        difficulty = difficulty + 1
                    End If
                Next cell
            End If
END_FOR:
            
            If y > 0 Then
                dummy = xx(0, 0, 0, 21, u, (difficulty + 1))
                dummy = xx(0, 0, 0, 21, u, 1)
            End If
        End If
    End If

End Sub

Public Sub ww()

    Dim q As Range
    Dim p As Integer
    Dim piece(1) As Byte
    
    bb = b
    For p = 21 To 98 Step 1
        If RangeExists("o_" & p) Then
            Set q = ThisWorkbook.Sheets(SHEET_NAME).Range("o_" & p)
            piece(0) = pieces(ii(p) And z, 0)
            piece(1) = pieces(ii(p) And z, 1)
            q.value = FromHex(piece)
            If p = bb Then
                q.Borders.Color = vbRed
                q.Borders.LineStyle = xlContinuous
                q.Borders.Weight = xlMedium
            Else
                q.Borders.LineStyle = xlLineStyleNone
            End If
        End If
    Next p

End Sub

Private Function RangeExists(ByVal rangeName As String) As Boolean

    On Error GoTo RANGE_ERROR
    Dim testRange As Range
    Set testRange = ThisWorkbook.Sheets(SHEET_NAME).Range(rangeName)
    RangeExists = True
    On Error GoTo 0
    Err.Clear
    Exit Function
    
RANGE_ERROR:
    RangeExists = False
    On Error GoTo 0
    Err.Clear

End Function


Public Function xx(ByVal w As Integer, _
    ByVal c As Long, _
    ByVal h As Integer, _
    ByVal e As Integer, _
    ByVal ss As Integer, _
    ByVal s As Integer) As Long
    
    Dim a As Integer
    Dim cc As Long
    Dim r As Integer
    Dim q As Integer
    Dim aa As Long
    Dim m As Integer
    Dim n As Integer
    Dim g As Integer
    Dim p As Integer
    Dim kk As Double
    Dim nn As Double
    Dim oo As Integer
    Dim d As Integer ' king in check flag
    Dim ee As Integer
    Dim ll As Long
    Dim o As Integer
    Dim t As Integer
    Dim loopCondition As Long
    
    Dim jj As Integer
    
    Dim exitCondition As Boolean
    
    oo = e
    nn = -mm * mm
    kk = LeftShift(78 - h, x)
    a = IIf(CBool(y), -x, x)
    y = y Xor 8
    gg = gg + 1
    
    If CBool(w) Then
        d = 1
    ElseIf Not CBool(w) And Not CBool(s) Then
        d = 0
    ElseIf Not CBool(w) And Not s >= h Then
        d = 0
    Else
        d = -CBool(w) Or -CBool(s) And -CBool(s >= h) And -CBool(xx(0, 0, 0, 21, 0, 0) > mm)
    End If
    
    Do
        p = oo
        o = ii(p)
        If CBool(o) Then
            q = o And z Xor y
            If q < 7 Then
                aa = IIf(CBool(q And 2), 8, 4)
                q = q - 1
                
                Dim tempArray(6) As Integer
                tempArray(0) = 53
                tempArray(1) = 47
                tempArray(2) = 61
                tempArray(3) = 51
                tempArray(4) = 47
                tempArray(5) = 47
                tempArray(6) = 0
                
                cc = IIf(CBool(o - 9 And z), tempArray(q), 57)
                Do
                    p = p + l(cc)
                    r = ii(p)
                    
                    If Not CBool(w) Or (p = w) Then
                        g = IIf(q Or p + a - ss, 0, ss)
                        If Not CBool(r) And (-CBool(q) Or aa < 3 Or -CBool(g)) Or ((r + 1 And z Xor y) > 9) And (q Or -CInt(aa > 2)) Then
                            m = -(Not CBool(r - 2 And 7))
                            If m Then
                                y = y Xor 8
                                ii(gg) = oo
                                gg = gg - 1
                                xx = kk
                                Exit Function
                            End If
                            n = o And z
                            jj = n
                            ee = ii(p - a) And z
                            If CBool(q Or ee - 7) Then
                                t = n
                            Else
                                n = n + 2
                                t = 6 Xor y
                            End If
                            
                            Do While n <= t
                                ll = IIf(CBool(r), l(r And 7 Or 32) - h - q, 0)
                                If CBool(s) Then
                                    ll = ll + IIf(CBool(1 - q), l((p - p Mod x) / x + 37) - l((oo - oo Mod x) / x + 37) _
                                        + l(p Mod x + 38) * IIf(CBool(q), 1, 2) - l(oo Mod x + 38) _
                                        + (o And 16) / 2, -CBool(m) * 9) _
                                        + IIf(Not CBool(q), -(Not CBool((ii(p - 1) Xor n))) + -(Not CBool(ii(p + 1) Xor n)) + l(n And 7 Or 32) - 99 + -CBool(g) * 99 + -CBool(aa < 2), 0) _
                                        + -(Not CBool(ee Xor y Xor 9))
                                End If
                                
                                If (s > h) Or -(1 < s) And (s = h) And (ll > (z Or d)) Then
                                    ii(p) = n
                                    If CBool(m) Then
                                        ii(g) = ii(m)
                                        ii(m) = 0
                                    Else
                                        If CBool(g) Then
                                            ii(g) = 0
                                        End If
                                    End If
                                    ii(oo) = 0
                                    
                                    jj = IIf((q Or aa) > 1, 0, p)
                                    ll = ll - xx(IIf(s > (h Or d), 0, p), ll - nn, h + 1, ii(gg + 1), IIf((q Or aa) > 1, 0, p), s)
                                    
                                    If Not CBool(-CBool(h) Or s - 1 Or bb - oo Or i - n Or p - b Or ll < -mm) Then
                                        Call ww
                                        gg = gg - 1
                                        u = jj
                                        xx = u
                                        Exit Function
                                    End If
                                    
                                    If q - 1 Or aa < 7 Then
                                        jj = q - 1 Or -CBool(aa < 7)
                                    ElseIf CBool(m) Then
                                        jj = m
                                    ElseIf s < 1 Or d Or r > 0 Or o < z Then
                                        jj = s < 1 Or d Or r > 0 Or o < z
                                    Else
                                        jj = -CBool(xx(0, 0, 0, 21, 0, 0) > mm)
                                    End If
                                    
                                    ii(oo) = o
                                    ii(p) = r
                                    If CBool(m) Then
                                        ii(m) = ii(g)
                                        ii(g) = 0
                                    ElseIf CBool(g) Then
                                        ii(g) = 9 Xor y
                                    End If
                                    
                                End If
                                If -(ll > nn) Or -(s > 1) And -(ll = nn) And -(Not CBool(h)) And -(Rnd() < 0.5) Then
                                    ii(gg) = oo
                                    If s > 1 Then
                                        If CBool(h) And c - ll < 0 Then
                                            y = y Xor 8
                                            gg = gg - 1
                                            xx = ll
                                            Exit Function
                                        End If
                                        
                                        If Not CBool(h) Then
                                            i = n
                                            bb = oo
                                            b = p
                                        End If
                                    End If
                                    nn = ll
                                End If
                                
                                If CBool(jj) Then
                                    n = n + 1
                                Else
                                    g = p
                                    m = IIf(p < oo, g - 3, g + 2)
                                    If ii(m) < z Or ii(m + oo - p) Or ii(p + p - oo) Then
                                        n = n + 1
                                    End If
                                    p = p + p - oo
                                End If
                                
                            Loop
                        End If
                    End If
                    
                    If Not CBool(r) And (q > 2) Then
                        loopCondition = True
                    Else
                        p = oo
                        If CBool(q Or -(aa > 2) Or (o > z) And -(Not CBool(r))) Then
                            cc = cc + 1
                            loopCondition = CBool(Not CBool(r) And (q > 2)) Or _
                                CBool(q Or -(aa > 2) Or (o > z) And -(Not CBool(r))) And _
                                CBool(cc * (aa - 1))
                            aa = aa - 1
                        Else
                            loopCondition = False
                        End If
                    End If
                Loop Until Not CBool(loopCondition)
        
            End If
        End If
        
        oo = oo + 1
        If oo > 98 Then
            oo = 20
            exitCondition = True
        Else
            exitCondition = CBool(e - oo)
        End If
    Loop Until Not exitCondition
    
    y = y Xor 8
    gg = gg - 1
    xx = IIf(CBool(nn + mm * mm) And CBool(nn > -kk + 1924 Or d), nn, 0)

End Function

Public Function RightShift(ByVal value As Long, ByVal Shift As Byte) As Long
    
    RightShift = value
    
    If Shift > 0 Then
        RightShift = Int(RightShift / (2 ^ Shift))
    End If
    
End Function

Public Function LeftShift(ByVal value As Long, ByVal Shift As Integer) As Long
    LeftShift = value
    
    If Shift = -1 Then
        If value Mod 2 = 1 Then
            LeftShift = -2147483648#
            Exit Function
        Else
            LeftShift = 0
            Exit Function
        End If
    End If
    
    If Shift > 0 Then
        LeftShift = value * (CLng(2) ^ Shift)
    End If
End Function
