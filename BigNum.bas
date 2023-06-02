Attribute VB_Name = "BigNumRoutines"
'Autor: Andrija Radovic
Type Result
    res As String
    rmm As String
End Type

Sub Main()
    BigNum.Show
End Sub

Sub SWAP(a As Variant, b As Variant)
    Dim c As Variant
    c = a
    a = b
    b = c
End Sub

Function NoLead0(ByVal a As String, ByVal d As String) As String
    Dim s As String
    a = Trim(a)
    s = Left$(a, 1)
    If s = "-" Then
        a = Mid$(a, 2)
    Else
        s = ""
    End If
    a = Replace$(LTrim$(Replace$(a, Left$(d, 1), " ")), " ", Left$(d, 1))
    NoLead0 = IIf(a = "", Left$(d, 1), s & a)
End Function

Function ABSS(ByVal a As String, d As String) As String
    a = NoLead0(a, d)
    ABSS = IIf(Left$(a, 1) = "-", Mid$(a, 2), a)
End Function

Function COMPS(ByVal b As String, ByVal c As String, d As String) As Integer
    Dim i As Long, j As Long, k As Long
    If Len(c) = Len(b) Then
        For i = 1 To Len(c)
            j = InStr(d, Mid$(b, i, 1))
            k = InStr(d, Mid$(c, i, 1))
            If j <> k Then
                If j < k Then
                    COMPS = -1
                    Exit Function
                Else
                    COMPS = 1
                    Exit Function
                End If
            End If
        Next
        COMPS = 0
    Else
        If Len(c) > Len(b) Then
            COMPS = -1
        Else
            COMPS = 1
        End If
    End If
End Function

Function ADDS(ByVal b As String, ByVal c As String, d As String) As String
    Dim n As Long, of As Long, i As Long, f As Long, f1 As Long, a As String, ppp As String
    n = Len(d)
    of = Len(c) - Len(b)
    If of < 0 Then
        SWAP b, c
        of = -of
    End If
    a = ""
    f = 0
    For i = Len(b) To 1 Step -1
        f1 = f + InStr(d, Mid$(b, i, 1)) + InStr(d, Mid$(c, of + i, 1)) - 2
        If f1 >= n Then
            f = 1
            f1 = f1 - n
        Else
            f = 0
        End If
        a = Mid$(d, 1 + f1, 1) & a
    Next
    If of Then
        For i = of To 1 Step -1
            If f Then
                f1 = f + InStr(d, Mid$(c, i, 1)) - 1
                If f1 >= n Then
                    f = 1
                    f1 = f1 - n
                Else
                    f = 0
                End If
                a = Mid$(d, 1 + f1, 1) & a
            Else
                a = Mid$(c, 1, i) & a
                Exit For
            End If
        Next
    End If
    If f Then ADDS = Mid$(d, 2, 1) & a Else ADDS = a
End Function

Function SUBS(ByVal b As String, ByVal c As String, d As String) As String
    Dim s As String, g As String, i As Long
    If COMPS(b, c, d) < 0 Then
        SWAP c, b
        g = "-"
    Else
        g = ""
    End If
    s = String$(Len(b) - Len(c), Right$(d, 1))
    For i = 1 To Len(c)
        s = s & Mid$(d, 1 + Len(d) - InStr(d, Mid$(c, i, 1)), 1)
    Next
    SUBS = g & NoLead0(Right$(ADDS(b, ADDS(s, Mid$(d, 2, 1), d), d), Len(b)), d)
End Function

Function MULS(ByVal b As String, ByVal c As String, d As String) As String
    Dim i As Long, j As Long, n As Long, f As Long, f1 As Long
    Dim m As Long, a As String, p As String, nul As String
    n = Len(d)
    If Len(b) > Len(c) Then SWAP b, c
    a = ""
    nul = ""
    For i = Len(b) To 1 Step -1
        m = InStr(d, Mid$(b, i, 1)) - 1
        p = ""
        f = 0
        For j = Len(c) To 1 Step -1
            f1 = f + m * (InStr(d, Mid$(c, j, 1)) - 1)
            If f1 >= n Then
                f = f1 \ n
                f1 = f1 Mod n
            Else
                f = 0
            End If
            p = Mid$(d, 1 + f1, 1) + p
        Next
        If f Then p = Mid$(d, 1 + f, 1) + p
        p = p + nul
        nul = nul + Left$(d, 1)
        a = ADDS(a, p, d)
    Next
    MULS = a
End Function

' The DIVS subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function DIVS(ByVal a As String, ByVal b As String, dg As String) As Result
    Dim d As String, c As String, p As String, Stack As New Collection
    If b = Left$(dg, 1) Then
        DIVS.res = 0
        DIVS.rmm = 0
    Else
        d = Mid$(dg, 2, 1)
        Do
            Stack.Add d
            Stack.Add b
            d = ADDS(d, d, dg)
            b = ADDS(b, b, dg)
        Loop While COMPS(b, a, dg) <= 0
        b = Stack.Item(Stack.Count): Stack.Remove Stack.Count
        b = SUBS(a, b, dg)
        p = Stack.Item(Stack.Count): Stack.Remove Stack.Count
        If p = Mid$(dg, 2, 1) Then
            If Left$(b, 1) = "-" Then
                p = Left$(dg, 1)
                b = a
            End If
        Else
            Do
                d = Stack.Item(Stack.Count): Stack.Remove Stack.Count
                c = Stack.Item(Stack.Count): Stack.Remove Stack.Count
                d = SUBS(d, b, dg)
                If d = Left$(dg, 1) Or Left$(d, 1) = "-" Then
                    b = IIf(Left$(d, 1) = "-", Mid$(d, 2), d)
                    p = ADDS(p, c, dg)
                End If
            Loop Until c = Mid$(dg, 2, 1)
        End If
        DIVS.res = p
        DIVS.rmm = b
    End If
End Function

Function POWS(ByVal c As String, ByVal b As String, d As String) As String
    Dim s As String, ppp As String
    Static ff As Object
    If TypeName(ff) <> "Dictionary" Then Set ff = CreateObject("Scripting.Dictionary")
    ppp = c & "|" & b & "|" & d
    If ff.Exists(ppp) Then
        POWS = ff(ppp)
    Else
        b = A2B(b, d, "01")
        s = Mid$(d, 2, 1)
        Do
            If Right$(b, 1) = "1" Then s = MULS(s, c, d)
            b = Left$(b, Len(b) - 1)
            If b <> "" Then c = MULS(c, c, d) Else Exit Do
        Loop
        ff.Add ppp, s
        POWS = s
    End If
End Function

Function XORS(ByVal b As String, ByVal c As String, d As String) As String
    Dim i As Long, s As String
    If Len(b) > Len(c) Then SWAP b, c
    If InStr("|1|2|4|8", "|" & Trim$(Replace$(Hex$(Len(d)), "0", " "))) Then
        s = ""
        b = StrReverse(b)
        c = StrReverse(c)
        For i = 1 To Len(b)
            s = s & Mid$(d, ((InStr(d, Mid$(b, i, 1)) - 1) Xor (InStr(d, Mid$(c, i, 1)) - 1)) + 1, 1)
        Next
        XORS = NoLead0(StrReverse(s & Mid$(c, Len(b) + 1)), d)
    Else
        b = StrReverse(A2B(b, d, "01234567"))
        c = StrReverse(A2B(c, d, "01234567"))
        s = ""
        For i = 1 To Len(b)
            s = s & Oct$(Val(Mid$(b, i, 1)) Xor Val(Mid$(c, i, 1)))
        Next
        XORS = A2B(StrReverse(s & Mid$(c, Len(b) + 1)), "01234567", d)
    End If
End Function

Function ORS(ByVal b As String, ByVal c As String, d As String) As String
    Dim i As Long, s As String
    If Len(b) > Len(c) Then SWAP b, c
    If InStr("|1|2|4|8", "|" & Trim$(Replace$(Hex$(Len(d)), "0", " "))) Then
        s = ""
        b = StrReverse(b)
        c = StrReverse(c)
        For i = 1 To Len(b)
            s = s & Mid$(d, ((InStr(d, Mid$(b, i, 1)) - 1) Or (InStr(d, Mid$(c, i, 1)) - 1)) + 1, 1)
        Next
        ORS = NoLead0(StrReverse(s & Mid$(c, Len(b) + 1)), d)
    Else
        b = StrReverse(A2B(b, d, "01234567"))
        c = StrReverse(A2B(c, d, "01234567"))
        s = ""
        For i = 1 To Len(b)
            s = s & Oct$(Val(Mid$(b, i, 1)) Or Val(Mid$(c, i, 1)))
        Next
        ORS = A2B(StrReverse(s & Mid$(c, Len(b) + 1)), "01234567", d)
    End If
End Function

Function ANDS(ByVal b As String, ByVal c As String, d As String) As String
    Dim i As Long, s As String
    If Len(b) > Len(c) Then SWAP b, c
    If InStr("|1|2|4|8", "|" & Trim$(Replace$(Hex$(Len(d)), "0", " "))) Then
        s = ""
        b = StrReverse(b)
        c = StrReverse(c)
        For i = 1 To Len(b)
            s = s & Mid$(d, ((InStr(d, Mid$(b, i, 1)) - 1) And (InStr(d, Mid$(c, i, 1)) - 1)) + 1, 1)
        Next
        ANDS = NoLead0(StrReverse(s), d)
    Else
        b = StrReverse(A2B(b, d, "01234567"))
        c = StrReverse(A2B(c, d, "01234567"))
        s = ""
        For i = 1 To Len(b)
            s = s & Oct$(Val(Mid$(b, i, 1)) And Val(Mid$(c, i, 1)))
        Next
        ANDS = A2B(StrReverse(s), "01234567", d)
    End If
End Function

Function NEGS(b As String, d As String) As String
    NEGS = A2B(Replace(Replace(Replace(A2B(b, d, "01"), "0", "X"), "1", "0"), "X", "1"), "01", d)
End Function

Function ChgSgn(a As String) As String
    ChgSgn = IIf(Left$(a, 1) = "-", Mid$(a, 2), "-" & a)
End Function

Function SUBA(ByVal b As String, ByVal c As String, d As String) As String
    Dim s As Integer
    b = NoLead0(b, d)
    c = NoLead0(c, d)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0
        SUBA = SUBS(b, c, d)
    Case 3
        SUBA = ChgSgn(SUBS(b, c, d))
    Case 1
        SUBA = "-" & ADDS(b, c, d)
    Case 2
        SUBA = ADDS(b, c, d)
    End Select
End Function

Function ORA(b As String, c As String, d As String) As String
    ORA = ORS(ABSS(b, d), ABSS(c, d), d)
End Function

Function ANDA(b As String, c As String, d As String) As String
    ANDA = ANDS(ABSS(b, d), ABSS(c, d), d)
End Function

Function XORA(ByVal b As String, ByVal c As String, d As String) As String
    XORA = XORS(ABSS(b, d), ABSS(c, d), d)
End Function

Function NEGA(b As String, d As String) As String
    NEGA = NEGS(ABSS(b, d), d)
End Function

Function ADDA(ByVal b As String, ByVal c As String, d As String) As String
    Dim s As Integer
    b = NoLead0(b, d)
    c = NoLead0(c, d)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0
        ADDA = ADDS(b, c, d)
    Case 3
        ADDA = "-" & ADDS(b, c, d)
    Case 1
        ADDA = SUBS(c, b, d)
    Case 2
        ADDA = SUBS(b, c, d)
    End Select
End Function

Function MULA(ByVal b As String, ByVal c As String, d As String) As String
    Dim s As Integer
    b = NoLead0(b, d)
    c = NoLead0(c, d)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0, 3
        MULA = MULS(b, c, d)
    Case 1, 2
        MULA = "-" & MULS(b, c, d)
    End Select
End Function

Function DIVA(ByVal b As String, ByVal c As String, d As String) As String
    Dim s As Integer
    b = NoLead0(b, d)
    c = NoLead0(c, d)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0, 3
        DIVA = DIVS(b, c, d).res
    Case 1, 2
        DIVA = "-" & DIVS(b, c, d).res
    End Select
End Function

' The POWA subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function POWA(ByVal b As String, ByVal c As String, d As String) As String
    Dim s As Integer
    b = NoLead0(b, d)
    c = NoLead0(c, d)
    c = IIf(Left$(c, 1) = "-", Mid$(c, 2), c)
    If Left$(b, 1) = "-" Then
        b = Mid$(b, 2)
        If InStr(d, Right$(c, 1)) And 1 Then
            POWA = POWS(b, c, d)
        Else
            POWA = "-" & POWS(b, c, d)
        End If
    Else
        POWA = POWS(b, c, d)
    End If
End Function

Function ChooseOp(a As String, b As String, c As String, d As String) As String
    Select Case a
    Case "+"
        ChooseOp = ADDA(b, c, d)
    Case "%"
        ChooseOp = ORA(b, c, d)
    Case "&"
        ChooseOp = ANDA(b, c, d)
    Case "@"
        ChooseOp = XORA(b, c, d)
    Case "-"
        ChooseOp = SUBA(b, c, d)
    Case "*"
        ChooseOp = MULA(b, c, d)
    Case "/"
        ChooseOp = DIVA(b, c, d)
    Case "^"
        ChooseOp = POWA(b, c, d)
    Case "~"
        ChooseOp = NEGA(c, d)
    Case "±"
        ChooseOp = ChgSgn(c)
    Case "!"
        ChooseOp = FACTA(b, d)
    End Select
End Function

' The EVAL subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function EVAL(ByVal a As String, l As String, d As String) As String
    Dim i As Long, j As Variant, b As String
    a = NoLead0(a, d)
    i = UBound(Split(a, "(")) - UBound(Split(a, ")"))
    If i > 0 Then a = a & String(i, ")")
    a = Replace(a, "#", l)
    i = InStr(a, "-")
    Do While i
        If InStr("(+-%&@/*^±~", Mid$(a, IIf(i = 1, 1, i - 1), 1)) Then Mid$(a, i, 1) = "±"
        i = InStr(i + 1, a, "-")
    Loop
    Do
        b = XB(a)
        If b = String$(Len(a), "X") Then
            a = Mid$(a, 2, Len(a) - 2)
        Else
            Exit Do
        End If
    Loop
    If b <> "Y" Then
        For Each j In Array("+", "-", "%", "&", "@", "/", "*", "^", "±", "~", "!")
            i = InStrRev(b, j)
            If i Then
                EVAL = ChooseOp(CStr(j), EVAL(Left$(a, i - 1), l, d), EVAL(Mid$(a, i + 1), l, d), d)
                Exit Function
            End If
        Next
        EVAL = NoLead0(a, d)
    Else
        EVAL = Left$(d, 1)
    End If
End Function

Function XB(a As String) As String
    Dim i As Long, j As Long
    XB = a
    If InStr(a, "(") Then
        j = 0
        For i = 1 To Len(a)
            Select Case Mid$(a, i, 1)
            Case "("
                j = j + 1
                If j > 0 Then Mid(XB, i, 1) = "X"
            Case ")"
                If j > 0 Then Mid(XB, i, 1) = "X"
                j = j - 1
            Case Else
                If j > 0 Then Mid(XB, i, 1) = "X"
            End Select
        Next
    End If
    If j Then XB = "Y"
End Function

Function A2BS(ByVal a As String, da As String, db As String) As String
    Dim sg As String, s As String, m As String, i As Long, ppp As String
    ReDim n(0 To Len(da)) As String
    Static ff As Object
    If TypeName(ff) <> "Dictionary" Then Set ff = CreateObject("Scripting.Dictionary")
    ppp = a & "|" & da & "|" & db
    If ff.Exists(ppp) Then
        A2BS = ff(ppp)
    Else
        A2BS = A2B(a, da, db)
        ff.Add ppp, A2BS
    End If
End Function

Function A2B(ByVal a As String, da As String, db As String) As String
    Dim sg As String, s As String, m As String, i As Long, ppp As String
    ReDim n(0 To Len(da)) As String
    If Left$(a, 1) = "-" Then
        sg = "-"
        a = Mid$(a, 2)
    Else
        sg = ""
    End If
    a = NoLead0(UCase$(a), da)
    m = Mid$(db, 2, 1)
    s = Left$(db, 1)
    n(0) = s
    For i = 1 To Len(da)
        n(i) = ADDS(m, n(i - 1), db)
    Next
    For i = Len(a) To 2 Step -1
        s = ADDS(s, MULS(n(InStr(da, Mid$(a, i, 1)) - 1), m, db), db)
        m = MULS(n(Len(da)), m, db)
    Next
    A2B = sg & ADDS(s, MULS(n(InStr(da, Mid$(a, i, 1)) - 1), m, db), db)
End Function

Function ValS(c As String, d As String) As Long
    Dim b As Long, l As Long, i As Long
    Select Case d
    Case "0123456789"
        ValS = CLng(c)
    Case "01234567"
        ValS = Val("&O" & c)
    Case "0123456789ABCDEF"
        ValS = Val("&H" & c)
    Case Else
        ValS = 0
        b = 1
        l = Len(d)
        For i = Len(c) To 1 Step -1
            ValS = ValS + b * (InStr(d, Mid$(c, i, 1)) - 1)
            b = b * l
        Next
    End Select
End Function

Function StrS(ByVal a As Long, d As String) As String
    Dim l As Long
    Select Case d
    Case "0123456789"
        StrS = CStr(a)
    Case "01234567"
        StrS = Oct$(a)
    Case "0123456789ABCDEF"
        StrS = Hex$(a)
    Case Else
        StrS = ""
        l = Len(d)
        Do While a
            StrS = Mid$(d, 1 + (a Mod l), 1) & StrS
            a = a \ l
        Loop
    End Select
End Function

' The FACTS subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function FACTS(aa As String, d As String) As String
    Static ff As Object
    Dim fa As Variant, a As Long, b As Long, c As Long, i As Long, dd As Long
    Dim f As String, cc As String
    fa = Array(3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, 103, _
    107, 109, 113, 127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, 199, _
    211, 223, 227, 229, 233, 239, 241, 251, 257, 263, 269, 271, 277, 281, 283, 293, 307, 311, 313, _
    317, 331, 337, 347, 349, 353, 359, 367, 373, 379, 383, 389, 397, 401, 409, 419, 421, 431, 433, _
    439, 443, 449, 457, 461, 463, 467, 479, 487, 491, 499, 503, 509, 521, 523, 541, 547, 557, 563, _
    569, 571, 577, 587, 593, 599, 601, 607, 613, 617, 619, 631, 641, 643, 647, 653, 659, 661, 673, _
    677, 683, 691, 701, 709, 719, 727, 733, 739, 743, 751, 757, 761, 769, 773, 787, 797, 809, 811, _
    821, 823, 827, 829, 839, 853, 857, 859, 863, 877, 881, 883, 887, 907, 911, 919, 929, 937, 941, _
    947, 953, 967, 971, 977, 983, 991, 997, 1009, 1013, 1019, 1021, 1031, 1033, 1039, 1049, 1051, _
    1061, 1063, 1069, 1087, 1091, 1093, 1097, 1103, 1109, 1117, 1123, 1129, 1151, 1153, 1163, _
    1171, 1181, 1187, 1193, 1201, 1213, 1217, 1223, 1229, 1231, 1237, 1249, 1259, 1277, 1279, _
    1283, 1289, 1291, 1297, 1301, 1303, 1307, 1319, 1321, 1327, 1361, 1367, 1373, 1381, 1399, _
    1409, 1423, 1427, 1429, 1433, 1439, 1447, 1451, 1453, 1459, 1471, 1481, 1483, 1487, 1489, _
    1493, 1499, 1511, 1523, 1531, 1543, 1549, 1553, 1559, 1567, 1571, 1579, 1583, 1597, 1601, _
    1607, 1609, 1613, 1619, 1621, 1627, 1637, 1657, 1663, 1667, 1669, 1693, 1697, 1699, 1709, _
    1721, 1723, 1733, 1741, 1747, 1753, 1759, 1777, 1783, 1787, 1789, 1801, 1811, 1823, 1831, _
    1847, 1861, 1867, 1871, 1873, 1877, 1879, 1889, 1901, 1907, 1913, 1931, 1933, 1949, 1951, _
    1973, 1979, 1987, 1993, 1997, 1999, 2003, 2011, 2017, 2027, 2029, 2039, 2053, 2063, 2069)
    If TypeName(ff) <> "Dictionary" Then Set ff = CreateObject("Scripting.Dictionary")
    If ff.Exists(aa & "|" & d) Then
        FACTS = ff(aa & "|" & d)
    Else
        a = ValS(aa, d)
        f = "1"
        cc = "2"
        c = 2
        i = 0
        Do
            b = a
            dd = 0
            Do
                b = b \ c
                dd = dd + b
            Loop While b > 1
            f = MULS$(f, POWS$(cc, StrS$(dd, d), d), d)
            c = fa(i)
            cc = StrS$(c, d)
            i = i + 1
        Loop Until c > a
        ff.Add aa & "|" & d, f
        FACTS = f
    End If
End Function

Function FACTA(ByVal a As String, d As String) As String
    a = NoLead0(a, d)
    FACTA = FACTS(IIf(Left$(a, 1) = "-", Mid$(a, 2), a), d)
End Function

Function Sintax(a As String, b As String, c As String) As Boolean
    Select Case a
    Case "~", "-"
        Select Case b
        Case "+", "%", "&", "@", "/", "*", "^", "!", "~", ")", "-"
            Sintax = False
        Case Else
            Sintax = True
        End Select
    Case "("
        Select Case b
        Case "+", "%", "&", "@", "/", "*", "^", "!", ")"
            Sintax = False
        Case Else
            Sintax = True
        End Select
    Case "+", "%", "&", "@", "/", "*", "^"
        Select Case b
        Case "+", "%", "&", "@", "/", "*", "^", "!", ")"
            Sintax = False
        Case Else
            Sintax = True
        End Select
    Case "!"
        Sintax = IIf(b = "!", False, True)
    Case Else
        Select Case b
        Case "~"
            Sintax = False
        Case ")"
            If UBound(Split(c, "(")) > UBound(Split(c, ")")) Then
                Sintax = True
            Else
                Sintax = False
            End If
        Case Else
            Sintax = True
        End Select
    End Select
End Function
