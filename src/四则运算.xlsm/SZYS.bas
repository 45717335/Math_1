Attribute VB_Name = "SZYS"
Option Explicit

Function sizeyunsuan(Optional op1 As String = "", Optional op2 As String = "") As String


    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer, i7 As Integer, i8 As Integer, i9 As Integer, i10 As Integer
    Dim j1 As Integer, j2 As Integer, j3 As Integer, j4 As Integer, j5 As Integer, j6 As Integer, j7 As Integer, j8 As Integer, j9 As Integer, j10 As Integer
    Dim i As Integer, j As Integer
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String
    Dim b_c As Boolean
    Dim JIA As String, JIAN As String, CHENG As String, CHU As String, DENG As String
    JIA = "+"
    JIAN = "-"
    CHENG = "×"
    CHU = "÷"
    DENG = "=      "
    If op1 = "+" Then op1 = JIA
    If op1 = "-" Then op1 = JIAN
    If op1 = "x" Then op1 = CHENG
    If op1 = "/" Then op1 = CHU
    If op2 = "+" Then op2 = JIA
    If op2 = "-" Then op2 = JIAN
    If op2 = "x" Then op2 = CHENG
    If op2 = "/" Then op2 = CHU
    If op1 = "JIA" Then op1 = JIA
    If op1 = "JIAN" Then op1 = JIAN
    If op1 = "CHENG" Then op1 = CHENG
    If op1 = "CHU" Then op1 = CHU
    If op2 = "JIA" Then op2 = JIA
    If op2 = "JIAN" Then op2 = JIAN
    If op2 = "CHENG" Then op2 = CHENG
    If op2 = "CHU" Then op2 = CHU
    If op1 = "加" Then op1 = JIA
    If op1 = "减" Then op1 = JIAN
    If op1 = "乘" Then op1 = CHENG
    If op1 = "除" Then op1 = CHU
    If op2 = "加" Then op2 = JIA
    If op2 = "减" Then op2 = JIAN
    If op2 = "乘" Then op2 = CHENG
    If op2 = "除" Then op2 = CHU
    If op1 = "除以" Then op2 = CHU
    If op2 = "除以" Then op2 = CHU
    b_c = True
    Do While b_c
        i1 = Int(Rnd() * 99 + 1)
        i2 = Int(Rnd() * 99 + 1)
        If i1 + i2 <= 100 Then
            b_c = False
        End If
    Loop
    b_c = True
    Do While b_c
        i3 = Int(Rnd() * 99 + 1)
        i4 = Int(Rnd() * 99 + 1)
        i5 = Int(Rnd() * 99 + 1)
        If i3 + i4 + i5 <= 100 Then
            b_c = False
        End If
    Loop
    b_c = True
    Do While b_c
        i6 = Int(Rnd() * 8 + 2)
        i7 = Int(Rnd() * 8 + 2)
        If i6 * i7 <= 100 Then
            b_c = False
        End If
    Loop
    b_c = True
    Do While b_c
        i8 = Int(Rnd() * 8 + 2)
        i9 = Int(Rnd() * 8 + 2)
        i10 = Int(Rnd() * 8 + 2)
        If i8 * i9 * i10 <= 100 Then
            b_c = False
        End If
    Loop
    If op1 = "" Then
        i = Int(Rnd() * 4 + 1)
        If i = 1 Then
            op1 = JIA
        ElseIf i = 2 Then
            op1 = JIAN
        ElseIf i = 3 Then
            op1 = CHENG
        Else
            op1 = CHU
        End If
        i = Int(Rnd() * 5 + 1)
        If i = 1 Then
            op2 = JIA
        ElseIf i = 2 Then
            op2 = JIAN
        ElseIf i = 3 Then
            op2 = CHENG
        ElseIf i = 4 Then
            op2 = CHU
        Else
            op2 = ""
        End If
    End If
    If op1 = JIA And op2 = "" Then
        sizeyunsuan = i1 & JIA & i2 & DENG & i1 + i2
    ElseIf op1 = JIAN And op2 = "" Then
        sizeyunsuan = i1 + i2 & JIAN & i1 & DENG & i2
    ElseIf op1 = CHENG And op2 = "" Then
        sizeyunsuan = i6 & CHENG & i7 & DENG & i6 * i7
    ElseIf op1 = CHU And op2 = "" Then
        sizeyunsuan = i6 * i7 & CHU & i6 & DENG & i7
    ElseIf op1 = JIA And op2 = JIA Then
        sizeyunsuan = i3 & JIA & i4 & JIA & i5 & DENG & i3 + i4 + i5
    ElseIf op1 = JIA And op2 = JIAN Then
        i = Int(Rnd() * (i1 + i2) + 1)
        sizeyunsuan = i1 & JIA & i2 & JIAN & i & DENG & i1 + i2 - i
    ElseIf op1 = JIA And op2 = CHENG Then
        i = Int(Rnd() * (100 - i6 * i7) + 1)
        sizeyunsuan = i & JIA & i6 & CHENG & i7 & DENG & i + i6 * i7
    ElseIf op1 = JIA And op2 = CHU Then
        i = Int(Rnd() * (100 - i6) + 1)
        sizeyunsuan = i & JIA & i6 * i7 & CHU & i7 & DENG & i + i6
    ElseIf op1 = JIAN And op2 = JIA Then
        i = Int(Rnd() * (100 - i2) + 1)
        sizeyunsuan = i1 + i2 & JIAN & i1 & JIA & i & DENG & i2 + i
    ElseIf op1 = JIAN And op2 = JIAN Then
        sizeyunsuan = i3 + i4 + i5 & JIAN & i4 & JIAN & i5 & DENG & i3
    ElseIf op1 = JIAN And op2 = CHENG Then
        i = Int(Rnd() * (100 - i6 * i7) + i6 * i7)
        sizeyunsuan = i & JIAN & i6 & CHENG & i7 & DENG & i - i6 * i7
    ElseIf op1 = JIAN And op2 = CHU Then
        i = Int(Rnd() * (100 - i6) + 1)
        sizeyunsuan = i & JIAN & i6 * i7 & CHU & i7 & DENG & i - i6
    ElseIf op1 = CHENG And op2 = JIA Then
        i = Int(Rnd() * (100 - i6 * i7) + 1)
        sizeyunsuan = i6 & CHENG & i7 & JIA & i & DENG & i + i6 * i7
    ElseIf op1 = CHENG And op2 = JIAN Then
        i = Int(Rnd() * i6 * i7 + 1)
        sizeyunsuan = i6 & CHENG & i7 & JIAN & i & DENG & i6 * i7 - i
        
        
    ElseIf op1 = CHENG And op2 = CHENG Then
        sizeyunsuan = i8 & CHENG & i9 & CHENG & i10 & DENG & i8 * i9 * i10
       
        
    ElseIf op1 = CHENG And op2 = CHU Then
        i = Int(Rnd() * Int(100 / (i6 * i7)) + 1)
        sizeyunsuan = i * i6 & CHENG & i7 & CHU & i & DENG & i6 * i7
    ElseIf op1 = CHU And op2 = JIA Then
        i = Int(Rnd() * (100 - i6) + 1)
        sizeyunsuan = i6 * i7 & CHU & i7 & JIA & i & DENG & i + i6
    ElseIf op1 = CHU And op2 = JIAN Then
        i = Int(Rnd() * i6 + 1)
        sizeyunsuan = i6 * i7 & CHU & i7 & JIAN & i & DENG & i6 - i
    ElseIf op1 = CHU And op2 = CHENG Then
        i = Int(Rnd() * (100 / i6))
        sizeyunsuan = i6 * i7 & CHU & i7 & CHENG & i & DENG & i * i6
    ElseIf op1 = CHU And op2 = CHU Then
        sizeyunsuan = i8 * i9 * i10 & CHU & i8 & CHU & i9 & DENG & i10
    End If
    
    If InStr(sizeyunsuan, CHENG & 1) > 0 Or InStr(sizeyunsuan, CHU & 1) > 0 Or InStr(sizeyunsuan, 1 & CHENG) > 0 Then
    sizeyunsuan = sizeyunsuan(op1, op2)
    ElseIf InStr(sizeyunsuan, CHENG & 0) > 0 Then
    sizeyunsuan = sizeyunsuan(op1, op2)
    End If
    
    
End Function
Function xx() As String
xx = "XXXXXXAFASF"
End Function
