Attribute VB_Name = "Functions"
Option Explicit
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As pointapi, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal nXOffset As Long, ByVal nYOffset As Long) As Long
Public Type pointapi
    X As Long
    Y As Long
End Type

Public bentukAkhir As Long


Public Function BuatElips(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Long
    Const PI As Double = 3.14159265358979
    Dim i As Long, pts(0 To 15) As pointapi
    Dim a As Double, b As Double, CX As Double, CY As Double, angle As Double
    a = (X2 - X1) / 2
    b = (Y2 - Y1) / 2
    CX = X1 + a
    CY = Y1 + b
    For i = 0 To 15
        angle = (2 * PI * i) / 16
        pts(i).X = CX + a * Cos(angle)
        pts(i).Y = CY + b * Sin(angle)
    Next i
    BuatElips = CreatePolygonRgn(pts(0), 16, 2)
End Function

Public Function BuatKotak(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Long
    Dim pts(0 To 3) As pointapi
    pts(0).X = X1: pts(0).Y = Y1
    pts(1).X = X2: pts(1).Y = Y1
    pts(2).X = X2: pts(2).Y = Y2
    pts(3).X = X1: pts(3).Y = Y2
    BuatKotak = CreatePolygonRgn(pts(0), 4, 2)
End Function

Public Function BuatHurufA() As Long
    Dim bentukA As Long
    Dim kakiKiri As Long, kakiKanan As Long, palangTengah As Long
    Dim pts(0 To 3) As pointapi

    pts(0).X = 152.5:  pts(0).Y = 140
    pts(1).X = 165:  pts(1).Y = 140
    pts(2).X = 130:  pts(2).Y = 250
    pts(3).X = 115:  pts(3).Y = 250
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)

    pts(0).X = 155:  pts(0).Y = 140
    pts(1).X = 167.5:  pts(1).Y = 140
    pts(2).X = 205:  pts(2).Y = 250
    pts(3).X = 190:  pts(3).Y = 250
    kakiKanan = CreatePolygonRgn(pts(0), 4, 2)

    palangTengah = BuatKotak(135, 210, 188, 225)

    bentukA = kakiKiri
    CombineRgn bentukA, bentukA, kakiKanan, 2
    CombineRgn bentukA, bentukA, palangTengah, 2

    DeleteObject kakiKanan
    DeleteObject palangTengah

    BuatHurufA = bentukA
End Function

Public Function BuatHurufBe() As Long
    Dim bentukBe As Long
    Dim palangAtas As Long, tiangVertikal As Long
    Dim lingkaranBawah As Long, lingkaranDalam As Long, setengahBawah As Long
    Dim pLuar(35) As pointapi, pDalam(35) As pointapi
    Dim i As Long
    Dim CX As Long, CY As Long
    Dim rxL As Long, ryL As Long
    Dim rxD As Long, ryD As Long

    palangAtas = BuatKotak(235, 140, 325, 155)
    tiangVertikal = BuatKotak(235, 140, 250, 250)

    CX = (130 + 325) / 2
    CY = (185 + 251) / 2
    rxL = (325 - 130) / 2
    ryL = (251 - 185) / 2
    rxD = (310 - 145) / 2
    ryD = (236 - 200) / 2

    For i = 0 To 35
        pLuar(i).X = CX + rxL * Cos(i * 3.14159265 / 18)
        pLuar(i).Y = CY + ryL * Sin(i * 3.14159265 / 18)
        pDalam(i).X = CX + rxD * Cos(i * 3.14159265 / 18)
        pDalam(i).Y = CY + ryD * Sin(i * 3.14159265 / 18)
    Next i

    lingkaranBawah = CreatePolygonRgn(pLuar(0), 36, 1)
    lingkaranDalam = CreatePolygonRgn(pDalam(0), 36, 1)
    setengahBawah = BuatKotak(235, 140, 325, 251)

    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalam, 3
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1

    bentukBe = tiangVertikal
    CombineRgn bentukBe, bentukBe, palangAtas, 2
    CombineRgn bentukBe, bentukBe, lingkaranBawah, 2

    DeleteObject palangAtas
    DeleteObject lingkaranBawah
    DeleteObject lingkaranDalam
    DeleteObject setengahBawah

    BuatHurufBe = bentukBe
End Function

Public Function BuatHurufVe() As Long
    Dim bentukVe As Long
    Dim tiangVertikal As Long
    Dim lingkaranAtas As Long, lingkaranDalamAtas As Long
    Dim lingkaranBawah As Long, lingkaranDalamBawah As Long
    Dim setengahBawah As Long
    Dim pLuarAtas(35) As pointapi, pDalamAtas(35) As pointapi
    Dim pLuarBawah(35) As pointapi, pDalamBawah(35) As pointapi
    Dim i As Long
    Dim cxA As Long, cyA As Long, rxLA As Long, ryLA As Long, rxDA As Long, ryDA As Long
    Dim cxB As Long, cyB As Long, rxLB As Long, ryLB As Long, rxDB As Long, ryDB As Long

    tiangVertikal = BuatKotak(355, 140, 370, 250)

    cxA = (225 + 445) / 2
    cyA = (140 + 203) / 2
    rxLA = (445 - 225) / 2
    ryLA = (203 - 140) / 2
    rxDA = (430 - 240) / 2
    ryDA = (188 - 155) / 2

    cxB = (225 + 445) / 2
    cyB = (188 + 251) / 2
    rxLB = (445 - 225) / 2
    ryLB = (251 - 188) / 2
    rxDB = (430 - 240) / 2
    ryDB = (236 - 203) / 2

    For i = 0 To 35
        pLuarAtas(i).X = cxA + rxLA * Cos(i * 3.14159265 / 18)
        pLuarAtas(i).Y = cyA + ryLA * Sin(i * 3.14159265 / 18)
        pDalamAtas(i).X = cxA + rxDA * Cos(i * 3.14159265 / 18)
        pDalamAtas(i).Y = cyA + ryDA * Sin(i * 3.14159265 / 18)

        pLuarBawah(i).X = cxB + rxLB * Cos(i * 3.14159265 / 18)
        pLuarBawah(i).Y = cyB + ryLB * Sin(i * 3.14159265 / 18)
        pDalamBawah(i).X = cxB + rxDB * Cos(i * 3.14159265 / 18)
        pDalamBawah(i).Y = cyB + ryDB * Sin(i * 3.14159265 / 18)
    Next i

    lingkaranAtas = CreatePolygonRgn(pLuarAtas(0), 36, 1)
    lingkaranDalamAtas = CreatePolygonRgn(pDalamAtas(0), 36, 1)
    lingkaranBawah = CreatePolygonRgn(pLuarBawah(0), 36, 1)
    lingkaranDalamBawah = CreatePolygonRgn(pDalamBawah(0), 36, 1)
    setengahBawah = BuatKotak(355, 140, 445, 250)

    CombineRgn lingkaranAtas, lingkaranAtas, lingkaranDalamAtas, 3
    CombineRgn lingkaranAtas, lingkaranAtas, setengahBawah, 1
    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalamBawah, 3
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1

    bentukVe = tiangVertikal
    CombineRgn bentukVe, bentukVe, lingkaranAtas, 2
    CombineRgn bentukVe, bentukVe, lingkaranBawah, 2

    DeleteObject lingkaranAtas
    DeleteObject lingkaranDalamAtas
    DeleteObject lingkaranBawah
    DeleteObject lingkaranDalamBawah
    DeleteObject setengahBawah

    BuatHurufVe = bentukVe
End Function

Public Function BuatHurufGe() As Long
    Dim bentukGe As Long
    Dim palangAtas As Long, tiangVertikal As Long

    palangAtas = BuatKotak(475, 140, 565, 155)
    tiangVertikal = BuatKotak(475, 140, 490, 250)

    bentukGe = tiangVertikal
    CombineRgn bentukGe, bentukGe, palangAtas, 2

    DeleteObject palangAtas
    BuatHurufGe = bentukGe
End Function

Public Function BuatHurufDe() As Long
    Dim bentukDe As Long
    Dim palangBawah As Long, potonganBawah As Long, bagianAtas As Long, potonganAtas As Long
    Dim bagianKiri As Long

    palangBawah = BuatKotak(595, 215, 685, 250)
    potonganBawah = BuatKotak(610, 230, 670, 250)
    bentukDe = palangBawah

    CombineRgn bentukDe, bentukDe, potonganBawah, 3

    bagianAtas = BuatKotak(615, 140, 678, 215)
    potonganAtas = BuatKotak(620, 155, 663, 215)
    CombineRgn bagianAtas, bagianAtas, potonganAtas, 3

    bagianKiri = BuatKotak(605, 140, 620, 215)

    CombineRgn bentukDe, bentukDe, bagianAtas, 2
    CombineRgn bentukDe, bentukDe, bagianKiri, 2

    BuatHurufDe = bentukDe
    DeleteObject potonganBawah
    DeleteObject bagianAtas
    DeleteObject potonganAtas
    DeleteObject bagianKiri
End Function

Public Function BuatHurufYe() As Long
    Dim bentukYe As Long
    Dim palangAtas As Long, palangTengah As Long, palangBawah As Long, tiangVertikal As Long

    palangAtas = BuatKotak(725, 140, 795, 155)
    palangTengah = BuatKotak(725, 185, 795, 200)
    palangBawah = BuatKotak(725, 235, 795, 250)

    tiangVertikal = BuatKotak(725, 140, 740, 250)

    bentukYe = tiangVertikal
    CombineRgn bentukYe, bentukYe, palangAtas, 2
    CombineRgn bentukYe, bentukYe, palangTengah, 2
    CombineRgn bentukYe, bentukYe, palangBawah, 2

    BuatHurufYe = bentukYe

    DeleteObject palangAtas
    DeleteObject palangTengah
    DeleteObject palangBawah
End Function

Public Function BuatHurufYo() As Long
    Dim bentukYo As Long
    Dim palangAtas As Long, palangTengah As Long, palangBawah As Long, tiangVertikal As Long
    Dim mata As Long, mata2 As Long

    palangAtas = BuatKotak(845, 140, 915, 155)
    palangTengah = BuatKotak(845, 185, 915, 200)
    palangBawah = BuatKotak(845, 235, 915, 250)
    
    mata = BuatElips(860, 115, 875, 130)
    mata2 = BuatElips(885, 115, 900, 130)

    tiangVertikal = BuatKotak(845, 140, 860, 250)

    bentukYo = tiangVertikal
    CombineRgn mata, mata, mata2, 2
    CombineRgn bentukYo, bentukYo, palangAtas, 2
    CombineRgn bentukYo, bentukYo, palangTengah, 2
    CombineRgn bentukYo, bentukYo, palangBawah, 2
    CombineRgn bentukYo, bentukYo, mata, 2

    BuatHurufYo = bentukYo

    DeleteObject palangAtas
    DeleteObject palangTengah
    DeleteObject palangBawah
    DeleteObject mata
    DeleteObject mata2
End Function

Public Function BuatHurufZhe() As Long
    Dim bentukZhe As Long
    Dim bagianAtasKanan As Long, potonganKanan1 As Long, potonganKanan2, bagianAtasKanan2, potonganKanan3, potonganKanan4
    Dim bagianAtasKiri As Long, potonganKiri1 As Long, potonganKiri2, bagianAtasKiri2, potonganKiri3, potonganKiri4
    Dim tiangVertikal As Long, palangTengah As Long, kakiKanan As Long, kakiKiri As Long
    Dim ptsKanan(0 To 3) As pointapi, ptsKiri(0 To 3) As pointapi
    Dim i As Integer, garisCerminX As Long
    garisCerminX = 1000

    bagianAtasKanan = BuatElips(992.5, 140, 1052.5, 205)
    potonganKanan1 = BuatElips(1007.5, 155, 1037.5, 190)
    potonganKanan2 = BuatKotak(1022.5, 170, 1052.5, 205)
    bagianAtasKanan2 = BuatElips(1037.5, 140, 1102.5, 205)
    potonganKanan3 = BuatElips(1052.5, 155, 1087.5, 190)
    potonganKanan4 = BuatKotak(1037.5, 140, 1067.5, 170)

    ptsKanan(0).X = 1010: ptsKanan(0).Y = 200
    ptsKanan(1).X = 1030: ptsKanan(1).Y = 200
    ptsKanan(2).X = 1065: ptsKanan(2).Y = 250
    ptsKanan(3).X = 1045: ptsKanan(3).Y = 250
    kakiKanan = CreatePolygonRgn(ptsKanan(0), 4, 2)

    bagianAtasKiri = BuatElips((2 * garisCerminX) - 1052.5, 140, (2 * garisCerminX) - 992.5, 205)
    potonganKiri1 = BuatElips((2 * garisCerminX) - 1037.5, 155, (2 * garisCerminX) - 1007.5, 190)
    potonganKiri2 = BuatKotak((2 * garisCerminX) - 1052.5, 170, (2 * garisCerminX) - 1022.5, 205)
    bagianAtasKiri2 = BuatElips((2 * garisCerminX) - 1102.5, 140, (2 * garisCerminX) - 1037.5, 205)
    potonganKiri3 = BuatElips((2 * garisCerminX) - 1087.5, 155, (2 * garisCerminX) - 1052.5, 190)
    potonganKiri4 = BuatKotak((2 * garisCerminX) - 1067.5, 140, (2 * garisCerminX) - 1037.5, 170)

    For i = 0 To 3
        ptsKiri(i).X = (2 * garisCerminX) - ptsKanan(i).X
        ptsKiri(i).Y = ptsKanan(i).Y
    Next i
    kakiKiri = CreatePolygonRgn(ptsKiri(0), 4, 2)

    tiangVertikal = BuatKotak(992.5, 140, 1007.5, 250)
    palangTengah = BuatKotak(977.5, 190, 1022.5, 205)

    bentukZhe = tiangVertikal

    CombineRgn bagianAtasKanan, bagianAtasKanan, potonganKanan1, 3
    CombineRgn bagianAtasKanan, bagianAtasKanan, potonganKanan2, 1
    CombineRgn bagianAtasKanan2, bagianAtasKanan2, potonganKanan3, 3
    CombineRgn bagianAtasKanan2, bagianAtasKanan2, potonganKanan4, 1

    CombineRgn bagianAtasKiri, bagianAtasKiri, potonganKiri1, 3
    CombineRgn bagianAtasKiri, bagianAtasKiri, potonganKiri2, 1
    CombineRgn bagianAtasKiri2, bagianAtasKiri2, potonganKiri3, 3
    CombineRgn bagianAtasKiri2, bagianAtasKiri2, potonganKiri4, 1

    CombineRgn bagianAtasKanan, bagianAtasKanan, bagianAtasKanan2, 2
    CombineRgn bagianAtasKiri, bagianAtasKiri, bagianAtasKiri2, 2

    CombineRgn bentukZhe, bentukZhe, bagianAtasKanan, 2
    CombineRgn bentukZhe, bentukZhe, bagianAtasKiri, 2

    CombineRgn bentukZhe, bentukZhe, palangTengah, 2

    CombineRgn bentukZhe, bentukZhe, kakiKanan, 2
    CombineRgn bentukZhe, bentukZhe, kakiKiri, 2

    BuatHurufZhe = bentukZhe

    DeleteObject bagianAtasKanan
    DeleteObject bagianAtasKanan2
    DeleteObject potonganKanan1
    DeleteObject potonganKanan2
    DeleteObject potonganKanan3
    DeleteObject potonganKanan4
    DeleteObject bagianAtasKiri
    DeleteObject bagianAtasKiri2
    DeleteObject potonganKiri1
    DeleteObject potonganKiri2
    DeleteObject potonganKiri3
    DeleteObject potonganKiri4
    DeleteObject palangTengah
    DeleteObject kakiKanan
    DeleteObject kakiKiri
    DeleteObject i
    DeleteObject garisCerminX
End Function

Public Function BuatHurufZe() As Long
    Dim bentukZe As Long
    Dim lingkaranAtas As Long, potonganAtas As Long
    Dim lingkaranBawah As Long, potonganBawah As Long
    Dim potonganTengah As Long
    Dim pts(0 To 3) As pointapi

    lingkaranAtas = BuatElips(1085, 140, 1155, 205)
    potonganAtas = BuatElips(1100, 155, 1140, 190)
    CombineRgn lingkaranAtas, lingkaranAtas, potonganAtas, 3

    lingkaranBawah = BuatElips(1085, 190, 1155, 250)
    potonganBawah = BuatElips(1100, 205, 1140, 235)
    CombineRgn lingkaranBawah, lingkaranBawah, potonganBawah, 3

    pts(0).X = 1050: pts(0).Y = 160
    pts(1).X = 1115: pts(1).Y = 170
    pts(2).X = 1115: pts(2).Y = 220
    pts(3).X = 1050: pts(3).Y = 230
    potonganTengah = CreatePolygonRgn(pts(0), 4, 2)

    bentukZe = lingkaranAtas
    CombineRgn bentukZe, bentukZe, lingkaranBawah, 2
    CombineRgn bentukZe, bentukZe, potonganTengah, 4

    BuatHurufZe = bentukZe

    DeleteObject lingkaranBawah
    DeleteObject potonganAtas
    DeleteObject potonganBawah
    DeleteObject potonganTengah
End Function

Public Function BuatHurufI() As Long
    Dim bentukI As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangDiagonal As Long
    Dim pts(0 To 3) As pointapi

    palangVertikal = BuatKotak(1195, 140, 1210, 250)
    palangVertikal2 = BuatKotak(1270, 140, 1285, 250)

    pts(0).X = 1270:  pts(0).Y = 140
    pts(1).X = 1285:  pts(1).Y = 140
    pts(2).X = 1210:  pts(2).Y = 250
    pts(3).X = 1195:  pts(3).Y = 250
    palangDiagonal = CreatePolygonRgn(pts(0), 4, 2)

    bentukI = palangVertikal
    CombineRgn bentukI, bentukI, palangVertikal2, 2
    CombineRgn bentukI, bentukI, palangDiagonal, 2

    BuatHurufI = bentukI

    DeleteObject palangVertikal2
    DeleteObject palangDiagonal
End Function

Public Function BuatHurufIy() As Long
    Dim bentukIy As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangDiagonal As Long, lengkungan As Long, lengkunganDalam As Long, potonganLengkung As Long
    Dim pts(0 To 3) As pointapi

    palangVertikal = BuatKotak(1315, 140, 1330, 250)
    palangVertikal2 = BuatKotak(1390, 140, 1405, 250)
    lengkungan = BuatElips(1335, 85, 1390, 130)
    lengkunganDalam = BuatElips(1350, 100, 1375, 115)
    potonganLengkung = BuatKotak(1320, 80, 1390, 105)

    pts(0).X = 1390:  pts(0).Y = 140
    pts(1).X = 1405:  pts(1).Y = 140
    pts(2).X = 1330:  pts(2).Y = 250
    pts(3).X = 1315:  pts(3).Y = 250
    palangDiagonal = CreatePolygonRgn(pts(0), 4, 2)

    bentukIy = palangVertikal
    CombineRgn bentukIy, bentukIy, palangVertikal2, 2
    CombineRgn bentukIy, bentukIy, palangDiagonal, 2
    CombineRgn bentukIy, bentukIy, lengkungan, 2
    CombineRgn bentukIy, bentukIy, lengkunganDalam, 3
    CombineRgn bentukIy, bentukIy, potonganLengkung, 4

    BuatHurufIy = bentukIy

    DeleteObject palangVertikal2
    DeleteObject palangDiagonal
    DeleteObject lengkungan
    DeleteObject lengkunganDalam
    DeleteObject potonganLengkung
End Function

Public Function BuatHurufKa() As Long
    Dim bentukKa As Long
    Dim bagianAtasKanan As Long, potonganKanan1 As Long, potonganKanan2 As Long, bagianAtasKanan2 As Long, potonganKanan3 As Long, potonganKanan4 As Long
    Dim tiangVertikal As Long, palangTengah As Long, kakiKanan As Long
    Dim ptsKanan(0 To 3) As pointapi

    bagianAtasKanan = BuatElips(1435, 140, 1501, 205)
    potonganKanan1 = BuatElips(1450, 155, 1486, 190)
    potonganKanan2 = BuatKotak(1472.5, 170, 1501, 205)
    bagianAtasKanan2 = BuatElips(1485, 140, 1550, 205)
    potonganKanan3 = BuatElips(1500, 155, 1535, 190)
    potonganKanan4 = BuatKotak(1485, 140, 1515, 170)

    ptsKanan(0).X = 1468: ptsKanan(0).Y = 195
    ptsKanan(1).X = 1485: ptsKanan(1).Y = 195
    ptsKanan(2).X = 1520: ptsKanan(2).Y = 250
    ptsKanan(3).X = 1502: ptsKanan(3).Y = 250
    kakiKanan = CreatePolygonRgn(ptsKanan(0), 4, 2)

    tiangVertikal = BuatKotak(1435, 140, 1450, 250)
    palangTengah = BuatKotak(1435, 187.5, 1475, 203)

    bentukKa = tiangVertikal

    CombineRgn bagianAtasKanan, bagianAtasKanan, potonganKanan1, 3
    CombineRgn bagianAtasKanan, bagianAtasKanan, potonganKanan2, 1
    CombineRgn bagianAtasKanan2, bagianAtasKanan2, potonganKanan3, 3
    CombineRgn bagianAtasKanan2, bagianAtasKanan2, potonganKanan4, 1

    CombineRgn bagianAtasKanan, bagianAtasKanan, bagianAtasKanan2, 2
    CombineRgn bentukKa, bentukKa, bagianAtasKanan, 2
    CombineRgn bentukKa, bentukKa, palangTengah, 2
    CombineRgn bentukKa, bentukKa, kakiKanan, 2

    BuatHurufKa = bentukKa

    DeleteObject bagianAtasKanan
    DeleteObject bagianAtasKanan2
    DeleteObject potonganKanan1
    DeleteObject potonganKanan2
    DeleteObject potonganKanan3
    DeleteObject potonganKanan4
    DeleteObject palangTengah
    DeleteObject kakiKanan
End Function

Public Function BuatHurufEl() As Long
    Dim bentukEl As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangHorizontal As Long
    Dim palangKiri As Long, potonganAtas As Long, potonganKiri As Long

    palangVertikal = BuatKotak(135, 300, 150, 380)
    palangVertikal2 = BuatKotak(190, 300, 205, 400)
    palangHorizontal = BuatKotak(135, 300, 205, 315)
    palangKiri = BuatKotak(120, 375, 145, 390)
    'potonganAtas = BuatKotak(-50, 180, 136, 385, 30, 30)
    potonganKiri = BuatKotak(0, 230, 115, 400)

    bentukEl = palangVertikal

    CombineRgn bentukEl, bentukEl, palangVertikal2, 2
    CombineRgn bentukEl, bentukEl, palangHorizontal, 2
    CombineRgn bentukEl, bentukEl, palangKiri, 2
    CombineRgn bentukEl, bentukEl, potonganAtas, 4
    CombineRgn bentukEl, bentukEl, potonganKiri, 4

    BuatHurufEl = bentukEl

    DeleteObject palangVertikal2
    DeleteObject palangHorizontal
    DeleteObject palangKiri
    DeleteObject potonganAtas
    DeleteObject potonganKiri
End Function

Public Function BuatHurufEm() As Long
    Dim bentukEm As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangDiagonal As Long, palangDiagonal2 As Long
    Dim ptsKanan(0 To 3) As pointapi, ptsKiri(0 To 3) As pointapi

    palangVertikal = BuatKotak(310, 310, 325, 400)
    palangVertikal2 = BuatKotak(235, 310, 250, 400)

    ptsKanan(0).X = 245: ptsKanan(0).Y = 310
    ptsKanan(1).X = 260: ptsKanan(1).Y = 310
    ptsKanan(2).X = 287.5: ptsKanan(2).Y = 400
    ptsKanan(3).X = 272.5: ptsKanan(3).Y = 400
    palangDiagonal = CreatePolygonRgn(ptsKanan(0), 4, 2)

    ptsKiri(0).X = 300: ptsKiri(0).Y = 310
    ptsKiri(1).X = 315: ptsKiri(1).Y = 310
    ptsKiri(2).X = 287.5: ptsKiri(2).Y = 400
    ptsKiri(3).X = 272.5: ptsKiri(3).Y = 400
    palangDiagonal2 = CreatePolygonRgn(ptsKiri(0), 4, 2)

    bentukEm = palangVertikal
    CombineRgn bentukEm, bentukEm, palangVertikal2, 2
    CombineRgn bentukEm, bentukEm, palangDiagonal, 2
    CombineRgn bentukEm, bentukEm, palangDiagonal2, 2

    BuatHurufEm = bentukEm

    DeleteObject palangVertikal2
    DeleteObject palangDiagonal
    DeleteObject palangDiagonal2
End Function

Public Function BuatHurufEn() As Long
    Dim bentukEn As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangTengah As Long

    palangVertikal = BuatKotak(365, 310, 380, 400)
    palangVertikal2 = BuatKotak(420, 310, 435, 400)
    palangTengah = BuatKotak(365, 347.5, 435, 362.5)

    bentukEn = palangVertikal

    CombineRgn bentukEn, bentukEn, palangVertikal2, 2
    CombineRgn bentukEn, bentukEn, palangTengah, 2

    BuatHurufEn = bentukEn

    DeleteObject palangVertikal2
    DeleteObject palangTengah
End Function

Public Function BuatHurufO() As Long
    Dim luar(0 To 15) As pointapi
    Dim dalam(0 To 15) As pointapi
    Dim i As Integer

    Dim CX As Long: CX = 520
    Dim CY As Long: CY = 360

    Dim dxOuter As Variant
    Dim dyOuter As Variant
    Dim scaleInner As Double: scaleInner = 0.7

    dxOuter = Array(0, 25, 45, 55, 60, 55, 45, 25, 0, -25, -45, -55, -60, -55, -45, -25)
    dyOuter = Array(-60, -55, -45, -25, 0, 25, 45, 55, 60, 55, 45, 25, 0, -25, -45, -55)

    For i = 0 To 15
        luar(i).X = CX + CLng(dxOuter(i))
        luar(i).Y = CY + CLng(dyOuter(i))

        dalam(i).X = CX + CLng(dxOuter(i) * scaleInner)
        dalam(i).Y = CY + CLng(dyOuter(i) * scaleInner)
    Next i

    Dim regLuar As Long, regDalam As Long, bentuk As Long
    regLuar = CreatePolygonRgn(luar(0), 16, 2)
    regDalam = CreatePolygonRgn(dalam(0), 16, 2)

    bentuk = regLuar
    CombineRgn bentuk, bentuk, regDalam, 4

    BuatHurufO = bentuk

    DeleteObject regDalam
End Function





Public Function BuatHurufPe() As Long
    Dim bentukPe As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangAtas As Long

    palangVertikal = BuatKotak(605, 310, 620, 400)
    palangVertikal2 = BuatKotak(660, 310, 675, 400)
    palangAtas = BuatKotak(605, 310, 675, 325)

    bentukPe = palangVertikal

    CombineRgn bentukPe, bentukPe, palangVertikal2, 2
    CombineRgn bentukPe, bentukPe, palangAtas, 2

    BuatHurufPe = bentukPe

    DeleteObject palangVertikal2
    DeleteObject palangAtas
End Function

Public Function BuatHurufEr() As Long
    Dim bentukEr As Long
    Dim tiangVertikal As Long, lingkaranAtas As Long, lingkaranDalam As Long, setengahAtas As Long

    tiangVertikal = BuatKotak(715, 310, 730, 400)

    'lingkaranAtas = BuatKotak(610, 300, 795, 361, 60, 60)
    'lingkaranDalam = BuatKotak(625, 315, 780, 346, 40, 40)
    setengahAtas = BuatKotak(715, 300, 805, 400)

    bentukEr = tiangVertikal

    CombineRgn lingkaranAtas, lingkaranAtas, lingkaranDalam, 3
    CombineRgn lingkaranAtas, lingkaranAtas, setengahAtas, 1

    CombineRgn bentukEr, bentukEr, lingkaranAtas, 2

    BuatHurufEr = bentukEr

    DeleteObject lingkaranAtas
    DeleteObject lingkaranDalam
    DeleteObject setengahAtas
End Function

Public Function BuatHurufEs() As Long
    Dim bentukEs As Long
    Dim luar As Long, dalam As Long, samping As Long
    Dim pts(0 To 3) As pointapi

    'luar = BuatKotak(835, 290, 925, 400, 90, 90)
    'dalam = BuatKotak(850, 305, 910, 385, 60, 60)

    pts(0).X = 875: pts(0).Y = 330
    pts(1).X = 940: pts(1).Y = 315
    pts(2).X = 940: pts(2).Y = 375
    pts(3).X = 875: pts(3).Y = 360
    samping = CreatePolygonRgn(pts(0), 4, 2)

    bentukEs = luar

    CombineRgn bentukEs, bentukEs, dalam, 3
    CombineRgn bentukEs, bentukEs, samping, 4

    BuatHurufEs = bentukEs

    DeleteObject dalam
    DeleteObject samping
End Function

Public Function BuatHurufTe() As Long
    Dim bentukTe As Long
    Dim palangVertikal As Long, palangAtas As Long

    palangVertikal = BuatKotak(992.5, 290, 1007.5, 400)
    palangAtas = BuatKotak(955, 290, 1045, 305)

    bentukTe = palangVertikal

    CombineRgn bentukTe, bentukTe, palangAtas, 2

    BuatHurufTe = bentukTe

    DeleteObject palangAtas
    BuatHurufTe = bentukTe
End Function

Public Function BuatHurufU() As Long
    Dim bentukU As Long
    Dim palangKiri As Long, potonganAtas As Long, potonganKiri As Long
    Dim kakiKiri As Long
    Dim kakiKanan As Long
    Dim pts(0 To 3) As pointapi

    pts(0).X = 1075:  pts(0).Y = 290
    pts(1).X = 1090:  pts(1).Y = 290
    pts(2).X = 1127.5:  pts(2).Y = 370
    pts(3).X = 1112.5:  pts(3).Y = 370
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)

    pts(0).X = 1112.5:  pts(0).Y = 375
    pts(1).X = 1127.5:  pts(1).Y = 370
    pts(2).X = 1165:  pts(2).Y = 290
    pts(3).X = 1150:  pts(3).Y = 290
    kakiKanan = CreatePolygonRgn(pts(0), 4, 2)

    'palangKiri = BuatKotak(1000, 347, 1128, 400, 50, 50)
    'potonganAtas = BuatKotak(1000, 305, 1113, 385, 25, 25)
    potonganKiri = BuatKotak(1000, 300, 1085, 400)

    CombineRgn palangKiri, palangKiri, potonganAtas, 4
    CombineRgn palangKiri, palangKiri, potonganKiri, 4

    bentukU = palangKiri

    CombineRgn bentukU, bentukU, kakiKanan, 2
    CombineRgn bentukU, bentukU, kakiKiri, 2

    DeleteObject potonganAtas
    DeleteObject potonganKiri
    DeleteObject kakiKanan
    DeleteObject kakiKiri

    BuatHurufU = bentukU
End Function

Public Function BuatHurufEf() As Long
    Dim bentukEf As Long
    Dim luar As Long, dalam As Long, palangTengah As Long

    'luar = BuatKotak(1195, 320, 1285, 390, 60, 60)
    'dalam = BuatKotak(1210, 335, 1270, 375, 40, 40)
    palangTengah = BuatKotak(1232.5, 310, 1247.5, 400)

    bentukEf = luar

    CombineRgn bentukEf, bentukEf, dalam, 3
    CombineRgn bentukEf, bentukEf, palangTengah, 2

    BuatHurufEf = bentukEf
End Function

Public Function BuatHurufKha() As Long
    Dim bentukKha As Long
    Dim kakiKiri As Long, kakiKanan As Long
    Dim pts(0 To 3) As pointapi

    pts(0).X = 1390:  pts(0).Y = 310
    pts(1).X = 1410:  pts(1).Y = 310
    pts(2).X = 1330:  pts(2).Y = 400
    pts(3).X = 1310:  pts(3).Y = 400
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)

    pts(0).X = 1390:  pts(0).Y = 400
    pts(1).X = 1410:  pts(1).Y = 400
    pts(2).X = 1330:  pts(2).Y = 310
    pts(3).X = 1310:  pts(3).Y = 310
    kakiKanan = CreatePolygonRgn(pts(0), 4, 2)

    bentukKha = kakiKiri
    CombineRgn bentukKha, bentukKha, kakiKanan, 2

    BuatHurufKha = bentukKha

    DeleteObject kakiKanan
End Function

Public Function BuatHurufTse() As Long
    Dim bentukTse As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangBawah As Long, palangKanan As Long

    palangVertikal = BuatKotak(1435, 290, 1450, 380)
    palangVertikal2 = BuatKotak(1500, 290, 1515, 380)
    palangBawah = BuatKotak(1435, 365, 1525, 380)
    palangKanan = BuatKotak(1510, 365, 1525, 400)

    bentukTse = palangVertikal

    CombineRgn bentukTse, bentukTse, palangVertikal2, 2
    CombineRgn bentukTse, bentukTse, palangBawah, 2
    CombineRgn bentukTse, bentukTse, palangKanan, 2

    BuatHurufTse = bentukTse

    DeleteObject palangVertikal2
    DeleteObject palangBawah
    DeleteObject palangKanan
End Function

Public Function BuatHurufChe() As Long
    Dim bentukChe As Long
    Dim lingkaranAtas As Long, lingkaranDalam As Long, setengahAtas As Long, palangVertikal As Long

    'lingkaranAtas = BuatKotak(125, 310, 195, 520, 55, 55)
    'lingkaranDalam = BuatKotak(140, 325, 190, 505, 40, 40)
    setengahAtas = BuatKotak(125, 250, 195, 460)
    palangVertikal = BuatKotak(180, 460, 195, 550)

    bentukChe = lingkaranAtas

    CombineRgn bentukChe, bentukChe, lingkaranDalam, 4
    CombineRgn bentukChe, bentukChe, setengahAtas, 4
    CombineRgn bentukChe, bentukChe, palangVertikal, 2

  BuatHurufChe = bentukChe
End Function

Public Function BuatHurufSha() As Long
    Dim bentukSha As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangVertikal3 As Long, palangBawah As Long

    palangVertikal = BuatKotak(235, 460, 250, 550)
    palangVertikal2 = BuatKotak(272.5, 460, 287.5, 550)
    palangVertikal3 = BuatKotak(310, 460, 325, 550)
    palangBawah = BuatKotak(235, 535, 325, 550)

    bentukSha = palangVertikal

    CombineRgn bentukSha, bentukSha, palangVertikal2, 2
    CombineRgn bentukSha, bentukSha, palangVertikal3, 2
    CombineRgn bentukSha, bentukSha, palangBawah, 2

    BuatHurufSha = bentukSha

    DeleteObject palangVertikal2
    DeleteObject palangVertikal3
    DeleteObject palangBawah
End Function

Public Function BuatHurufShcha() As Long
    Dim bentukShcha As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangVertikal3 As Long, palangBawah As Long, palangKanan As Long

    palangVertikal = BuatKotak(355, 460, 370, 550)
    palangVertikal2 = BuatKotak(392.5, 460, 407.5, 550)
    palangVertikal3 = BuatKotak(430, 460, 445, 550)
    palangBawah = BuatKotak(355, 535, 445, 550)
    palangKanan = BuatKotak(440, 535, 455, 570)

    bentukShcha = palangVertikal

    CombineRgn bentukShcha, bentukShcha, palangVertikal2, 2
    CombineRgn bentukShcha, bentukShcha, palangVertikal3, 2
    CombineRgn bentukShcha, bentukShcha, palangBawah, 2
    CombineRgn bentukShcha, bentukShcha, palangKanan, 2

    BuatHurufShcha = bentukShcha

    DeleteObject palangVertikal2
    DeleteObject palangVertikal3
    DeleteObject palangBawah
    DeleteObject palangKanan
End Function

Public Function BuatHurufYerry() As Long
    Dim bentukYerry As Long
    Dim tiangVertikal As Long, lingkaranBawah As Long, lingkaranDalam As Long, setengahBawah As Long, palangKanan As Long

    palangKanan = BuatKotak(560, 460, 575, 550)
    tiangVertikal = BuatKotak(475, 460, 490, 550)

    'lingkaranBawah = BuatKotak(420, 495, 555, 551, 70, 70)
    'lingkaranDalam = BuatKotak(435, 510, 540, 536, 50, 50)
    setengahBawah = BuatKotak(475, 440, 580, 551)

    bentukYerry = tiangVertikal

    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalam, 3
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1
    CombineRgn lingkaranBawah, lingkaranBawah, palangKanan, 2

    CombineRgn bentukYerry, bentukYerry, lingkaranBawah, 2

    BuatHurufYerry = bentukYerry

    DeleteObject lingkaranBawah
    DeleteObject lingkaranDalam
    DeleteObject setengahBawah
    DeleteObject palangKanan
End Function

Public Function BuatHurufE() As Long
    Dim bentukE As Long
    Dim luar As Long, dalam As Long, samping As Long, palangTengah As Long
    Dim pts(0 To 3) As pointapi

    'luar = BuatKotak(605, 460, 675, 550, 70, 70)
    'dalam = BuatKotak(620, 475, 660, 535, 45, 45)
    palangTengah = BuatKotak(635, 497.5, 670, 512.5)

    pts(0).X = 640: pts(0).Y = 495
    pts(1).X = 575: pts(1).Y = 480
    pts(2).X = 575: pts(2).Y = 535
    pts(3).X = 640: pts(3).Y = 520
    samping = CreatePolygonRgn(pts(0), 4, 2)

    bentukE = luar

    CombineRgn bentukE, bentukE, dalam, 3
    CombineRgn bentukE, bentukE, samping, 4
    CombineRgn bentukE, bentukE, palangTengah, 2

    BuatHurufE = bentukE

    DeleteObject dalam
    DeleteObject samping
    DeleteObject palangTengah
End Function

Public Function BuatHurufYu() As Long
    Dim bentukYu As Long
    Dim luar As Long, dalam As Long, palangKiri As Long, palangTengah As Long

    palangKiri = BuatKotak(705, 455, 720, 550)
    palangTengah = BuatKotak(705, 495, 730, 510)
    'luar = BuatKotak(730, 455, 805, 550, 70, 70)
    'dalam = BuatKotak(745, 470, 790, 535, 45, 45)

    bentukYu = luar

    CombineRgn bentukYu, bentukYu, dalam, 3
    CombineRgn bentukYu, bentukYu, palangKiri, 2
    CombineRgn bentukYu, bentukYu, palangTengah, 2

    BuatHurufYu = bentukYu

    DeleteObject dalam
End Function

Public Function BuatHurufYa() As Long
    Dim bentukYa As Long
    Dim tiangVertikal As Long, kakiKiri As Long, lengkung As Long, lengkungDalam As Long
    Dim pts(0 To 7) As pointapi
    
    '==== Batang kanan vertikal ====
    pts(0).X = 910: pts(0).Y = 440
    pts(1).X = 925: pts(1).Y = 440
    pts(2).X = 925: pts(2).Y = 550
    pts(3).X = 910: pts(3).Y = 550
    tiangVertikal = CreatePolygonRgn(pts(0), 4, 2)
    
    '==== Lengkungan kiri atas polygon ====
    ' Membuat lengkungan manual menggantikan RoundRect

    pts(0).X = 850: pts(0).Y = 440   ' atas kiri
    pts(1).X = 910: pts(1).Y = 440   ' atas kanan (ke tiang)
    pts(2).X = 910: pts(2).Y = 500   ' turun kanan
    pts(3).X = 850: pts(3).Y = 500   ' bawah kanan lengkung
    pts(4).X = 840: pts(4).Y = 490   ' bawah tengah
    pts(5).X = 830: pts(5).Y = 475   ' bawah kiri lengkung
    pts(6).X = 830: pts(6).Y = 465   ' naik kiri
    pts(7).X = 840: pts(7).Y = 450   ' titik mendekati atas
    lengkung = CreatePolygonRgn(pts(0), 8, 2)


    pts(0).X = 860: pts(0).Y = 450   ' atas kiri
    pts(1).X = 910: pts(1).Y = 450   ' atas kanan (ke tiang)
    pts(2).X = 910: pts(2).Y = 490   ' turun kanan
    pts(3).X = 860: pts(3).Y = 490   ' bawah kanan lengkung
    pts(4).X = 850: pts(4).Y = 480   ' bawah tengah
    pts(5).X = 845: pts(5).Y = 475   ' bawah kiri lengkung
    pts(6).X = 845: pts(6).Y = 465   ' naik kiri
    pts(7).X = 850: pts(7).Y = 460   ' titik mendekati atas
    lengkungDalam = CreatePolygonRgn(pts(0), 8, 2)
    
    '==== Kaki kiri ====
    pts(0).X = 870: pts(0).Y = 490
    pts(1).X = 890: pts(1).Y = 490
    pts(2).X = 850: pts(2).Y = 550
    pts(3).X = 830: pts(3).Y = 550
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)

    '==== Combine ====
    bentukYa = tiangVertikal
    
    CombineRgn bentukYa, bentukYa, lengkung, 2
    CombineRgn bentukYa, bentukYa, lengkungDalam, 4
    CombineRgn bentukYa, bentukYa, kakiKiri, 2
    
    BuatHurufYa = bentukYa

    DeleteObject lengkung
    DeleteObject kakiKiri

End Function

