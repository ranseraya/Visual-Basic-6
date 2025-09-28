Attribute VB_Name = "Functions"
Option Explicit
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As pointapi, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Type pointapi
X As Long
Y As Long
End Type

Public bentukAkhir As Long

Public Function BuatHurufA() As Long
    Dim bentukA As Long
    Dim kakiKiri As Long, kakiKanan As Long, palangTengah As Long
    Dim pts(0 To 3) As pointapi
    
    pts(0).X = 130:  pts(0).Y = 100
    pts(1).X = 140:  pts(1).Y = 100
    pts(2).X = 110:  pts(2).Y = 195
    pts(3).X = 100:  pts(3).Y = 195
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)


    pts(0).X = 130:  pts(0).Y = 100
    pts(1).X = 140:  pts(1).Y = 100
    pts(2).X = 170:  pts(2).Y = 195
    pts(3).X = 160:  pts(3).Y = 195
    kakiKanan = CreatePolygonRgn(pts(0), 4, 2)

    palangTengah = CreateRectRgn(120, 155, 158, 165)

    bentukA = kakiKiri
    CombineRgn bentukA, bentukA, kakiKanan, 2
    CombineRgn bentukA, bentukA, palangTengah, 2
    
    DeleteObject kakiKanan
    DeleteObject palangTengah
    
    BuatHurufA = bentukA
End Function

Public Function BuatHurufBe() As Long
    Dim bentukBe As Long
    Dim palangAtas As Long, tiangVertikal As Long, lingkaranBawah As Long, lingkaranDalam As Long, setengahBawah As Long
    
    palangAtas = CreateRectRgn(215, 100, 285, 110)
    tiangVertikal = CreateRectRgn(215, 100, 225, 195)
        
    lingkaranBawah = CreateRoundRectRgn(140, 145, 285, 196, 70, 70)
    lingkaranDalam = CreateRoundRectRgn(150, 155, 275, 186, 50, 50)
    setengahBawah = CreateRectRgn(215, 140, 285, 200)
    
    bentukBe = tiangVertikal
    
    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalam, 3
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1
    
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
    Dim tiangVertikal As Long, lingkaranAtas As Long, lingkaranDalamAtas As Long, lingkaranBawah As Long, lingkaranDalamBawah As Long, setengahBawah As Long
    
    tiangVertikal = CreateRectRgn(330, 100, 340, 195)
        
    lingkaranAtas = CreateRoundRectRgn(265, 100, 400, 153, 70, 70)
    lingkaranDalamAtas = CreateRoundRectRgn(275, 110, 390, 143, 50, 50)
    lingkaranBawah = CreateRoundRectRgn(265, 143, 400, 196, 70, 70)
    lingkaranDalamBawah = CreateRoundRectRgn(275, 153, 390, 186, 50, 50)
    setengahBawah = CreateRectRgn(330, 100, 425, 200)
    
    bentukVe = tiangVertikal
    
    CombineRgn lingkaranAtas, lingkaranAtas, lingkaranDalamAtas, 3
    CombineRgn lingkaranAtas, lingkaranAtas, setengahBawah, 1
    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalamBawah, 3
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1
    
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

    palangAtas = CreateRectRgn(445, 100, 515, 110)
    tiangVertikal = CreateRectRgn(445, 100, 455, 195)
    
    bentukGe = tiangVertikal
    CombineRgn bentukGe, bentukGe, palangAtas, 2

    DeleteObject palangAtas
    BuatHurufGe = bentukGe
End Function

Public Function BuatHurufDe() As Long
    Dim bentukDe As Long
    Dim palangBawah As Long, potonganBawah As Long, bagianAtas As Long, potonganAtas As Long
    Dim bagianKiri As Long, potonganKiri As Long, potonganKiri2 As Long
    
    palangBawah = CreateRectRgn(560, 165, 630, 195)
    potonganBawah = CreateRectRgn(570, 175, 620, 195)
    bentukDe = palangBawah
    
    CombineRgn bentukDe, bentukDe, potonganBawah, 3
    
    bagianAtas = CreateRectRgn(580, 100, 623, 175)
    potonganAtas = CreateRectRgn(580, 110, 613, 175)
    CombineRgn bagianAtas, bagianAtas, potonganAtas, 3
    
    bagianKiri = CreateRoundRectRgn(-50, -200, 590, 220, 200, 200)
    potonganKiri = CreateRoundRectRgn(-50, -200, 580, 220, 200, 200)
    potonganKiri2 = CreateRectRgn(0, 0, 565, 300)
    potonganAtas = CreateRectRgn(560, 0, 630, 100)
    
    CombineRgn bagianKiri, bagianKiri, potonganKiri, 3
    CombineRgn bagianKiri, bagianKiri, potonganKiri2, 4
    CombineRgn bagianKiri, bagianKiri, potonganAtas, 4
    
    CombineRgn bentukDe, bentukDe, bagianAtas, 2
    CombineRgn bentukDe, bentukDe, bagianKiri, 2

    BuatHurufDe = bentukDe
    DeleteObject potonganBawah
    DeleteObject bagianAtas
    DeleteObject potonganAtas
    DeleteObject bagianKiri
    DeleteObject potonganKiri
    DeleteObject potonganKiri2
End Function

Public Function BuatHurufYe() As Long
    Dim bentukYe As Long
    Dim palangAtas As Long, palangTengah As Long, palangBawah As Long, tiangVertikal As Long
    
    palangAtas = CreateRectRgn(675, 100, 745, 110)
    palangTengah = CreateRectRgn(675, 142, 745, 152)
    palangBawah = CreateRectRgn(675, 185, 745, 195)
    
    tiangVertikal = CreateRectRgn(675, 100, 685, 195)
        
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
    
    palangAtas = CreateRectRgn(790, 100, 860, 110)
    palangTengah = CreateRectRgn(790, 142, 860, 152)
    palangBawah = CreateRectRgn(790, 185, 860, 195)
    mata = CreateEllipticRgn(805, 75, 820, 90)
    mata2 = CreateEllipticRgn(835, 75, 850, 90)
    
    tiangVertikal = CreateRectRgn(790, 100, 800, 195)
        
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
    garisCerminX = 940
    
    bagianAtasKanan = CreateEllipticRgn(920, 100, 970, 150)
    potonganKanan1 = CreateEllipticRgn(930, 110, 960, 140)
    potonganKanan2 = CreateRectRgn(935, 125, 970, 150)
    bagianAtasKanan2 = CreateEllipticRgn(960, 100, 1010, 150)
    potonganKanan3 = CreateEllipticRgn(970, 110, 1000, 140)
    potonganKanan4 = CreateRectRgn(960, 100, 985, 125)
    
    ptsKanan(0).X = 945: ptsKanan(0).Y = 145
    ptsKanan(1).X = 955: ptsKanan(1).Y = 145
    ptsKanan(2).X = 985: ptsKanan(2).Y = 195
    ptsKanan(3).X = 975: ptsKanan(3).Y = 195
    kakiKanan = CreatePolygonRgn(ptsKanan(0), 4, 2)
    
    bagianAtasKiri = CreateEllipticRgn((2 * garisCerminX) - 970, 100, (2 * garisCerminX) - 920, 150)
    potonganKiri1 = CreateEllipticRgn((2 * garisCerminX) - 960, 110, (2 * garisCerminX) - 930, 140)
    potonganKiri2 = CreateRectRgn((2 * garisCerminX) - 970, 125, (2 * garisCerminX) - 935, 150)
    bagianAtasKiri2 = CreateEllipticRgn((2 * garisCerminX) - 1010, 100, (2 * garisCerminX) - 960, 150)
    potonganKiri3 = CreateEllipticRgn((2 * garisCerminX) - 1000, 110, (2 * garisCerminX) - 970, 140)
    potonganKiri4 = CreateRectRgn((2 * garisCerminX) - 985, 100, (2 * garisCerminX) - 960, 125)
    
    For i = 0 To 3
        ptsKiri(i).X = (2 * garisCerminX) - ptsKanan(i).X
        ptsKiri(i).Y = ptsKanan(i).Y
    Next i
    kakiKiri = CreatePolygonRgn(ptsKiri(0), 4, 2)
    
    tiangVertikal = CreateRectRgn(935, 100, 945, 195)
    palangTengah = CreateRectRgn(925, 140, 945, 150)
    
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
        
    lingkaranAtas = CreateEllipticRgn(1020, 100, 1090, 152.5)
    potonganAtas = CreateEllipticRgn(1030, 110, 1080, 142.5)
    CombineRgn lingkaranAtas, lingkaranAtas, potonganAtas, 3
    
    lingkaranBawah = CreateEllipticRgn(1020, 142.5, 1090, 195)
    potonganBawah = CreateEllipticRgn(1030, 152.5, 1080, 185)
    CombineRgn lingkaranBawah, lingkaranBawah, potonganBawah, 3
    
    pts(0).X = 1000: pts(0).Y = 120
    pts(1).X = 1055: pts(1).Y = 130
    pts(2).X = 1055: pts(2).Y = 165
    pts(3).X = 1000: pts(3).Y = 175
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

    palangVertikal = CreateRectRgn(1135, 100, 1145, 195)
    palangVertikal2 = CreateRectRgn(1195, 100, 1205, 195)
    
    pts(0).X = 1195:  pts(0).Y = 100
    pts(1).X = 1205: pts(1).Y = 100
    pts(2).X = 1145:  pts(2).Y = 195
    pts(3).X = 1135:  pts(3).Y = 195
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

    palangVertikal = CreateRectRgn(1250, 100, 1260, 195)
    palangVertikal2 = CreateRectRgn(1310, 100, 1320, 195)
    lengkungan = CreateEllipticRgn(1265, 55, 1305, 95)
    lengkunganDalam = CreateEllipticRgn(1275, 65, 1295, 85)
    potonganLengkung = CreateRectRgn(1250, 55, 1320, 75)
    
    pts(0).X = 1310:  pts(0).Y = 100
    pts(1).X = 1320: pts(1).Y = 100
    pts(2).X = 1260:  pts(2).Y = 195
    pts(3).X = 1250:  pts(3).Y = 195
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
    
    bagianAtasKanan = CreateEllipticRgn(1375, 100, 1425, 150)
    potonganKanan1 = CreateEllipticRgn(1385, 110, 1415, 140)
    potonganKanan2 = CreateRectRgn(1400, 125, 1425, 150)
    bagianAtasKanan2 = CreateEllipticRgn(1415, 100, 1465, 150)
    potonganKanan3 = CreateEllipticRgn(1425, 110, 1455, 140)
    potonganKanan4 = CreateRectRgn(1415, 100, 1440, 125)
    
    ptsKanan(0).X = 1385: ptsKanan(0).Y = 145
    ptsKanan(1).X = 1400: ptsKanan(1).Y = 145
    ptsKanan(2).X = 1445: ptsKanan(2).Y = 195
    ptsKanan(3).X = 1430: ptsKanan(3).Y = 195
    kakiKanan = CreatePolygonRgn(ptsKanan(0), 4, 2)
    
    tiangVertikal = CreateRectRgn(1375, 100, 1385, 195)
    palangTengah = CreateRectRgn(1375, 140, 1400, 150)
    
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
    
    palangVertikal = CreateRectRgn(100, 230, 110, 270)
    palangVertikal2 = CreateRectRgn(160, 230, 170, 325)
    palangHorizontal = CreateRectRgn(100, 230, 170, 240)
    palangKiri = CreateRoundRectRgn(-20, 240, 111, 325, 40, 40)
    potonganAtas = CreateRoundRectRgn(-50, 150, 101, 315, 30, 30)
    potonganKiri = CreateRectRgn(0, 30, 80, 335)
    
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
    
    palangVertikal = CreateRectRgn(275, 230, 285, 315)
    palangVertikal2 = CreateRectRgn(215, 230, 225, 315)
    
    ptsKanan(0).X = 220: ptsKanan(0).Y = 230
    ptsKanan(1).X = 230: ptsKanan(1).Y = 230
    ptsKanan(2).X = 255: ptsKanan(2).Y = 315
    ptsKanan(3).X = 245: ptsKanan(3).Y = 315
    palangDiagonal = CreatePolygonRgn(ptsKanan(0), 4, 2)
    
    ptsKiri(0).X = 270: ptsKiri(0).Y = 230
    ptsKiri(1).X = 280: ptsKiri(1).Y = 230
    ptsKiri(2).X = 255: ptsKiri(2).Y = 315
    ptsKiri(3).X = 245: ptsKiri(3).Y = 315
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

    palangVertikal = CreateRectRgn(330, 230, 340, 315)
    palangVertikal2 = CreateRectRgn(390, 230, 400, 315)
    palangTengah = CreateRectRgn(330, 265, 400, 275)
        
    bentukEn = palangVertikal
    
    CombineRgn bentukEn, bentukEn, palangVertikal2, 2
    CombineRgn bentukEn, bentukEn, palangTengah, 2
    
    BuatHurufEn = bentukEn
    
    DeleteObject palangVertikal2
    DeleteObject palangTengah
End Function

Public Function BuatHurufO() As Long
    Dim bentukO As Long
    Dim luar As Long, dalam As Long
    
    luar = CreateRoundRectRgn(445, 230, 515, 315, 70, 70)
    dalam = CreateRoundRectRgn(460, 245, 500, 300, 45, 45)
    
    bentukO = luar
    
    CombineRgn bentukO, bentukO, dalam, 3
    
    BuatHurufO = bentukO
    
    DeleteObject dalam
End Function

Public Function BuatHurufPe() As Long
    Dim bentukPe As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangAtas As Long

    palangVertikal = CreateRectRgn(560, 230, 570, 315)
    palangVertikal2 = CreateRectRgn(620, 230, 630, 315)
    palangAtas = CreateRectRgn(560, 230, 630, 240)
    
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
    
    tiangVertikal = CreateRectRgn(675, 230, 685, 315)
        
    lingkaranAtas = CreateRoundRectRgn(600, 230, 745, 281, 70, 70)
    lingkaranDalam = CreateRoundRectRgn(610, 240, 735, 271, 50, 50)
    setengahAtas = CreateRectRgn(675, 230, 745, 315)
    
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
    
    luar = CreateRoundRectRgn(790, 230, 860, 315, 70, 70)
    dalam = CreateRoundRectRgn(805, 245, 845, 300, 45, 45)
    
    pts(0).X = 825: pts(0).Y = 260
    pts(1).X = 880: pts(1).Y = 250
    pts(2).X = 880: pts(2).Y = 305
    pts(3).X = 825: pts(3).Y = 275
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

    palangVertikal = CreateRectRgn(935, 230, 945, 315)
    palangAtas = CreateRectRgn(905, 230, 975, 240)
    
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
    
    pts(0).X = 1020:  pts(0).Y = 230
    pts(1).X = 1030:  pts(1).Y = 230
    pts(2).X = 1060:  pts(2).Y = 285
    pts(3).X = 1050:  pts(3).Y = 285
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)

    pts(0).X = 1046:  pts(0).Y = 295
    pts(1).X = 1056:  pts(1).Y = 295
    pts(2).X = 1090:  pts(2).Y = 230
    pts(3).X = 1080:  pts(3).Y = 230
    kakiKanan = CreatePolygonRgn(pts(0), 4, 2)
    
    palangKiri = CreateRoundRectRgn(1000, 270, 1060, 315, 50, 50)
    potonganAtas = CreateRoundRectRgn(1000, 260, 1050, 305, 25, 25)
    potonganKiri = CreateRectRgn(900, 285, 1030, 355)
    
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
    
    luar = CreateRoundRectRgn(1135, 250, 1205, 305, 50, 50)
    dalam = CreateRoundRectRgn(1145, 260, 1195, 295, 30, 30)
    palangTengah = CreateRectRgn(1165, 245, 1175, 310)
    
    bentukEf = luar
    
    CombineRgn bentukEf, bentukEf, dalam, 3
    CombineRgn bentukEf, bentukEf, palangTengah, 2
    
    BuatHurufEf = bentukEf
End Function

Public Function BuatHurufKha() As Long
    Dim bentukKha As Long
    Dim kakiKiri As Long, kakiKanan As Long
    Dim pts(0 To 3) As pointapi
    
    pts(0).X = 1310:  pts(0).Y = 230
    pts(1).X = 1320:  pts(1).Y = 230
    pts(2).X = 1260:  pts(2).Y = 315
    pts(3).X = 1250:  pts(3).Y = 315
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)


    pts(0).X = 1310:  pts(0).Y = 315
    pts(1).X = 1320:  pts(1).Y = 315
    pts(2).X = 1260:  pts(2).Y = 230
    pts(3).X = 1250:  pts(3).Y = 230
    kakiKanan = CreatePolygonRgn(pts(0), 4, 2)

    bentukKha = kakiKiri
    CombineRgn bentukKha, bentukKha, kakiKanan, 2
    
    BuatHurufKha = bentukKha
    
    DeleteObject kakiKanan
End Function

Public Function BuatHurufTse() As Long
    Dim bentukTse As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangBawah As Long, palangKanan As Long

    palangVertikal = CreateRectRgn(1375, 230, 1385, 295)
    palangVertikal2 = CreateRectRgn(1420, 230, 1430, 295)
    palangBawah = CreateRectRgn(1375, 305, 1435, 295)
    palangKanan = CreateRectRgn(1425, 295, 1435, 315)
    
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
    
    lingkaranAtas = CreateRoundRectRgn(100, 200, 170, 410, 55, 55)
    lingkaranDalam = CreateRoundRectRgn(115, 215, 165, 395, 40, 40)
    setengahAtas = CreateRectRgn(100, 200, 200, 350)
    palangVertikal = CreateRectRgn(155, 350, 170, 445)
    
    bentukChe = lingkaranAtas
    
    CombineRgn bentukChe, bentukChe, lingkaranDalam, 4
    CombineRgn bentukChe, bentukChe, setengahAtas, 4
    CombineRgn bentukChe, bentukChe, palangVertikal, 2
    
  BuatHurufChe = bentukChe
End Function

Public Function BuatHurufSha() As Long
    Dim bentukSha As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangVertikal3 As Long, palangBawah As Long

    palangVertikal = CreateRectRgn(215, 350, 225, 435)
    palangVertikal2 = CreateRectRgn(255, 350, 265, 435)
    palangVertikal3 = CreateRectRgn(295, 350, 305, 435)
    palangBawah = CreateRectRgn(215, 425, 305, 435)
    
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

    palangVertikal = CreateRectRgn(355, 350, 365, 435)
    palangVertikal2 = CreateRectRgn(395, 350, 405, 435)
    palangVertikal3 = CreateRectRgn(435, 350, 445, 435)
    palangBawah = CreateRectRgn(355, 425, 445, 435)
    palangKanan = CreateRectRgn(440, 425, 450, 445)
    
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
    
    palangKanan = CreateRectRgn(575, 350, 585, 445)
    tiangVertikal = CreateRectRgn(500, 350, 510, 445)
        
    lingkaranBawah = CreateRoundRectRgn(425, 395, 570, 446, 70, 70)
    lingkaranDalam = CreateRoundRectRgn(435, 405, 560, 436, 50, 50)
    setengahBawah = CreateRectRgn(500, 395, 580, 446)
    
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
    
    luar = CreateRoundRectRgn(635, 350, 705, 445, 70, 70)
    dalam = CreateRoundRectRgn(650, 365, 690, 430, 45, 45)
    palangTengah = CreateRectRgn(665, 390, 700, 405)
    
    pts(0).X = 670: pts(0).Y = 385
    pts(1).X = 605: pts(1).Y = 370
    pts(2).X = 605: pts(2).Y = 425
    pts(3).X = 670: pts(3).Y = 410
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
    
    palangKiri = CreateRectRgn(755, 350, 770, 445)
    palangTengah = CreateRectRgn(755, 390, 780, 405)
    luar = CreateRoundRectRgn(780, 350, 855, 445, 70, 70)
    dalam = CreateRoundRectRgn(795, 365, 840, 430, 45, 45)
    
    bentukYu = luar
    
    CombineRgn bentukYu, bentukYu, dalam, 3
    CombineRgn bentukYu, bentukYu, palangKiri, 2
    CombineRgn bentukYu, bentukYu, palangTengah, 2
    
    BuatHurufYu = bentukYu
    
    DeleteObject dalam
End Function

Public Function BuatHurufYa() As Long
    Dim bentukYa As Long
    Dim tiangVertikal As Long, lingkaranBawah As Long, lingkaranDalam As Long, setengahBawah As Long, kakiKiri As Long
    Dim pts(0 To 3) As pointapi
    
    tiangVertikal = CreateRectRgn(965, 350, 975, 445)
    
    lingkaranBawah = CreateRoundRectRgn(905, 350, 1050, 401, 70, 70)
    lingkaranDalam = CreateRoundRectRgn(915, 360, 1040, 391, 50, 50)
    setengahBawah = CreateRectRgn(905, 350, 975, 401)
    
    pts(0).X = 930: pts(0).Y = 399
    pts(1).X = 940: pts(1).Y = 399
    pts(2).X = 915: pts(2).Y = 445
    pts(3).X = 905: pts(3).Y = 445
    kakiKiri = CreatePolygonRgn(pts(0), 4, 2)
    bentukYa = tiangVertikal
    
    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalam, 3
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1
    CombineRgn lingkaranBawah, lingkaranBawah, kakiKiri, 2
    
    CombineRgn bentukYa, bentukYa, lingkaranBawah, 2
    
    BuatHurufYa = bentukYa
    
    DeleteObject lingkaranBawah
    DeleteObject lingkaranDalam
    DeleteObject setengahBawah
    DeleteObject kakiKiri
End Function
