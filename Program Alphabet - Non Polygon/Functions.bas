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
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal nXOffset As Long, ByVal nYOffset As Long) As Long

Type pointapi
X As Long
Y As Long
End Type

Public bentukAkhir As Long


Public Function BuatHurufA() As Long
    Dim bentukA As Long, palangTengah As Long
    Dim kk1 As Long, kk2 As Long, kk3 As Long, kk4 As Long, kk5 As Long, kk6 As Long, kk7 As Long
    Dim kr1 As Long, kr2 As Long, kr3 As Long, kr4 As Long, kr5 As Long, kr6 As Long, kr7 As Long
      
    kk1 = CreateRectRgn(120, 230, 130, 250)
    kk2 = CreateRectRgn(125, 215, 135, 235)
    kk3 = CreateRectRgn(130, 200, 140, 220)
    kk4 = CreateRectRgn(135, 185, 145, 205)
    kk5 = CreateRectRgn(140, 170, 150, 190)
    kk6 = CreateRectRgn(145, 155, 155, 175)
    kk7 = CreateRectRgn(150, 140, 160, 160)
    
    kr1 = CreateRectRgn(185, 230, 195, 250)
    kr2 = CreateRectRgn(180, 215, 190, 235)
    kr3 = CreateRectRgn(175, 200, 185, 220)
    kr4 = CreateRectRgn(170, 185, 180, 205)
    kr5 = CreateRectRgn(165, 170, 175, 190)
    kr6 = CreateRectRgn(160, 155, 170, 175)
    kr7 = CreateRectRgn(155, 140, 165, 160)
    
    palangTengah = CreateRectRgn(140, 200, 175, 215)
            
    bentukA = kk1
    CombineRgn bentukA, bentukA, kk2, 2
    CombineRgn bentukA, bentukA, kk3, 2
    CombineRgn bentukA, bentukA, kk4, 2
    CombineRgn bentukA, bentukA, kk5, 2
    CombineRgn bentukA, bentukA, kk6, 2
    CombineRgn bentukA, bentukA, kk7, 2

    CombineRgn bentukA, bentukA, kr1, 2
    CombineRgn bentukA, bentukA, kr2, 2
    CombineRgn bentukA, bentukA, kr3, 2
    CombineRgn bentukA, bentukA, kr4, 2
    CombineRgn bentukA, bentukA, kr5, 2
    CombineRgn bentukA, bentukA, kr6, 2
    CombineRgn bentukA, bentukA, kr7, 2
    
    CombineRgn bentukA, bentukA, palangTengah, 2
    
    BuatHurufA = bentukA
    DeleteObject palangTengah
    
End Function


Public Function BuatHurufBe() As Long
    Dim bentukBe As Long
    Dim palangAtas As Long, tiangVertikal As Long, lingkaranBawah As Long, lingkaranDalam As Long, setengahBawah As Long
    
    palangAtas = CreateRectRgn(235, 140, 325, 155)
    tiangVertikal = CreateRectRgn(235, 140, 250, 250)
        
    lingkaranBawah = CreateRoundRectRgn(130, 185, 325, 251, 60, 60)
    lingkaranDalam = CreateRoundRectRgn(145, 200, 310, 236, 40, 40)
    setengahBawah = CreateRectRgn(235, 140, 325, 251)
    
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
    
    tiangVertikal = CreateRectRgn(355, 140, 370, 250)
        
    lingkaranAtas = CreateRoundRectRgn(225, 140, 445, 203, 70, 70)
    lingkaranDalamAtas = CreateRoundRectRgn(240, 155, 430, 188, 50, 50)
    lingkaranBawah = CreateRoundRectRgn(225, 188, 445, 251, 70, 70)
    lingkaranDalamBawah = CreateRoundRectRgn(240, 203, 430, 236, 50, 50)
    setengahBawah = CreateRectRgn(355, 140, 445, 250)
    
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

    palangAtas = CreateRectRgn(475, 140, 565, 155)
    tiangVertikal = CreateRectRgn(475, 140, 490, 250)
    
    bentukGe = tiangVertikal
    CombineRgn bentukGe, bentukGe, palangAtas, 2

    DeleteObject palangAtas
    BuatHurufGe = bentukGe
End Function

Public Function BuatHurufDe() As Long
    Dim bentukDe As Long
    Dim palangBawah As Long, potonganBawah As Long, bagianAtas As Long, potonganAtas As Long
    Dim bagianKiri As Long, potonganKiri As Long, potonganKiri2 As Long
    
    palangBawah = CreateRectRgn(595, 215, 685, 250)
    potonganBawah = CreateRectRgn(610, 230, 670, 250)
    bentukDe = palangBawah
    
    CombineRgn bentukDe, bentukDe, potonganBawah, 3
    
    bagianAtas = CreateRectRgn(615, 140, 678, 215)
    potonganAtas = CreateRectRgn(615, 155, 663, 215)
    CombineRgn bagianAtas, bagianAtas, potonganAtas, 3
    
    bagianKiri = CreateRoundRectRgn(-50, -200, 630, 265, 200, 200)
    potonganKiri = CreateRoundRectRgn(-50, -200, 615, 265, 200, 200)
    potonganKiri2 = CreateRectRgn(0, 0, 600, 300)
    potonganAtas = CreateRectRgn(595, 0, 665, 140)
    
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
    
    palangAtas = CreateRectRgn(725, 140, 795, 155)
    palangTengah = CreateRectRgn(725, 185, 795, 200)
    palangBawah = CreateRectRgn(725, 235, 795, 250)
    
    tiangVertikal = CreateRectRgn(725, 140, 740, 250)
        
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
    
    palangAtas = CreateRectRgn(845, 140, 915, 155)
    palangTengah = CreateRectRgn(845, 185, 915, 200)
    palangBawah = CreateRectRgn(845, 235, 915, 250)
    mata = CreateEllipticRgn(860, 115, 875, 130)
    mata2 = CreateEllipticRgn(885, 115, 900, 130)
    
    tiangVertikal = CreateRectRgn(845, 140, 860, 250)
        
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
    Dim bagianAtasKanan As Long, potonganAtasKanan1 As Long, potonganAtasKanan2, bagianAtasKanan2, potonganAtasKanan3, potonganAtasKanan4
    Dim bagianAtasKiri As Long, potonganAtasKiri1 As Long, potonganAtasKiri2, bagianAtasKiri2, potonganAtasKiri3, potonganAtasKiri4
    Dim bagianBawahKanan As Long, potonganBawahKanan1 As Long, potonganBawahKanan2, bagianBawahKanan2, potonganBawahKanan3, potonganBawahKanan4
    Dim bagianBawahKiri As Long, potonganBawahKiri1 As Long, potonganBawahKiri2, bagianBawahKiri2, potonganBawahKiri3, potonganBawahKiri4
    Dim tiangVertikal As Long, palangTengah As Long, kakiKanan As Long, kakiKiri As Long

    Dim ptsKanan(0 To 3) As pointapi, ptsKiri(0 To 3) As pointapi
    Dim i As Integer, garisCerminX As Long
    garisCerminX = 1000
    
    bagianAtasKanan = CreateEllipticRgn(992.5, 140, 1052.5, 205)
    potonganAtasKanan1 = CreateEllipticRgn(1007.5, 155, 1037.5, 190)
    potonganAtasKanan2 = CreateRectRgn(1022.5, 170, 1052.5, 205)
    bagianAtasKanan2 = CreateEllipticRgn(1037.5, 140, 1102.5, 205)
    potonganAtasKanan3 = CreateEllipticRgn(1052.5, 155, 1087.5, 190)
    potonganAtasKanan4 = CreateRectRgn(1037.5, 140, 1067.5, 170)
    
    bagianAtasKiri = CreateEllipticRgn((2 * garisCerminX) - 1052.5, 140, (2 * garisCerminX) - 992.5, 205)
    potonganAtasKiri1 = CreateEllipticRgn((2 * garisCerminX) - 1037.5, 155, (2 * garisCerminX) - 1007.5, 190)
    potonganAtasKiri2 = CreateRectRgn((2 * garisCerminX) - 1052.5, 170, (2 * garisCerminX) - 1022.5, 205)
    bagianAtasKiri2 = CreateEllipticRgn((2 * garisCerminX) - 1102.5, 140, (2 * garisCerminX) - 1037.5, 205)
    potonganAtasKiri3 = CreateEllipticRgn((2 * garisCerminX) - 1087.5, 155, (2 * garisCerminX) - 1052.5, 190)
    potonganAtasKiri4 = CreateRectRgn((2 * garisCerminX) - 1067.5, 140, (2 * garisCerminX) - 1037.5, 170)
    
    bagianBawahKanan = CreateEllipticRgn(992.5, 190, 1052.5, 255)
    potonganBawahKanan1 = CreateEllipticRgn(1007.5, 205, 1037.5, 240)
    potonganBawahKanan2 = CreateRectRgn(1022.5, 190, 1052.5, 230)
    bagianBawahKanan2 = CreateEllipticRgn(1037.5, 190, 1102.5, 255)
    potonganBawahKanan3 = CreateEllipticRgn(1052.5, 205, 1087.5, 240)
    potonganBawahKanan4 = CreateRectRgn(1037.5, 230, 1067.5, 255)
    
    bagianBawahKiri = CreateEllipticRgn((2 * garisCerminX) - 1052.5, 190, (2 * garisCerminX) - 992.5, 255)
    potonganBawahKiri1 = CreateEllipticRgn((2 * garisCerminX) - 1037.5, 205, (2 * garisCerminX) - 1007.5, 240)
    potonganBawahKiri2 = CreateRectRgn((2 * garisCerminX) - 1052.5, 190, (2 * garisCerminX) - 1022.5, 230)
    bagianBawahKiri2 = CreateEllipticRgn((2 * garisCerminX) - 1102.5, 190, (2 * garisCerminX) - 1037.5, 255)
    potonganBawahKiri3 = CreateEllipticRgn((2 * garisCerminX) - 1087.5, 205, (2 * garisCerminX) - 1052.5, 240)
    potonganBawahKiri4 = CreateRectRgn((2 * garisCerminX) - 1067.5, 230, (2 * garisCerminX) - 1037.5, 255)

    tiangVertikal = CreateRectRgn(992.5, 140, 1007.5, 250)
    palangTengah = CreateRectRgn(977.5, 190, 1022.5, 205)
    
    bentukZhe = tiangVertikal
    
    CombineRgn bagianAtasKanan, bagianAtasKanan, potonganAtasKanan1, 3
    CombineRgn bagianAtasKanan, bagianAtasKanan, potonganAtasKanan2, 1
    CombineRgn bagianAtasKanan2, bagianAtasKanan2, potonganAtasKanan3, 3
    CombineRgn bagianAtasKanan2, bagianAtasKanan2, potonganAtasKanan4, 1
    
    CombineRgn bagianAtasKiri, bagianAtasKiri, potonganAtasKiri1, 3
    CombineRgn bagianAtasKiri, bagianAtasKiri, potonganAtasKiri2, 1
    CombineRgn bagianAtasKiri2, bagianAtasKiri2, potonganAtasKiri3, 3
    CombineRgn bagianAtasKiri2, bagianAtasKiri2, potonganAtasKiri4, 1
    
    CombineRgn bagianAtasKanan, bagianAtasKanan, bagianAtasKanan2, 2
    CombineRgn bagianAtasKiri, bagianAtasKiri, bagianAtasKiri2, 2

    CombineRgn bagianBawahKanan, bagianBawahKanan, potonganBawahKanan1, 3
    CombineRgn bagianBawahKanan, bagianBawahKanan, potonganBawahKanan2, 1
    CombineRgn bagianBawahKanan2, bagianBawahKanan2, potonganBawahKanan3, 3
    CombineRgn bagianBawahKanan2, bagianBawahKanan2, potonganBawahKanan4, 1
    
    CombineRgn bagianBawahKiri, bagianBawahKiri, potonganBawahKiri1, 3
    CombineRgn bagianBawahKiri, bagianBawahKiri, potonganBawahKiri2, 1
    CombineRgn bagianBawahKiri2, bagianBawahKiri2, potonganBawahKiri3, 3
    CombineRgn bagianBawahKiri2, bagianBawahKiri2, potonganBawahKiri4, 1
    
    
    CombineRgn bagianBawahKanan, bagianBawahKanan, bagianBawahKanan2, 2
    CombineRgn bagianBawahKiri, bagianBawahKiri, bagianBawahKiri2, 2

    CombineRgn bentukZhe, bentukZhe, bagianAtasKanan, 2
    CombineRgn bentukZhe, bentukZhe, bagianAtasKiri, 2
    CombineRgn bentukZhe, bentukZhe, bagianBawahKanan, 2
    CombineRgn bentukZhe, bentukZhe, bagianBawahKiri, 2

    CombineRgn bentukZhe, bentukZhe, palangTengah, 2
    
    BuatHurufZhe = bentukZhe
    
    DeleteObject bagianAtasKanan
    DeleteObject bagianAtasKanan2
    DeleteObject potonganAtasKanan1
    DeleteObject potonganAtasKanan2
    DeleteObject potonganAtasKanan3
    DeleteObject potonganAtasKanan4
    DeleteObject bagianAtasKiri
    DeleteObject bagianAtasKiri2
    DeleteObject potonganAtasKiri1
    DeleteObject potonganAtasKiri2
    DeleteObject potonganAtasKiri3
    DeleteObject potonganAtasKiri4
    DeleteObject bagianBawahKanan
    DeleteObject bagianBawahKanan2
    DeleteObject potonganBawahKanan1
    DeleteObject potonganBawahKanan2
    DeleteObject potonganBawahKanan3
    DeleteObject potonganBawahKanan4
    DeleteObject bagianBawahKiri
    DeleteObject bagianBawahKiri2
    DeleteObject potonganBawahKiri1
    DeleteObject potonganBawahKiri2
    DeleteObject potonganBawahKiri3
    DeleteObject potonganBawahKiri4
    DeleteObject palangTengah
    DeleteObject garisCerminX
End Function

Public Function BuatHurufEl() As Long
    Dim bentukEl As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangHorizontal As Long
    Dim palangKiri As Long, potonganAtas As Long, potonganKiri As Long
    
    palangVertikal = CreateRectRgn(135, 300, 150, 380)
    palangVertikal2 = CreateRectRgn(190, 300, 205, 400)
    palangHorizontal = CreateRectRgn(135, 300, 205, 315)
    palangKiri = CreateRoundRectRgn(-20, 340, 151, 400, 50, 50)
    potonganAtas = CreateRoundRectRgn(-50, 180, 136, 385, 30, 30)
    potonganKiri = CreateRectRgn(0, 230, 115, 400)
    
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

Public Function BuatHurufEn() As Long
    Dim bentukEn As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangTengah As Long

    palangVertikal = CreateRectRgn(365, 310, 380, 400)
    palangVertikal2 = CreateRectRgn(420, 310, 435, 400)
    palangTengah = CreateRectRgn(365, 347.5, 435, 362.5)
        
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
    
    luar = CreateRoundRectRgn(475, 290, 565, 400, 90, 90)
    dalam = CreateRoundRectRgn(490, 305, 550, 385, 60, 60)
    
    bentukO = luar
    
    CombineRgn bentukO, bentukO, dalam, 3
    
    BuatHurufO = bentukO
    
    DeleteObject dalam
End Function

Public Function BuatHurufPe() As Long
    Dim bentukPe As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangAtas As Long

    palangVertikal = CreateRectRgn(605, 310, 620, 400)
    palangVertikal2 = CreateRectRgn(660, 310, 675, 400)
    palangAtas = CreateRectRgn(605, 310, 675, 325)
    
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
    
    tiangVertikal = CreateRectRgn(715, 310, 730, 400)
        
    lingkaranAtas = CreateRoundRectRgn(610, 300, 795, 361, 60, 60)
    lingkaranDalam = CreateRoundRectRgn(625, 315, 780, 346, 40, 40)
    setengahAtas = CreateRectRgn(715, 300, 805, 400)
    
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
    
    luar = CreateRoundRectRgn(835, 290, 925, 400, 90, 90)
    dalam = CreateRoundRectRgn(850, 305, 910, 385, 60, 60)
    
    pts(0).X = 875: pts(0).Y = 330
    pts(1).X = 940: pts(1).Y = 315
    pts(2).X = 940: pts(2).Y = 375
    pts(3).X = 875: pts(3).Y = 360
    
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

    palangVertikal = CreateRectRgn(992.5, 290, 1007.5, 400)
    palangAtas = CreateRectRgn(955, 290, 1045, 305)
    
    bentukTe = palangVertikal

    CombineRgn bentukTe, bentukTe, palangAtas, 2
    
    BuatHurufTe = bentukTe

    DeleteObject palangAtas
    BuatHurufTe = bentukTe
End Function

Public Function BuatHurufEf() As Long
    Dim bentukEf As Long
    Dim luar As Long, dalam As Long, palangTengah As Long
    
    luar = CreateRoundRectRgn(1195, 320, 1285, 390, 60, 60)
    dalam = CreateRoundRectRgn(1210, 335, 1270, 375, 40, 40)
    palangTengah = CreateRectRgn(1232.5, 310, 1247.5, 400)
    
    bentukEf = luar
    
    CombineRgn bentukEf, bentukEf, dalam, 3
    CombineRgn bentukEf, bentukEf, palangTengah, 2
    
    BuatHurufEf = bentukEf
End Function

Public Function BuatHurufTse() As Long
    Dim bentukTse As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangBawah As Long, palangKanan As Long

    palangVertikal = CreateRectRgn(1435, 290, 1450, 380)
    palangVertikal2 = CreateRectRgn(1500, 290, 1515, 380)
    palangBawah = CreateRectRgn(1435, 365, 1525, 380)
    palangKanan = CreateRectRgn(1510, 365, 1525, 400)
    
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
    
    lingkaranAtas = CreateRoundRectRgn(125, 310, 195, 520, 55, 55)
    lingkaranDalam = CreateRoundRectRgn(140, 325, 190, 505, 40, 40)
    setengahAtas = CreateRectRgn(125, 250, 195, 460)
    palangVertikal = CreateRectRgn(180, 460, 195, 550)
    
    bentukChe = lingkaranAtas
    
    CombineRgn bentukChe, bentukChe, lingkaranDalam, 4
    CombineRgn bentukChe, bentukChe, setengahAtas, 4
    CombineRgn bentukChe, bentukChe, palangVertikal, 2
    
  BuatHurufChe = bentukChe
End Function

Public Function BuatHurufSha() As Long
    Dim bentukSha As Long
    Dim palangVertikal As Long, palangVertikal2 As Long, palangVertikal3 As Long, palangBawah As Long

    palangVertikal = CreateRectRgn(235, 460, 250, 550)
    palangVertikal2 = CreateRectRgn(272.5, 460, 287.5, 550)
    palangVertikal3 = CreateRectRgn(310, 460, 325, 550)
    palangBawah = CreateRectRgn(235, 535, 325, 550)
    
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

    palangVertikal = CreateRectRgn(355, 460, 370, 550)
    palangVertikal2 = CreateRectRgn(392.5, 460, 407.5, 550)
    palangVertikal3 = CreateRectRgn(430, 460, 445, 550)
    palangBawah = CreateRectRgn(355, 535, 445, 550)
    palangKanan = CreateRectRgn(440, 535, 455, 570)
    
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
    
    palangKanan = CreateRectRgn(560, 460, 575, 550)
    tiangVertikal = CreateRectRgn(475, 460, 490, 550)
        
    lingkaranBawah = CreateRoundRectRgn(420, 495, 555, 551, 70, 70)
    lingkaranDalam = CreateRoundRectRgn(435, 510, 540, 536, 50, 50)
    setengahBawah = CreateRectRgn(475, 440, 580, 551)
    
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

Public Function BuatHurufYu() As Long
    Dim bentukYu As Long
    Dim luar As Long, dalam As Long, palangKiri As Long, palangTengah As Long
    
    palangKiri = CreateRectRgn(705, 455, 720, 550)
    palangTengah = CreateRectRgn(705, 495, 730, 510)
    luar = CreateRoundRectRgn(730, 455, 805, 550, 70, 70)
    dalam = CreateRoundRectRgn(745, 470, 790, 535, 45, 45)
    
    bentukYu = luar
    
    CombineRgn bentukYu, bentukYu, dalam, 3
    CombineRgn bentukYu, bentukYu, palangKiri, 2
    CombineRgn bentukYu, bentukYu, palangTengah, 2
    
    BuatHurufYu = bentukYu
    
    DeleteObject dalam
End Function

Public Function BuatHurufYa() As Long
    Dim bentukYa As Long
    Dim tiangVertikal As Long, lingkaranBawah As Long, lingkaranDalam As Long
    Dim setengahBawah As Long
    Dim kaki1 As Long, kaki2 As Long, kaki3 As Long, kaki4 As Long, kaki5 As Long

    ' Tiang kanan
    tiangVertikal = CreateRectRgn(910, 440, 925, 550)

    ' Lingkaran bagian atas kiri (kurva)
    lingkaranBawah = CreateRoundRectRgn(835, 440, 990, 501, 70, 70)
    lingkaranDalam = CreateRoundRectRgn(850, 455, 975, 486, 50, 50)
    setengahBawah = CreateRectRgn(835, 440, 925, 551)

    ' Kaki kiri dibuat bertingkat (tanpa polygon)
    kaki1 = CreateRectRgn(875, 495, 890, 510)
    kaki2 = CreateRectRgn(865, 505, 880, 520)
    kaki3 = CreateRectRgn(855, 515, 870, 530)
    kaki4 = CreateRectRgn(845, 525, 860, 540)
    kaki5 = CreateRectRgn(835, 535, 850, 550)
    
    ' Buat bentuk lingkaran berlubang
    CombineRgn lingkaranBawah, lingkaranBawah, lingkaranDalam, 4
    CombineRgn lingkaranBawah, lingkaranBawah, setengahBawah, 1

    ' Tambah kaki bertingkat
    CombineRgn lingkaranBawah, lingkaranBawah, kaki1, 2
    CombineRgn lingkaranBawah, lingkaranBawah, kaki2, 2
    CombineRgn lingkaranBawah, lingkaranBawah, kaki3, 2
    CombineRgn lingkaranBawah, lingkaranBawah, kaki4, 2
    CombineRgn lingkaranBawah, lingkaranBawah, kaki5, 2
    
    bentukYa = tiangVertikal
    CombineRgn bentukYa, bentukYa, lingkaranBawah, 2

    BuatHurufYa = bentukYa

    ' Bersihkan
    DeleteObject lingkaranBawah
    DeleteObject lingkaranDalam
    DeleteObject setengahBawah
    DeleteObject kaki1
    DeleteObject kaki2
    DeleteObject kaki3
End Function

