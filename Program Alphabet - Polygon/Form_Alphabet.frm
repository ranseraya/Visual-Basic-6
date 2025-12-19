VERSION 5.00
Begin VB.Form Form_Alphabet 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   23760
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form_Alphabet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    Dim pts(0 To 3) As pointapi
    pts(0).X = 0: pts(0).Y = 0
    pts(1).X = 1: pts(1).Y = 0
    pts(2).X = 0: pts(2).Y = 1
    pts(3).X = 1: pts(3).Y = 1
       
    bentukAkhir = CreatePolygonRgn(pts(0), 4, 2)
    
    hurufA = BuatHurufA()
    hurufBe = BuatHurufBe()
    hurufVe = BuatHurufVe()
    hurufGe = BuatHurufGe()
    hurufDe = BuatHurufDe()
    hurufYe = BuatHurufYe()
    hurufYo = BuatHurufYo()
    hurufZhe = BuatHurufZhe()
    hurufZe = BuatHurufZe()
    hurufI = BuatHurufI()
    hurufIy = BuatHurufIy()
    hurufKa = BuatHurufKa()
    hurufEl = BuatHurufEl()
    hurufEm = BuatHurufEm()
    hurufEn = BuatHurufEn()
    hurufO = BuatHurufO()
    hurufPe = BuatHurufPe()
    hurufTe = BuatHurufTe()
    hurufKha = BuatHurufKha()
    hurufTse = BuatHurufTse()
    hurufSha = BuatHurufSha()
    hurufShcha = BuatHurufShcha()
    hurufYa = BuatHurufYa()
    
    Call OffsetRgn(hurufKha, -230, 0)
    Call OffsetRgn(hurufTse, -230, 0)
    Call OffsetRgn(hurufSha, 480, -150)
    Call OffsetRgn(hurufShcha, 480, -150)
        
    CombineRgn bentukAkhir, bentukAkhir, hurufA, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufBe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufVe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufGe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufDe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYo, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufZhe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufZe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufI, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufIy, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufKa, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEl, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEm, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEn, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufO, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufPe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufTe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufKha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufTse, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufSha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufShcha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYa, 2
    
    SetWindowRgn Me.hwnd, bentukAkhir, True
    
    DeleteObject hurufA
    DeleteObject hurufBe
    DeleteObject hurufVe
    DeleteObject hurufGe
    DeleteObject hurufDe
    DeleteObject bentukYe
    DeleteObject bentukYo
    DeleteObject bentukZhe
    DeleteObject bentukZe
    DeleteObject bentukI
    DeleteObject bentukIy
    DeleteObject bentukKa
    DeleteObject bentukEl
    DeleteObject bentukEm
    DeleteObject bentukEn
    DeleteObject bentukO
    DeleteObject bentukPe
    DeleteObject bentukTe
    DeleteObject bentukKha
    DeleteObject bentukTse
    DeleteObject bentukSha
    DeleteObject bentukShcha
    'DeleteObject bentukYa
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub
