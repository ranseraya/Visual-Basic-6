VERSION 5.00
Begin VB.Form Form_Alphabet 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   23880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   23880
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
    
    bentukAkhir = CreateRectRgn(0, 0, 1, 1)
    
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
    hurufEr = BuatHurufEr()
    hurufEs = BuatHurufEs()
    hurufTe = BuatHurufTe()
    hurufU = BuatHurufU()
    hurufEf = BuatHurufEf()
    hurufKha = BuatHurufKha()
    hurufTse = BuatHurufTse()
    hurufChe = BuatHurufChe()
    hurufSha = BuatHurufSha()
    hurufShcha = BuatHurufShcha()
    hurufYerry = BuatHurufYerry()
    hurufE = BuatHurufE()
    hurufYu = BuatHurufYu()
    hurufYa = BuatHurufYa()
    
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
    CombineRgn bentukAkhir, bentukAkhir, hurufEr, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEs, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufTe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufU, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEf, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufKha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufTse, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufChe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufSha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufShcha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYerry, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufE, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYu, 2
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
    DeleteObject bentukEr
    DeleteObject bentukEs
    DeleteObject bentukTe
    DeleteObject bentukU
    DeleteObject bentukEf
    DeleteObject bentukKha
    DeleteObject bentukTse
    DeleteObject bentukChe
    DeleteObject bentukSha
    DeleteObject bentukShcha
    DeleteObject bentukYerry
    DeleteObject bentukE
    DeleteObject bentukYu
    DeleteObject bentukYa
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub
