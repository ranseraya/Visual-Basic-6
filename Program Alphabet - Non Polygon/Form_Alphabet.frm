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
    
    bentukAkhir = CreateRectRgn(0, 0, 1, 1)
    
    hurufA = BuatHurufA()
    hurufBe = BuatHurufBe()
    hurufVe = BuatHurufVe()
    hurufGe = BuatHurufGe()
    hurufDe = BuatHurufDe()
    hurufYe = BuatHurufYe()
    hurufYo = BuatHurufYo()
    hurufZhe = BuatHurufZhe()
    hurufEn = BuatHurufEn()
    hurufO = BuatHurufO()
    hurufPe = BuatHurufPe()
    hurufEr = BuatHurufEr()
    hurufEs = BuatHurufEs()
    hurufTe = BuatHurufTe()
    hurufEf = BuatHurufEf()
    hurufTse = BuatHurufTse()
    hurufChe = BuatHurufChe()
    hurufSha = BuatHurufSha()
    hurufShcha = BuatHurufShcha()
    hurufYerry = BuatHurufYerry()
    hurufYu = BuatHurufYu()
    hurufYa = BuatHurufYa()
    
    Call OffsetRgn(hurufTse, -130, 0)
    Call OffsetRgn(hurufChe, 0, -150)
    Call OffsetRgn(hurufSha, 0, -150)
    Call OffsetRgn(hurufShcha, 720, -150)
    Call OffsetRgn(hurufYerry, 780, -300)
    Call OffsetRgn(hurufYu, 420, -300)
    
    CombineRgn bentukAkhir, bentukAkhir, hurufA, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufBe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufVe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufGe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufDe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYo, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufZhe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEl, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEn, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufO, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufPe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEr, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEs, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufTe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufEf, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufTse, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufChe, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufSha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufShcha, 2
    CombineRgn bentukAkhir, bentukAkhir, hurufYerry, 2
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
    DeleteObject bentukEl
    DeleteObject bentukEn
    DeleteObject bentukO
    DeleteObject bentukPe
    DeleteObject bentukEr
    DeleteObject bentukEs
    DeleteObject bentukTe
    DeleteObject bentukEf
    DeleteObject bentukTse
    DeleteObject bentukChe
    DeleteObject bentukSha
    DeleteObject bentukShcha
    DeleteObject bentukYerry
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
