VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public m_Texture As cTexture        ' texture awal

Private m_Quadric As Long


Public Function DrawGLScene(bFireColor As Boolean) As Boolean
    Static xrot As GLfloat
    Static yrot As GLfloat
    Static zrot As GLfloat
    
    ' Clear the backbuffer and the depth buffer
    glClear clrColorBufferBit Or clrDepthBufferBit
    ' Reset the modelview matrix
    glLoadIdentity
    
    ' Translate out of the scene
    glTranslatef 0#, 0#, gflZ
    ' Rotate the scene along the x and y axis
    glRotatef xrot, 1#, 0#, 0#
    glRotatef yrot, 0#, 1#, 0#
    glRotatef zrot, 0#, 0#, 1#
    
    ' BAGIAN ATAS (KERUCUT / CONE)
    
    glPushMatrix
    
        ' Warna Kerucut
        glColor3f 1#, 0.2, 0.2
        
        ' Putar -90 derajat agar "berdiri"
        glRotatef -90, 1, 0, 0
        ' Pindahkan ke posisi Y=2
        glTranslatef 0, 0, 2
        
        ' Gambar Kerucut: (Object, RadiusBawah, RadiusAtas, Tinggi, Slices, Stacks)
        gluCylinder m_Quadric, 1.2, 0, 2, 16, 2
        
    glPopMatrix
    

    
    ' BADAN (TABUNG)
    
    glPushMatrix
    
        ' Warna Badan
        glColor3f 0.9, 0.9, 1#
        
        ' Putar -90 derajat agar "berdiri"
        glRotatef -90, 1, 0, 0
        ' Pindahkan ke posisi Y=-3
        glTranslatef 0, 0, -3
        
        ' Gambar Tabung: (Object, RadiusBawah, RadiusAtas, Tinggi, Slices, Stacks)
        gluCylinder m_Quadric, 1.2, 1.2, 5, 16, 5

 
        glColor3f 0.5, 0.5, 0.5 ' Warna abu-abu
        gluDisk m_Quadric, 0, 1.2, 16, 1 ' (Obj, InnerRadius, OuterRadius, Slices, Loops)

    glPopMatrix
    

    
    ' KAKI ROKET (3 KAKI)
    
    ' Kita tetap pakai GL_TRIANGLES untuk sirip, karena lebih cocok
    glBegin GL_TRIANGLES

        ' Warna Kaki (Abu-abu)
        glColor3f 0.5, 0.5, 0.5
        
        ' Kaki 1 (Depan)
        glVertex3f 0#, -2#, 1.2     ' Nempel di badan
        glVertex3f 0#, -4#, 1.2     ' Ujung bawah
        glVertex3f 0#, -2#, 2.5     ' Ujung luar
        
        ' Kaki 2 (Kiri Belakang)
        glVertex3f -1.03, -2#, -0.6 ' Nempel di badan
        glVertex3f -1.03, -4#, -0.6 ' Ujung bawah
        glVertex3f -2.16, -2#, -1.25 ' Ujung luar
        
        ' Kaki 3 (Kanan Belakang)
        glVertex3f 1.03, -2#, -0.6  ' Nempel di badan
        glVertex3f 1.03, -4#, -0.6  ' Ujung bawah
        glVertex3f 2.16, -2#, -1.25 ' Ujung luar
        
    glEnd


    ' API (KERUCUT TERBALIK)
    
    glPushMatrix
    
        ' warna kelap-kelip
        If bFireColor Then
            glColor3f 1#, 0.5, 0# ' Oranye
        Else
            glColor3f 1#, 1#, 0#   ' Kuning
        End If
        
        ' Putar 90 derajat agar "menghadap ke bawah"
        glRotatef 90, 1, 0, 0
        ' Pindahkan ke posisi Y=-3 (pangkal api)
        glTranslatef 0, 0, 3

        ' Gambar Kerucut: (Obj, RadiusBawah, RadiusAtas, Tinggi, Slices, Stacks)
        ' RadiusAtas = 0 berarti kerucut
        gluCylinder m_Quadric, 1#, 0, 2, 12, 1 ' Tinggi 2 (dari Y=-3 ke Y=-5)

    glPopMatrix


    xrot = xrot + gflXSpeed
    yrot = yrot + gflYSpeed
    zrot = zrot + gflZSpeed
    
    DrawGLScene = True
End Function

Public Function InitGL() As Boolean
    Dim aflLightAmbient(4) As GLfloat
    Dim aflLightDiffuse(4) As GLfloat
    Dim aflLightPosition(4) As GLfloat
    
    ' Create new texture
    'Set m_Texture = New cTexture
    'm_Texture.loadTexture App.Path & "\Data\Crate.tga", FILETYPE_TGA
    
    ' Enable texture mapping
    glDisable glcTexture2D
    ' Smooth shading
    glShadeModel smSmooth
    
    ' Set the clear colour
    glClearColor 0#, 0#, 0#, 0#
    ' Set the clear depth
    glClearDepth 1#
    
    
    ' Enable Z-buffer
    glEnable glcDepthTest
    ' Set test type
    glDepthFunc cfLEqual
    ' Best perspective correction
    glHint htPerspectiveCorrectionHint, hmNicest
      
    
    aflLightAmbient(0) = 0.5
    aflLightAmbient(1) = 0.5
    aflLightAmbient(2) = 0.5
    aflLightAmbient(3) = 1#
    aflLightDiffuse(0) = 1#
    aflLightDiffuse(1) = 1#
    aflLightDiffuse(2) = 1#
    aflLightDiffuse(3) = 1#
    aflLightPosition(0) = 0#
    aflLightPosition(1) = 0#
    aflLightPosition(2) = 2#
    aflLightPosition(3) = 1#
      
    glLightfv ltLight1, lpmAmbient, aflLightAmbient(0)
    glLightfv ltLight1, lpmDiffuse, aflLightDiffuse(0)
    glLightfv ltLight1, lpmPosition, aflLightPosition(0)
    
    glDisable glcLight1
    
    m_Quadric = gluNewQuadric()
    
    gluQuadricDrawStyle m_Quadric, GLU_FILL
    
    gluQuadricNormals m_Quadric, GLU_SMOOTH
    ' ======================================
    
    InitGL = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Set the key to be pressed
    gbKeys(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Set the key to be not pressed
    gbKeys(KeyCode) = False
End Sub

Private Sub Form_Load()
    Dim bFullscreen As Boolean
    Dim frm As frmMain
    Dim bLightSwitched As Boolean
    Dim bFilterSwitched As Boolean
    Dim bLightOn As Boolean
    Dim giCurrFilter As Integer
    
    Dim bFireColor As Boolean
    Dim iFireCounter As Integer

    ' Put us into fullscreen automatically
    bFullscreen = True
    bLightSwitched = False
    bFilterSwitched = False
    bLightOn = False
    gflZ = -10#

    ' Save the current display settings
    SaveDisplaySettings

    ' Show this form
    Me.Show
    ' Attempt to create a compatible GL window and set the display mode
    If (CreateGLWindow(Me, 640, 480, 32, bFullscreen) = False) Then
        Unload Me
    End If
    
    ' Attempt to set up OpenGL
    If (InitGL() = False) Then
        Unload Me
    End If
  
    ' Loop until the form is unloaded, process windows events every time we're not rendering
    Do While DoEvents()
    
        ' LOGIKA KELAP-KELIP
        iFireCounter = iFireCounter + 1
        If iFireCounter > 5 Then
            iFireCounter = 0
            bFireColor = Not bFireColor
        End If
    
        ' Render the scene, if it failed or the user has pressed the escape key then exit the program
        If (DrawGLScene(bFireColor) = False) Or (gbKeys(vbKeyEscape)) Then
       
             Unload Me
        Else
            ' Swap the front and back buffers to display what we've just rendered
            SwapBuffers Me.hDC
      
     
            '
            ' Toggle lighting
            'If (gbKeys(vbKeyL) = True) And (bLightSwitched = False) Then
            '
            '    bLightOn = Not (bLightOn)
            '
            '    If (bLightOn) Then
            '        glEnable glcLighting
            '    Else
            '        glDisable glcLighting
            '    End If
            '
            '    bLightSwitched = True
            'End If
            '
            'If (gbKeys(vbKeyL) = False) Then
            '    bLightSwitched = False
            'End If
      
     
            '
            ' Toggle filtering
            'If (gbKeys(vbKeyF) = True) And (bFilterSwitched = False) Then
            '    giCurrFilter = m_Texture.getFilter
            '    giCurrFilter = giCurrFilter + 1
            '    If giCurrFilter > 2 Then giCurrFilter = 0
            '    Select Case giCurrFilter
            '        Case 0:
            '            m_Texture.setFilter FILTER_NEAREST
            '        Case 1:
            '            m_Texture.setFilter FILTER_LINEAR
            '        Case 2:
            '            m_Texture.setFilter FILTER_MIPMAPPED
            '    End Select
            '
            '    bFilterSwitched = True
            'End If
            '
            'If (gbKeys(vbKeyF) = False) Then
            '    bFilterSwitched = False
            'End If
        
         
            ' Zoom in and out
            If (gbKeys(vbKeyPageUp) = True) Then
                gflZ = gflZ - 0.02
            End If
            
            If (gbKeys(vbKeyPageDown) = True) Then
                gflZ = gflZ + 0.02
            End If
            
            ' Increase / decrease cube's rotation amount
            If (gbKeys(vbKeyUp) = True) Then
                gflXSpeed = gflXSpeed - 0.01
            End If
    
         
            If (gbKeys(vbKeyDown) = True) Then
                gflXSpeed = gflXSpeed + 0.01
            End If
            
            If (gbKeys(vbKeyA) = True) Then
                gflYSpeed = gflYSpeed - 0.01
            End If
    
            If (gbKeys(vbKeyD) = True) Then
                gflYSpeed = gflYSpeed + 0.01
            End If
            
            If (gbKeys(vbKeyLeft) = True) Then
                gflZSpeed = gflZSpeed - 0.01
            End If
            
            If (gbKeys(vbKeyRight) = True) Then
                gflZSpeed = gflZSpeed + 0.01
            End If
            
            If (gbKeys(vbKeySpace) = True) Then
                gflXSpeed = 0
                gflYSpeed = 0
                gflZSpeed = 0
            End If
         
            ' Key escape has been pressed, so exit the program!
            If (gbKeys(vbKeyEscape) = True) Then
                Unload Me
            End If
        End If
    Loop
End Sub

Private Sub Form_Resize()
        ' Primarily used when in windowed mode
    ReSizeGLScene ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If m_Quadric Then gluDeleteQuadric m_Quadric
    
    ' Shut down OpenGL
    KillGLWindow Me
End Sub
