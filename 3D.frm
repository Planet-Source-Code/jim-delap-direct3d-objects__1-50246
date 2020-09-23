VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim DX As DirectX8
Dim D3DX As D3DX8
Dim D3D As Direct3D8
Dim D3DD As Direct3DDevice8
Dim DM As D3DDISPLAYMODE
Dim DPP As D3DPRESENT_PARAMETERS
Dim DI As DirectInput8
Dim DID As DirectInputDevice8
Dim DIS As DIKEYBOARDSTATE

Const FVF = D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1
Const PI = 3.14159265358979
Const RAD = PI / 180

Const RED = &HFF0000
Const GREEN = &HFF00
Const BLUE = &HFF
Const WHITE = &HFFFFFF
    
Private Type TLV
    X As Single
    Y As Single
    Z As Single
    Color As Long
End Type

Private Type UdtObj
    TLV(0 To 199, 0 To 9) As TLV
End Type

Private Type UdtFrame
    Rate As Byte
    LC As Single
    Drawn As Byte
    MF As D3DXFont
    MFD As IFont
    vRECT As RECT
End Type

Dim Angle As Single
Dim Shape As Integer
Dim PrimAmount(199) As Integer

Dim TEMP As Single
Dim TEMPDIR As Integer
Dim Col, Col2 As Long
Dim ColTemp, ColD As Single

Dim matWorld As D3DMATRIX
Dim matTemp As D3DMATRIX
Dim matProj As D3DMATRIX
Dim matView As D3DMATRIX

Dim Obj As UdtObj
Dim Frame As UdtFrame

Private Sub Form_Load()
    Me.Show
    StartDX
    Init
    Start
End Sub

Sub StartDX()
    Set DX = New DirectX8
    Set D3DX = New D3DX8
    Set D3D = DX.Direct3DCreate()
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DM
    With DPP
        .BackBufferCount = 1
        .BackBufferFormat = DM.Format
        .BackBufferWidth = 640 'DM.Width
        .BackBufferHeight = 480 'DM.Height
        .hDeviceWindow = Me.hWnd
        .AutoDepthStencilFormat = D3DFMT_D16
        .EnableAutoDepthStencil = True
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .Windowed = 0
    End With
    Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, DPP)
    
    D3DD.SetVertexShader FVF
    D3DD.SetRenderState D3DRS_LIGHTING, 0
    D3DD.SetRenderState D3DRS_ZENABLE, 1
    D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    Set DI = DX.DirectInputCreate
    Set DID = DI.CreateDevice("GUID_SysKeyboard")
    DID.SetCommonDataFormat DIFORMAT_KEYBOARD
    DID.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DID.Acquire

    
    D3DXMatrixIdentity matWorld
    D3DD.SetTransform D3DTS_WORLD, matWorld
    
    D3DXMatrixLookAtLH matView, MV(0, 5, 2), MV(0, 0, 0), MV(0, 1, 0)
    D3DD.SetTransform D3DTS_VIEW, matView
    
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1, 0.1, 75
    D3DD.SetTransform D3DTS_PROJECTION, matProj
End Sub

Sub EndDX()
    DID.Unacquire
    Set D3DX = Nothing
    Set D3DD = Nothing
    Set D3D = Nothing
    Set DX = Nothing
End Sub

Sub Init()

    With Frame
        .LC = 0
        .Drawn = 0
        .Rate = 0
        Font.Size = DPP.BackBufferWidth / (DPP.BackBufferWidth / 10)
        Font.Bold = True
        Set .MFD = Font
        Set .MF = D3DX.CreateFont(D3DD, .MFD.hFont)
        With .vRECT
            .Left = 1
            .Top = 1
            .Right = DPP.BackBufferWidth / 2
            .bottom = DPP.BackBufferHeight / 2
        End With
    End With
    
    Shape = 0
    'Angle = 0
    TEMP = -1.5
    TEMPDIR = 1
    ColTemp = 0
    Col = RGB(255, ColTemp, ColTemp)
    Col2 = RGB(ColTemp, 255, 255)
    
    'CUBE
    PrimAmount(0) = 12
    DoTLV 0
        
    'TRIANGLE
    PrimAmount(1) = 8
    DoTLV 1
        
    'HEXAGON
    PrimAmount(2) = 20
    DoTLV 2
        
    'TRIANGLE 2
    PrimAmount(3) = 4
    DoTLV 3
        
    'CUBE 2
    PrimAmount(4) = 24
    DoTLV 4
    
    'TRIANGLE 3
    PrimAmount(5) = 34
    DoTLV 5
    
    'SQUARE
    PrimAmount(6) = 32
    DoTLV 6
End Sub

Sub Start()
    Do
        CheckKeys
        Render
        DoEvents
    Loop
End Sub

Sub CheckKeys()
    DID.GetDeviceStateKeyboard DIS
    If DIS.Key(DIK_ESCAPE) Then
        EndDX
        End
    End If
    
    'If DIS.Key(DIK_LEFT) Then
    '    If Not Angle <= 0 Then Angle = Angle - 1 Else Angle = 359
    'ElseIf DIS.Key(DIK_RIGHT) Then
    '    If Not Angle >= 360 Then Angle = Angle + 1 Else Angle = 1
    'End If
    
    If DIS.Key(DIK_1) Then
        Shape = 0
    ElseIf DIS.Key(DIK_2) Then
        Shape = 1
    ElseIf DIS.Key(DIK_3) Then
        Shape = 2
    ElseIf DIS.Key(DIK_4) Then
        Shape = 3
    ElseIf DIS.Key(DIK_5) Then
        Shape = 4
    ElseIf DIS.Key(DIK_6) Then
        Shape = 5
    ElseIf DIS.Key(DIK_7) Then
        Shape = 6
    End If
End Sub

Sub Render()
    D3DD.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    D3DD.BeginScene
    
    DoObj Shape
    
    With Obj
        D3DD.DrawPrimitiveUP D3DPT_TRIANGLELIST, PrimAmount(Shape), .TLV(0, Shape), Len(.TLV(0, Shape))
    End With
    
    With Frame
        D3DX.DrawText .MF, D3DColorRGBA(255, 255, 255, 255), "FPS: " & GetFrameRate & vbNewLine & "Coder: Jim De Lap", .vRECT, DT_LEFT Or DT_TOP
    End With
    
    D3DD.EndScene
    D3DD.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Function CTLV(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Color As Long) As TLV
    With CTLV
        .X = X
        .Y = Y
        .Z = Z
        .Color = Color
    End With
End Function

Private Function MV(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    With MV
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Sub DoObj(ByVal Index As Integer)

    D3DXMatrixIdentity matWorld
        
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationX matTemp, Angle * RAD
    D3DXMatrixMultiply matWorld, matWorld, matTemp
        
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationY matTemp, Angle * RAD
    D3DXMatrixMultiply matWorld, matWorld, matTemp
        
    D3DXMatrixIdentity matTemp
    D3DXMatrixRotationZ matTemp, Angle * RAD
    D3DXMatrixMultiply matWorld, matWorld, matTemp
        
    DoTLV Index
    If Not Angle >= 360 Then Angle = Angle + 0.5 Else Angle = 0.5
    If TEMPDIR = 1 Then If Not TEMP >= 1.5 Then TEMP = TEMP + 0.01 Else TEMPDIR = 2
    If TEMPDIR = 2 Then If Not TEMP <= -1.5 Then TEMP = TEMP - 0.01 Else TEMPDIR = 1
    
    D3DD.SetTransform D3DTS_WORLD, matWorld
End Sub

Sub DoTLV(ByVal Index As Integer)
    With Obj
        Select Case Index
            Case 0
                'CUBE
                'FRONT
                .TLV(0, 0) = CTLV(-1, -1, -1, RED)
                .TLV(1, 0) = CTLV(1, -1, -1, BLUE)
                .TLV(2, 0) = CTLV(1, 1, -1, WHITE)
                
                .TLV(3, 0) = CTLV(1, 1, -1, WHITE)
                .TLV(4, 0) = CTLV(-1, 1, -1, GREEN)
                .TLV(5, 0) = CTLV(-1, -1, -1, RED)
                
                'BACK
                .TLV(6, 0) = CTLV(-1, -1, 1, RED)
                .TLV(7, 0) = CTLV(1, -1, 1, BLUE)
                .TLV(8, 0) = CTLV(1, 1, 1, WHITE)
                
                .TLV(9, 0) = CTLV(1, 1, 1, WHITE)
                .TLV(10, 0) = CTLV(-1, 1, 1, GREEN)
                .TLV(11, 0) = CTLV(-1, -1, 1, RED)
                
                'TOP
                .TLV(12, 0) = CTLV(-1, 1, -1, GREEN)
                .TLV(13, 0) = CTLV(1, 1, -1, WHITE)
                .TLV(14, 0) = CTLV(1, 1, 1, WHITE)
                
                .TLV(15, 0) = CTLV(1, 1, 1, WHITE)
                .TLV(16, 0) = CTLV(-1, 1, 1, GREEN)
                .TLV(17, 0) = CTLV(-1, 1, -1, GREEN)
                
                'BOTTOM
                .TLV(18, 0) = CTLV(-1, -1, -1, RED)
                .TLV(19, 0) = CTLV(1, -1, -1, BLUE)
                .TLV(20, 0) = CTLV(1, -1, 1, BLUE)
                
                .TLV(21, 0) = CTLV(1, -1, 1, BLUE)
                .TLV(22, 0) = CTLV(-1, -1, 1, RED)
                .TLV(23, 0) = CTLV(-1, -1, -1, RED)
                
                'LEFT
                .TLV(24, 0) = CTLV(-1, -1, 1, RED)
                .TLV(25, 0) = CTLV(-1, -1, -1, RED)
                .TLV(26, 0) = CTLV(-1, 1, -1, GREEN)
                
                .TLV(27, 0) = CTLV(-1, 1, -1, GREEN)
                .TLV(28, 0) = CTLV(-1, 1, 1, GREEN)
                .TLV(29, 0) = CTLV(-1, -1, 1, RED)
                
                'RIGHT
                .TLV(30, 0) = CTLV(1, -1, -1, BLUE)
                .TLV(31, 0) = CTLV(1, -1, 1, BLUE)
                .TLV(32, 0) = CTLV(1, 1, 1, WHITE)
                
                .TLV(33, 0) = CTLV(1, 1, 1, WHITE)
                .TLV(34, 0) = CTLV(1, 1, -1, WHITE)
                .TLV(35, 0) = CTLV(1, -1, -1, BLUE)
            Case 1
                'TRIANGLE
                'FRONT
                .TLV(0, 1) = CTLV(-1, -1, -1, RED)
                .TLV(1, 1) = CTLV(1, -1, -1, WHITE)
                .TLV(2, 1) = CTLV(0, 1, -1, BLUE)
    
                'BACK
                .TLV(3, 1) = CTLV(-1, -1, 1, BLUE)
                .TLV(4, 1) = CTLV(1, -1, 1, WHITE)
                .TLV(5, 1) = CTLV(0, 1, 1, RED)
            
                'LEFT
                .TLV(6, 1) = CTLV(-1, -1, -1, RED)
                .TLV(7, 1) = CTLV(-1, -1, 1, BLUE)
                .TLV(8, 1) = CTLV(0, 1, 1, WHITE)
        
                .TLV(9, 1) = CTLV(0, 1, 1, WHITE)
                .TLV(10, 1) = CTLV(0, 1, -1, WHITE)
                .TLV(11, 1) = CTLV(-1, -1, -1, RED)
        
                'RIGHT
                .TLV(12, 1) = CTLV(1, -1, -1, BLUE)
                .TLV(13, 1) = CTLV(1, -1, 1, RED)
                .TLV(14, 1) = CTLV(0, 1, 1, WHITE)
        
                .TLV(15, 1) = CTLV(0, 1, 1, WHITE)
                .TLV(16, 1) = CTLV(0, 1, -1, WHITE)
                .TLV(17, 1) = CTLV(1, -1, -1, BLUE)
        
                'BOTTOM
                .TLV(18, 1) = CTLV(-1, -1, -1, RED)
                .TLV(19, 1) = CTLV(1, -1, -1, BLUE)
                .TLV(20, 1) = CTLV(1, -1, 1, RED)
        
                .TLV(21, 1) = CTLV(1, -1, 1, RED)
                .TLV(22, 1) = CTLV(-1, -1, 1, BLUE)
                .TLV(23, 1) = CTLV(-1, -1, -1, RED)
            Case 2
                'HEXAGON
                'FRONT
                .TLV(0, 2) = CTLV(-1, 0, -1, GREEN)
                .TLV(1, 2) = CTLV(-0.5, -1, -1, BLUE)
                .TLV(2, 2) = CTLV(-0.5, 1, -1, RED)
        
                .TLV(3, 2) = CTLV(-0.5, -1, -1, BLUE)
                .TLV(4, 2) = CTLV(0.5, -1, -1, RED)
                .TLV(5, 2) = CTLV(0.5, 1, -1, BLUE)
        
                .TLV(6, 2) = CTLV(0.5, 1, -1, BLUE)
                .TLV(7, 2) = CTLV(-0.5, 1, -1, RED)
                .TLV(8, 2) = CTLV(-0.5, -1, -1, BLUE)
        
                .TLV(9, 2) = CTLV(1, 0, -1, GREEN)
                .TLV(10, 2) = CTLV(0.5, -1, -1, RED)
                .TLV(11, 2) = CTLV(0.5, 1, -1, BLUE)
        
                'BACK
                .TLV(12, 2) = CTLV(-1, 0, 1, GREEN)
                .TLV(13, 2) = CTLV(-0.5, -1, 1, BLUE)
                .TLV(14, 2) = CTLV(-0.5, 1, 1, RED)
        
                .TLV(15, 2) = CTLV(-0.5, -1, 1, BLUE)
                .TLV(16, 2) = CTLV(0.5, -1, 1, RED)
                .TLV(17, 2) = CTLV(0.5, 1, 1, BLUE)
        
                .TLV(18, 2) = CTLV(0.5, 1, 1, BLUE)
                .TLV(19, 2) = CTLV(-0.5, 1, 1, RED)
                .TLV(20, 2) = CTLV(-0.5, -1, 1, BLUE)
        
                .TLV(21, 2) = CTLV(1, 0, 1, GREEN)
                .TLV(22, 2) = CTLV(0.5, -1, 1, RED)
                .TLV(23, 2) = CTLV(0.5, 1, 1, BLUE)
        
                'BOTTOM LEFT
                .TLV(24, 2) = CTLV(-1, 0, -1, GREEN)
                .TLV(25, 2) = CTLV(-0.5, -1, -1, BLUE)
                .TLV(26, 2) = CTLV(-0.5, -1, 1, BLUE)
        
                .TLV(27, 2) = CTLV(-0.5, -1, 1, BLUE)
                .TLV(28, 2) = CTLV(-1, 0, 1, GREEN)
                .TLV(29, 2) = CTLV(-1, 0, -1, GREEN)
        
                'BOTTOM MIDDLE
                .TLV(30, 2) = CTLV(-0.5, -1, -1, BLUE)
                .TLV(31, 2) = CTLV(0.5, -1, -1, RED)
                .TLV(32, 2) = CTLV(0.5, -1, 1, RED)
        
                .TLV(33, 2) = CTLV(0.5, -1, 1, RED)
                .TLV(34, 2) = CTLV(-0.5, -1, 1, BLUE)
                .TLV(35, 2) = CTLV(-0.5, -1, -1, BLUE)
        
                'BOTTOM RIGHT
                .TLV(36, 2) = CTLV(0.5, -1, 1, RED)
                .TLV(37, 2) = CTLV(0.5, -1, -1, RED)
                .TLV(38, 2) = CTLV(1, 0, -1, GREEN)
        
                .TLV(39, 2) = CTLV(1, 0, -1, GREEN)
                .TLV(40, 2) = CTLV(1, 0, 1, GREEN)
                .TLV(41, 2) = CTLV(0.5, -1, 1, RED)
        
                'TOP LEFT
                .TLV(42, 2) = CTLV(-1, 0, -1, GREEN)
                .TLV(43, 2) = CTLV(-0.5, 1, -1, RED)
                .TLV(44, 2) = CTLV(-0.5, 1, 1, RED)
        
                .TLV(45, 2) = CTLV(-0.5, 1, 1, RED)
                .TLV(46, 2) = CTLV(-1, 0, 1, GREEN)
                .TLV(47, 2) = CTLV(-1, 0, -1, GREEN)
        
                'TOP MIDDLE
                .TLV(48, 2) = CTLV(-0.5, 1, -1, RED)
                .TLV(49, 2) = CTLV(0.5, 1, -1, BLUE)
                .TLV(50, 2) = CTLV(0.5, 1, 1, BLUE)
        
                .TLV(51, 2) = CTLV(0.5, 1, 1, BLUE)
                .TLV(52, 2) = CTLV(-0.5, 1, 1, RED)
                .TLV(53, 2) = CTLV(-0.5, 1, -1, RED)
        
                'TOP RIGHT
                .TLV(54, 2) = CTLV(0.5, 1, -1, BLUE)
                .TLV(55, 2) = CTLV(1, 0, -1, GREEN)
                .TLV(56, 2) = CTLV(1, 0, 1, GREEN)
        
                .TLV(57, 2) = CTLV(1, 0, 1, GREEN)
                .TLV(58, 2) = CTLV(0.5, 1, 1, BLUE)
                .TLV(59, 2) = CTLV(0.5, 1, -1, BLUE)
            Case 3
                'TRIANGLE 2
                'FRONT
                .TLV(0, 3) = CTLV(-1, -1, -1, BLUE)
                .TLV(1, 3) = CTLV(1, -1, -1, GREEN)
                .TLV(2, 3) = CTLV(0, 1, -0.5, WHITE)
        
                'LEFT
                .TLV(3, 3) = CTLV(-1, -1, -1, BLUE)
                .TLV(4, 3) = CTLV(0, 0, 1, RED)
                .TLV(5, 3) = CTLV(0, 1, -0.5, WHITE)
        
                'RIGHT
                .TLV(6, 3) = CTLV(1, -1, -1, GREEN)
                .TLV(7, 3) = CTLV(0, 0, 1, RED)
                .TLV(8, 3) = CTLV(0, 1, -0.5, WHITE)
        
                'BOTTOM
                .TLV(9, 3) = CTLV(-1, -1, -1, BLUE)
                .TLV(10, 3) = CTLV(0, 0, 1, RED)
                .TLV(11, 3) = CTLV(1, -1, -1, GREEN)
            Case 4
                'CUBE 2
                'FRONT
                .TLV(0, 4) = CTLV(-1, -1, -1, RED)
                .TLV(1, 4) = CTLV(0, 0, -TEMP, WHITE)
                .TLV(2, 4) = CTLV(-1, 1, -1, RED)
        
                .TLV(3, 4) = CTLV(-1, 1, -1, RED)
                .TLV(4, 4) = CTLV(0, 0, -TEMP, WHITE)
                .TLV(5, 4) = CTLV(1, 1, -1, RED)
    
                .TLV(6, 4) = CTLV(1, 1, -1, RED)
                .TLV(7, 4) = CTLV(0, 0, -TEMP, WHITE)
                .TLV(8, 4) = CTLV(1, -1, -1, RED)
        
                .TLV(9, 4) = CTLV(1, -1, -1, RED)
                .TLV(10, 4) = CTLV(0, 0, -TEMP, WHITE)
                .TLV(11, 4) = CTLV(-1, -1, -1, RED)
        
                'BACK
                .TLV(12, 4) = CTLV(-1, -1, 1, RED)
                .TLV(13, 4) = CTLV(0, 0, TEMP, WHITE)
                .TLV(14, 4) = CTLV(-1, 1, 1, RED)
        
                .TLV(15, 4) = CTLV(-1, 1, 1, RED)
                .TLV(16, 4) = CTLV(0, 0, TEMP, WHITE)
                .TLV(17, 4) = CTLV(1, 1, 1, RED)
        
                .TLV(18, 4) = CTLV(1, 1, 1, RED)
                .TLV(19, 4) = CTLV(0, 0, TEMP, WHITE)
                .TLV(20, 4) = CTLV(1, -1, 1, RED)
        
                .TLV(21, 4) = CTLV(1, -1, 1, RED)
                .TLV(22, 4) = CTLV(0, 0, TEMP, WHITE)
                .TLV(23, 4) = CTLV(-1, -1, 1, RED)
        
                'TOP
                .TLV(24, 4) = CTLV(-1, 1, -1, RED)
                .TLV(25, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(26, 4) = CTLV(-1, 1, 1, RED)
        
                .TLV(27, 4) = CTLV(-1, 1, 1, RED)
                .TLV(28, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(29, 4) = CTLV(1, 1, 1, RED)
        
                .TLV(30, 4) = CTLV(1, 1, 1, RED)
                .TLV(31, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(32, 4) = CTLV(1, 1, -1, RED)
        
                .TLV(33, 4) = CTLV(1, 1, -1, RED)
                .TLV(34, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(35, 4) = CTLV(-1, 1, -1, RED)
        
                'BOTTOM
                .TLV(36, 4) = CTLV(-1, -1, -1, RED)
                .TLV(37, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(38, 4) = CTLV(-1, -1, 1, RED)
        
                .TLV(39, 4) = CTLV(-1, -1, 1, RED)
                .TLV(40, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(41, 4) = CTLV(1, -1, 1, RED)
        
                .TLV(42, 4) = CTLV(1, -1, 1, RED)
                .TLV(43, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(44, 4) = CTLV(1, -1, -1, RED)
        
                .TLV(45, 4) = CTLV(1, -1, -1, RED)
                .TLV(46, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(47, 4) = CTLV(-1, -1, -1, RED)
        
                'LEFT
                .TLV(48, 4) = CTLV(-1, -1, 1, RED)
                .TLV(49, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(50, 4) = CTLV(-1, 1, 1, RED)
        
                .TLV(51, 4) = CTLV(-1, 1, 1, RED)
                .TLV(52, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(53, 4) = CTLV(-1, 1, -1, RED)
        
                .TLV(54, 4) = CTLV(-1, 1, -1, RED)
                .TLV(55, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(56, 4) = CTLV(-1, -1, -1, RED)
        
                .TLV(57, 4) = CTLV(-1, -1, -1, RED)
                .TLV(58, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(59, 4) = CTLV(-1, -1, 1, RED)
        
                'RIGHT
                .TLV(60, 4) = CTLV(1, -1, -1, RED)
                .TLV(61, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(62, 4) = CTLV(1, 1, -1, RED)
        
                .TLV(63, 4) = CTLV(1, 1, -1, RED)
                .TLV(64, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(65, 4) = CTLV(1, 1, 1, RED)
        
                .TLV(66, 4) = CTLV(1, 1, 1, RED)
                .TLV(67, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(68, 4) = CTLV(1, -1, 1, RED)
        
                .TLV(69, 4) = CTLV(1, -1, 1, RED)
                .TLV(70, 4) = CTLV(0, 0, 0, WHITE)
                .TLV(71, 4) = CTLV(1, -1, -1, RED)
            Case 5
                'FRONT
                .TLV(0, 5) = CTLV(-1, 1, -1, RED)
                .TLV(1, 5) = CTLV(-0.5, 0, -1, WHITE)
                .TLV(2, 5) = CTLV(0, 0.5, -1, WHITE)
                
                .TLV(3, 5) = CTLV(1, 1, -1, RED)
                .TLV(4, 5) = CTLV(0.5, 0, -1, WHITE)
                .TLV(5, 5) = CTLV(0, 0.5, -1, WHITE)
                
                .TLV(6, 5) = CTLV(1, -1, -1, RED)
                .TLV(7, 5) = CTLV(0, -0.5, -1, WHITE)
                .TLV(8, 5) = CTLV(0.5, 0, -1, WHITE)
                
                .TLV(9, 5) = CTLV(-1, -1, -1, RED)
                .TLV(10, 5) = CTLV(0, -0.5, -1, WHITE)
                .TLV(11, 5) = CTLV(-0.5, 0, -1, WHITE)
                
                'MIDDLE OF FRONT
                .TLV(12, 5) = CTLV(0, 0.5, -1, WHITE)
                .TLV(13, 5) = CTLV(0.5, 0, -1, WHITE)
                .TLV(14, 5) = CTLV(0, -0.5, -1, WHITE)
                
                .TLV(15, 5) = CTLV(0, 0.5, -1, WHITE)
                .TLV(16, 5) = CTLV(-0.5, 0, -1, WHITE)
                .TLV(17, 5) = CTLV(0, -0.5, -1, WHITE)
                
                'BACK
                .TLV(18, 5) = CTLV(0, 1, 1, WHITE)
                .TLV(19, 5) = CTLV(-0.3, 0.3, 1, BLUE)
                .TLV(20, 5) = CTLV(0.3, 0.3, 1, BLUE)
                
                .TLV(21, 5) = CTLV(1, 0, 1, WHITE)
                .TLV(22, 5) = CTLV(0.3, 0.3, 1, BLUE)
                .TLV(23, 5) = CTLV(0.3, -0.3, 1, BLUE)
                
                .TLV(24, 5) = CTLV(0, -1, 1, WHITE)
                .TLV(25, 5) = CTLV(0.3, -0.3, 1, BLUE)
                .TLV(26, 5) = CTLV(-0.3, -0.3, 1, BLUE)
                
                .TLV(27, 5) = CTLV(-1, 0, 1, WHITE)
                .TLV(28, 5) = CTLV(-0.3, 0.3, 1, BLUE)
                .TLV(29, 5) = CTLV(-0.3, -0.3, 1, BLUE)
                
                'MIDDLE OF BACK
                .TLV(30, 5) = CTLV(-0.3, 0.3, 1, BLUE)
                .TLV(31, 5) = CTLV(0.3, 0.3, 1, BLUE)
                .TLV(32, 5) = CTLV(0.3, -0.3, 1, BLUE)
                
                .TLV(33, 5) = CTLV(0.3, -0.3, 1, BLUE)
                .TLV(34, 5) = CTLV(-0.3, -0.3, 1, BLUE)
                .TLV(35, 5) = CTLV(-0.3, 0.3, 1, BLUE)
                
                'TOP 4
                .TLV(36, 5) = CTLV(0, 0.5, -1, WHITE)
                .TLV(37, 5) = CTLV(0, 1, 1, WHITE)
                .TLV(38, 5) = CTLV(-1, 1, -1, RED)
                
                .TLV(39, 5) = CTLV(0, 1, 1, WHITE)
                .TLV(40, 5) = CTLV(-0.3, 0.3, 1, BLUE)
                .TLV(41, 5) = CTLV(-1, 1, -1, RED)
                
                .TLV(42, 5) = CTLV(0, 0.5, -1, WHITE)
                .TLV(43, 5) = CTLV(0, 1, 1, WHITE)
                .TLV(44, 5) = CTLV(1, 1, -1, RED)
                
                .TLV(45, 5) = CTLV(0.3, 0.3, 1, BLUE)
                .TLV(46, 5) = CTLV(1, 1, -1, RED)
                .TLV(47, 5) = CTLV(0, 1, 1, WHITE)
                
                'RIGHT 4
                .TLV(48, 5) = CTLV(0.5, 0, -1, WHITE)
                .TLV(49, 5) = CTLV(1, 0, 1, WHITE)
                .TLV(50, 5) = CTLV(1, 1, -1, RED)
                
                .TLV(51, 5) = CTLV(1, 0, 1, WHITE)
                .TLV(52, 5) = CTLV(0.3, 0.3, 1, BLUE)
                .TLV(53, 5) = CTLV(1, 1, -1, RED)
                
                .TLV(54, 5) = CTLV(0.5, 0, -1, WHITE)
                .TLV(55, 5) = CTLV(1, 0, 1, WHITE)
                .TLV(56, 5) = CTLV(1, -1, -1, RED)
                
                .TLV(57, 5) = CTLV(1, -1, -1, RED)
                .TLV(58, 5) = CTLV(1, 0, 1, WHITE)
                .TLV(59, 5) = CTLV(0.3, -0.3, 1, BLUE)
                
                'BOTTOM 4
                .TLV(60, 5) = CTLV(0.3, -0.3, 1, BLUE)
                .TLV(61, 5) = CTLV(0, -1, 1, WHITE)
                .TLV(62, 5) = CTLV(1, -1, -1, RED)
                
                .TLV(63, 5) = CTLV(0, -0.5, -1, WHITE)
                .TLV(64, 5) = CTLV(0, -1, 1, WHITE)
                .TLV(65, 5) = CTLV(1, -1, -1, RED)
                
                .TLV(66, 5) = CTLV(0, -0.5, -1, WHITE)
                .TLV(67, 5) = CTLV(0, -1, 1, WHITE)
                .TLV(68, 5) = CTLV(-1, -1, -1, RED)
                
                .TLV(69, 5) = CTLV(-0.3, -0.3, 1, BLUE)
                .TLV(70, 5) = CTLV(0, -1, 1, WHITE)
                .TLV(71, 5) = CTLV(-1, -1, -1, RED)
                
                'LEFT 4
                .TLV(72, 5) = CTLV(-0.3, -0.3, 1, BLUE)
                .TLV(73, 5) = CTLV(-1, 0, 1, WHITE)
                .TLV(74, 5) = CTLV(-1, -1, -1, RED)
                
                .TLV(75, 5) = CTLV(-0.5, 0, -1, WHITE)
                .TLV(76, 5) = CTLV(-1, 0, 1, WHITE)
                .TLV(77, 5) = CTLV(-1, -1, -1, RED)
                
                .TLV(78, 5) = CTLV(-0.5, 0, -1, WHITE)
                .TLV(79, 5) = CTLV(-1, 0, 1, WHITE)
                .TLV(80, 5) = CTLV(-1, 1, -1, RED)
                
                .TLV(81, 5) = CTLV(-0.3, 0.3, 1, BLUE)
                .TLV(82, 5) = CTLV(-1, 0, 1, WHITE)
                .TLV(83, 5) = CTLV(-1, 1, -1, RED)
                
            Case 6
                'SQUARE
                'FRONT
                'LEFT
                .TLV(0, 6) = CTLV(-1, -1, -1, RED)
                .TLV(1, 6) = CTLV(-1, 1, -1, RED)
                .TLV(2, 6) = CTLV(-0.8, 1, -1, WHITE)
                
                .TLV(3, 6) = CTLV(-1, -1, -1, RED)
                .TLV(4, 6) = CTLV(-0.8, -1, -1, WHITE)
                .TLV(5, 6) = CTLV(-0.8, 1, -1, WHITE)
                
                'RIGHT
                .TLV(6, 6) = CTLV(1, -1, -1, RED)
                .TLV(7, 6) = CTLV(1, 1, -1, RED)
                .TLV(8, 6) = CTLV(0.8, 1, -1, WHITE)
                
                .TLV(9, 6) = CTLV(0.8, -1, -1, WHITE)
                .TLV(10, 6) = CTLV(1, -1, -1, RED)
                .TLV(11, 6) = CTLV(0.8, 1, -1, WHITE)
                
                'TOP
                .TLV(12, 6) = CTLV(-0.8, 1, -1, WHITE)
                .TLV(13, 6) = CTLV(-0.8, 0.8, -1, WHITE)
                .TLV(14, 6) = CTLV(0.8, 0.8, -1, WHITE)
                
                .TLV(15, 6) = CTLV(0.8, 0.8, -1, WHITE)
                .TLV(16, 6) = CTLV(0.8, 1, -1, WHITE)
                .TLV(17, 6) = CTLV(-0.8, 1, -1, WHITE)
                
                'BOTTOM
                .TLV(18, 6) = CTLV(-0.8, -1, -1, WHITE)
                .TLV(19, 6) = CTLV(0.8, -1, -1, WHITE)
                .TLV(20, 6) = CTLV(0.8, -0.8, -1, WHITE)
                
                .TLV(21, 6) = CTLV(0.8, -0.8, -1, WHITE)
                .TLV(22, 6) = CTLV(-0.8, -0.8, -1, WHITE)
                .TLV(23, 6) = CTLV(-0.8, -1, -1, WHITE)
                
                'BACK
                'LEFT
                .TLV(24, 6) = CTLV(-1, -1, 1, RED)
                .TLV(25, 6) = CTLV(-1, 1, 1, RED)
                .TLV(26, 6) = CTLV(-0.8, 1, 1, WHITE)
                
                .TLV(27, 6) = CTLV(-1, -1, 1, RED)
                .TLV(28, 6) = CTLV(-0.8, -1, 1, WHITE)
                .TLV(29, 6) = CTLV(-0.8, 1, 1, WHITE)
                
                'RIGHT
                .TLV(30, 6) = CTLV(1, -1, 1, RED)
                .TLV(31, 6) = CTLV(1, 1, 1, RED)
                .TLV(32, 6) = CTLV(0.8, 1, 1, WHITE)
                
                .TLV(33, 6) = CTLV(0.8, -1, 1, WHITE)
                .TLV(34, 6) = CTLV(1, -1, 1, RED)
                .TLV(35, 6) = CTLV(0.8, 1, 1, WHITE)
                
                'TOP
                .TLV(36, 6) = CTLV(-0.8, 1, 1, WHITE)
                .TLV(37, 6) = CTLV(-0.8, 0.8, 1, WHITE)
                .TLV(38, 6) = CTLV(0.8, 0.8, 1, WHITE)
                
                .TLV(39, 6) = CTLV(0.8, 0.8, 1, WHITE)
                .TLV(40, 6) = CTLV(0.8, 1, 1, WHITE)
                .TLV(41, 6) = CTLV(-0.8, 1, 1, WHITE)
                
                'BOTTOM
                .TLV(42, 6) = CTLV(-0.8, -1, 1, WHITE)
                .TLV(43, 6) = CTLV(0.8, -1, 1, WHITE)
                .TLV(44, 6) = CTLV(0.8, -0.8, 1, WHITE)
                
                .TLV(45, 6) = CTLV(0.8, -0.8, 1, WHITE)
                .TLV(46, 6) = CTLV(-0.8, -0.8, 1, WHITE)
                .TLV(47, 6) = CTLV(-0.8, -1, 1, WHITE)
                
                'OUTSIDE TOP
                .TLV(48, 6) = CTLV(-1, 1, -1, WHITE)
                .TLV(49, 6) = CTLV(-1, 1, 1, BLUE)
                .TLV(50, 6) = CTLV(1, 1, 1, WHITE)
                
                .TLV(51, 6) = CTLV(1, 1, 1, WHITE)
                .TLV(52, 6) = CTLV(1, 1, -1, BLUE)
                .TLV(53, 6) = CTLV(-1, 1, -1, WHITE)
                
                'OUTSIDE BOTTOM
                .TLV(54, 6) = CTLV(-1, -1, -1, WHITE)
                .TLV(55, 6) = CTLV(-1, -1, 1, BLUE)
                .TLV(56, 6) = CTLV(1, -1, 1, WHITE)
                
                .TLV(57, 6) = CTLV(1, -1, 1, WHITE)
                .TLV(58, 6) = CTLV(1, -1, -1, BLUE)
                .TLV(59, 6) = CTLV(-1, -1, -1, WHITE)
                
                'OUTSIDE LEFT
                .TLV(60, 6) = CTLV(-1, -1, -1, WHITE)
                .TLV(61, 6) = CTLV(-1, -1, 1, BLUE)
                .TLV(62, 6) = CTLV(-1, 1, 1, WHITE)
                
                .TLV(63, 6) = CTLV(-1, 1, 1, WHITE)
                .TLV(64, 6) = CTLV(-1, 1, -1, BLUE)
                .TLV(65, 6) = CTLV(-1, -1, -1, WHITE)
                
                'OUTSIDE RIGHT
                .TLV(66, 6) = CTLV(1, -1, -1, WHITE)
                .TLV(67, 6) = CTLV(1, -1, 1, BLUE)
                .TLV(68, 6) = CTLV(1, 1, 1, WHITE)
                
                .TLV(69, 6) = CTLV(1, 1, 1, WHITE)
                .TLV(70, 6) = CTLV(1, 1, -1, WHITE)
                .TLV(71, 6) = CTLV(1, -1, -1, WHITE)
                
                
                'INSIDE LEFT
                'BOTTOM
                .TLV(72, 6) = CTLV(-0.8, -0.8, -1, WHITE)
                .TLV(73, 6) = CTLV(-0.8, -0.8, 1, WHITE)
                .TLV(74, 6) = CTLV(0, 0, 0, RED)
                
                'TOP
                .TLV(75, 6) = CTLV(-0.8, 0.8, -1, WHITE)
                .TLV(76, 6) = CTLV(-0.8, 0.8, 1, WHITE)
                .TLV(77, 6) = CTLV(0, 0, 0, RED)
                
                'BACK
                .TLV(78, 6) = CTLV(-0.8, -0.8, 1, WHITE)
                .TLV(79, 6) = CTLV(-0.8, 0.8, 1, WHITE)
                .TLV(80, 6) = CTLV(0, 0, 0, RED)
                
                'FRONT
                .TLV(81, 6) = CTLV(-0.8, -0.8, -1, WHITE)
                .TLV(82, 6) = CTLV(-0.8, 0.8, -1, WHITE)
                .TLV(83, 6) = CTLV(0, 0, 0, RED)
                
                'INSIDE RIGHT
                'BOTTOM
                .TLV(84, 6) = CTLV(0.8, -0.8, -1, WHITE)
                .TLV(85, 6) = CTLV(0.8, -0.8, 1, WHITE)
                .TLV(86, 6) = CTLV(0, 0, 0, RED)
                
                'TOP
                .TLV(87, 6) = CTLV(0.8, 0.8, -1, WHITE)
                .TLV(88, 6) = CTLV(0.8, 0.8, 1, WHITE)
                .TLV(89, 6) = CTLV(0, 0, 0, RED)
                
                'BACK
                .TLV(90, 6) = CTLV(0.8, -0.8, 1, WHITE)
                .TLV(91, 6) = CTLV(0.8, 0.8, 1, WHITE)
                .TLV(92, 6) = CTLV(0, 0, 0, RED)
                
                'FRONT
                .TLV(93, 6) = CTLV(0.8, -0.8, -1, WHITE)
                .TLV(94, 6) = CTLV(0.8, 0.8, -1, WHITE)
                .TLV(95, 6) = CTLV(0, 0, 0, RED)
        End Select
    End With
End Sub

Private Function GetFrameRate() As Long
    With Frame
        If GetTickCount - .LC >= 1000 Then
            .Rate = .Drawn
            .LC = GetTickCount
            .Drawn = 0
        End If
        .Drawn = .Drawn + 1
        GetFrameRate = .Rate
    End With
End Function
