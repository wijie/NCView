VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   5415
   ForeColor       =   &H80000007&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":08CA
   ScaleHeight     =   4215
   ScaleWidth      =   5415
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   2040
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   480
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H8000000C&
      Height          =   732
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      Begin VB.PictureBox picDraw 
         Height          =   372
         Index           =   1
         Left            =   720
         MouseIcon       =   "frmMain.frx":0BD4
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   4
         Top             =   120
         Width           =   372
      End
      Begin VB.PictureBox picDraw 
         Height          =   372
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmMain.frx":0EDE
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   2
         Top             =   120
         Width           =   372
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1440
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12FA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '�㑵��
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�J��"
            Object.ToolTipText     =   "�J��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "���"
            Object.ToolTipText     =   "���"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�{��"
            Object.ToolTipText     =   "�{��"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDDE 
      BorderStyle     =   1  '����
      Caption         =   "DDE�p���x��"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "�J��(&O)"
      End
      Begin VB.Menu mnuNTIn 
         Caption         =   "NT�̓Ǎ���(&I)"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "���(&P)"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "NCView�̏I��(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�\��(&V)"
      Begin VB.Menu mnuLook 
         Caption         =   "�{��(&G)"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeisun 
         Caption         =   "�����\��(&L)"
      End
      Begin VB.Menu mnuStandard 
         Caption         =   "�S�̕\��(&S)"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "°�(&T)"
      Begin VB.Menu mnuInfo 
         Caption         =   "�H����(&I)"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "��߼��(&O)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "�ް�ޮݏ��(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tagRECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

' �E�B���h�E�̋�`�T�C�Y���擾
Private Declare Function GetWindowRect Lib "user32.dll" _
    (ByVal hwnd As Long, _
     lpRect As tagRECT) As Long

' �}�E�X�J�[�\���̈ړ��͈͂��w�肷��֐�
Private Declare Function ClipCursor Lib "user32.dll" _
    (lpRect As Any) As Long

' �V�X�e���̐ݒ��V�X�e�����g���b�N�̒l���擾����֐�
Private Declare Function GetSystemMetrics Lib "user32.dll" _
    (ByVal nIntex As Long) As Long

Private Const SM_CYCAPTION = 4 ' �^�C�g���o�[�̍������擾
Private Const SM_CYMENU = 15 ' �N���C�A���g�E�B���h�E�̃��j���[�̍������擾

Private mblnDisp As Boolean ' ��ʕ\���ォ�ۂ�������Flag
Private mblnPanMode As Boolean ' picDraw�̈ړ���/�s������Flag
Private msngDragDistX As Single ' MouseDown���̃}�E�X�|�C���^�[��X���W
Private msngDragDistY As Single ' MouseDown���̃}�E�X�|�C���^�[��X���W
Private msngCurrentTop As Single ' MouseMove����picDraw��Top�v���p�e�B
Private msngCurrentLeft As Single ' MouseMove����picDraw��Left�v���p�e�B
Private mblnHMove As Boolean ' picDraw���������ֈړ���/�s������Flag
Private mblnVMove As Boolean ' picDraw���c�����ֈړ���/�s������Flag
Private mstrFileName As String ' ���ݕҏW����NC�t�@�C����(�v���p�e�B�p)
Private mintTHNT As Integer ' ���ݕҏW����NC��TH/NT��������(�v���p�e�B�p)
Private mdblScaleFactor As Double ' �s�N�`���{�b�N�X�ɕ\�����鎞�̃t�@�N�^(�v���p�e�B�p)

'*********************************************************
' �p  �r: ���ݍ�ƒ���NC�t�@�C����(NCFileName�v���p�e�B)�̎擾
' ��  ��: ����
' �߂�l: NCFileName�v���p�e�B�̒l
'*********************************************************

Public Property Get NCFileName() As String

    NCFileName = mstrFileName

End Property

'*********************************************************
' �p  �r: NCFileName�v���p�e�B�Ɍ��ݍ�ƒ���NC�t�@�C�������Z�b�g
' ��  ��: strFileName: ���ݍ�ƒ���NC�t�@�C����
' �߂�l: ����
'*********************************************************

Public Property Let NCFileName(ByVal strFileName As String)

    mstrFileName = strFileName

End Property

'*********************************************************
' �p  �r: ���ݍ�ƒ���NC��TH/NT���������l(THNT�v���p�e�B)�̎擾
' ��  ��: ����
' �߂�l: THNT�v���p�e�B�̒l
'*********************************************************

Public Property Get THNT() As Integer

    THNT = mintTHNT

End Property

'*********************************************************
' �p  �r: THNT�v���p�e�B�Ɍ��ݍ�ƒ���NC��TH/NT���������l���Z�b�g
' ��  ��: intTHNT: TH/NT�������l(0 or 1)
' �߂�l: ����
'*********************************************************

Public Property Let THNT(ByVal intTHNT As Integer)

    mintTHNT = intTHNT

End Property

'*********************************************************
' �p  �r: �s�N�`���{�b�N�X�ɕ\������ׂ̃X�P�[���t�@�N�^
'         (ScaleFactor�v���p�e�B)�̎擾
' ��  ��: ����
' �߂�l: ScaleFactor�v���p�e�B�̒l
'*********************************************************

Public Property Get ScaleFactor() As Double

    ScaleFactor = mdblScaleFactor

End Property

'*********************************************************
' �p  �r: ScaleFactor�v���p�e�B�Ƀs�N�`���{�b�N�X�ɕ\������
'         �ׂ̃X�P�[���t�@�N�^���Z�b�g
' ��  ��: dblScaleFactor: �X�P�[���t�@�N�^
' �߂�l: ����
'*********************************************************

Public Property Let ScaleFactor(ByVal dblScaleFactor As Double)

    mdblScaleFactor = dblScaleFactor

End Property

'*********************************************************
' �p  �r: frmMain��KeyDown�C�x���g
' ��  ��: KeyCode: �L�[ �R�[�h�������萔
'         Shift: �C�x���g��������Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
' �߂�l: ����
'*********************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then gblnCancel = False

End Sub

'*********************************************************
' �p  �r: frmMain��Unload�C�x���g
' ��  ��: Cancel: �t�H�[������ʂ���������邩�ǂ������w�肷��
'                 �����l(0�ŏ���, ���̑��͏������Ȃ�)
' �߂�l: ����
'*********************************************************

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Unload frmToolInfo

    If Dir(fTempPath & conTempFileName) <> "" Then
        Kill fTempPath & conTempFileName ' �e���|�����t�@�C�����폜
    End If

    If WindowState = vbNormal Then
        ' Form�̈ʒu�Ƒ傫�������W�X�g���ɕۑ�
        SaveSetting "NCView", "Position", "Top", top
        SaveSetting "NCView", "Position", "Left", left
        SaveSetting "NCView", "Position", "Height", Height
        SaveSetting "NCView", "Position", "Width", Width
    End If

    ' �ŏ����̎���WindowState��ۑ����Ȃ�
    If WindowState <> vbMinimized Then
        SaveSetting "NCView", "Position", "WindowState", WindowState
    End If

End Sub

'*********************************************************
' �p  �r: frmMain.mnuAbout��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuAbout_Click()

    Load frmAbout
    frmAbout.Show vbModal

End Sub

'*********************************************************
' �p  �r: frmMain.mnuInfo��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuInfo_Click()

    Load frmToolInfo
    frmToolInfo.Show vbModal

End Sub

'*********************************************************
' �p  �r: frmMain.mnuLook��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuLook_Click()

    Dim strNC As String
    Dim blnRet As Boolean
    Dim i As Integer

    gblnCancel = True

    ' NC��ǂݍ���
    strNC = fReadNC(Me.NCFileName)

    ' �R���g���[���̏����l�̐ݒ�
    Call sSetControl

    If Me.THNT = TH Then
        picDraw(0).Cls ' �S�̕\���p�s�N�`���[�{�b�N�X
        picDraw(1).Cls ' �����\���p�s�N�`���[�{�b�N�X
        Call sInit
    End If
    For i = 1 To intRow
        With gudtToolInfo(Me.THNT, i)
            Select Case .strColor
                Case "BLACK"
                    .lngColor = RGB(0, 0, 0)
                Case "RED"
                    .lngColor = RGB(255, 0, 0)
                Case "GREEN"
                    .lngColor = RGB(0, 200, 0)
                Case "BLUE"
                    .lngColor = RGB(0, 0, 200)
                Case "MAGENTA"
                    .lngColor = RGB(240, 0, 240)
                Case "CYAN"
                    .lngColor = RGB(0, 150, 255)
            End Select
        End With
    Next
    blnRet = fConvertNC(strNC, _
                        gudtNCInfo(Me.THNT), _
                        gudtToolInfo, _
                        gudtWBInfo, _
                        ProgressBar1)
    If blnRet = False Then Exit Sub
    Call sDispNC(picDraw, ProgressBar1)
    mblnDisp = True

End Sub

'*********************************************************
' �p  �r: frmMain.mnuNTIn��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuNTIn_Click()

    Dim strNCFileName As String
    Dim strInput As String
    Dim strRet As String
    Dim strNC As String
    Dim strXY() As String
    Dim strStack As String

    On Error GoTo Trap

    Me.THNT = NT ' NT��Tab���A�N�e�B�u�ɂ���ׂɂ����Őݒ肵�Ă���
    
    ' �ϐ����Z�b�g����Ă��Ȃ���΃v���Z�X�ԒʐM�����݂�
    With gudtWBInfo
        If .intSosu = 0 And _
           .lngWBS(X) = 0 And _
           .lngWBS(Y) = 0 And _
           .lngStack(X) = 0 And _
           .lngStack(Y) = 0 Then
            Call sDDElink(gudtWBInfo)
        End If
    End With

    If gudtNCInfo(TH).strFileName = "" Then
        MsgBox "TH��ݒ肵�Ȃ���NT�͍����o���܂���B"
        Exit Sub
    Else
        strNCFileName = fGetInputFile()
        If strNCFileName = "" Then Exit Sub
    End If

    Me.NCFileName = strNCFileName
    Caption = conCaption & " - " & Me.NCFileName

    Load frmToolInfo
    frmToolInfo.Show vbModal
    If frmToolInfo.NTLoadFlag = False Then Exit Sub

    With gudtWBInfo
        If .lngStack(Y) = 0 Then
            strRet = InputBox("�w����?", "�w���̓���")
            If strRet = "" Then Exit Sub
            .intSosu = CInt(strRet)
            If .intSosu > 2 Then
'                .lngStack(X) = 5& * int1mm
                strNC = fReadNC(strNCFileName) ' NT�̓ǂݍ���
                strRet = fGetNTIdou(strNC)
                If strRet <> "" Then
                    strXY = Split(Mid(strRet, 2), "Y", -1, vbTextCompare)
                    strStack = CLng(strXY(Y)) / int1mm
                End If
            Else ' ���ʔ̏ꍇ
'                .lngStack(X) = 4& * int1mm
                With gudtNCInfo(TH)
                    strNC = fReadNC(.strFileName) ' TH�̓ǂݍ���
                    strRet = fGetNTIdou(strNC)
                    If strRet <> "" Then
                        strXY = Split(Mid(strRet, 2), "Y", -1, vbTextCompare)
                        strStack = 180 - (CLng(strXY(Y)) / int1mm)
                    Else
                        strStack = "180"
                    End If
                End With
            End If
            strInput = InputBox("��mm�X�^�b�N�ł���?", "�X�^�b�N�̓���", strStack)
            .lngStack(Y) = CLng(CSng(strInput) * int1mm)
        End If
    End With

    Exit Sub

Trap:
    MsgBox "���̓G���[�ł�"

End Sub

'*********************************************************
' �p  �r: frmMain.mnuOpen��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuOpen_Click()

    Dim strNCFileName As String
    Dim intRet As Integer ' MsgBox�̖߂�l

    intRet = MsgBox("���݂̍H����͔j������܂��B", 49, "�m�F")
    If intRet = vbCancel Then Exit Sub

    ' �ϐ�������������
    Call sInitialize

    ' frmToolInfo�̃v���p�e�B������������
    With frmToolInfo
        .THLoadFlag = False
        .NTLoadFlag = False
    End With

    Me.THNT = TH ' TH��Tab���A�N�e�B�u�ɂ���ׂɂ����Őݒ肵�Ă���
    strNCFileName = fGetInputFile()
    If strNCFileName = "" Then Exit Sub

    Me.NCFileName = strNCFileName
    Caption = conCaption & " - " & Me.NCFileName

    Load frmToolInfo
    frmToolInfo.Show vbModal

End Sub

'*********************************************************
' �p  �r: frmMain.mnuOption��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuOption_Click()

    Load frmOption
    frmOption.Show vbModal

End Sub

'*********************************************************
' �p  �r: frmMain.mnuPrint��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuPrint_Click()

    Dim i As Integer

    ' �R���g���[���̏����l�̐ݒ�
    Call sSetControl

    For i = 1 To intRow
        With gudtToolInfo(Me.THNT, i)
            Select Case .strColor
                Case "BLACK"
                    .lngColor = RGB(0, 0, 0)
                Case "RED"
                    .lngColor = RGB(255, 0, 0)
                Case "GREEN"
                    .lngColor = RGB(0, 255, 0)
                Case "BLUE"
                    .lngColor = RGB(0, 0, 255)
                Case "MAGENTA"
                    .lngColor = RGB(255, 0, 255)
                Case "CYAN"
                    .lngColor = RGB(0, 255, 255)
            End Select
        End With
    Next

    Load frmPlot
    frmPlot.Show vbModal

End Sub

'*********************************************************
' �p  �r: frmMain.mnuQuit��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuQuit_Click()

    ' NCView�̏I��
    Unload Me
    End

End Sub

'*********************************************************
' �p  �r: frmMain.mnuSeisun��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuSeisun_Click()

    picDraw(0).Visible = False ' �S�̕\��
    picDraw(1).Visible = True ' �����\��

End Sub

'*********************************************************
' �p  �r: frmMain.mnuStandard��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuStandard_Click()

    picDraw(0).Visible = True ' �S�̕\��
    picDraw(1).Visible = False ' �����\��

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw()��KeyDown�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         KeyCode: �L�[ �R�[�h�������萔
'         Shift: �C�x���g��������Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyUp
            picDraw(Index).top = picDraw(Index).top - 200
        Case vbKeyDown
            picDraw(Index).top = picDraw(Index).top + 200
        Case vbKeyLeft
            picDraw(Index).left = picDraw(Index).left - 200
        Case vbKeyRight
            picDraw(Index).left = picDraw(Index).left + 200
    End Select

    ' �s�N�`���[�{�b�N�X���R���e�i����͂ݏo���Ȃ��l�ɂ���
    Call sPicPosition(Index)

End Sub

'*********************************************************
' �p  �r: frmMain.Toolbar1��ButtonClick�C�x���g
' ��  ��: Button: �N���b�N���ꂽ Button �I�u�W�F�N�g�ւ̎Q��
' �߂�l: ����
'*********************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

    Select Case Button.Key
        Case "�J��"
            Call mnuOpen_Click
        Case "���"
            Call mnuPrint_Click
        Case "�{��"
            Call mnuLook_Click
    End Select

End Sub

'*********************************************************
' �p  �r: frmMain��Load�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Load()

    ' �O��I������Form�̈ʒu�Ƒ傫���𕜌�
    top = GetSetting("NCView", "Position", "Top", "0")
    left = GetSetting("NCView", "Position", "Left", "0")
    Height = GetSetting("NCView", "Position", "Height", Height)
    Width = GetSetting("NCView", "Position", "Width", Width)
    WindowState = GetSetting("NCView", "Position", "WindowState", vbNormal)

    KeyPreview = True
    gblnCancel = True

    ' �v���O���X�o�[���\���ɂ���
    ProgressBar1.Visible = False

    ' �^�C�g��
    Caption = conCaption

    ' DDE�ʐM�p���x���R���g���[�����\���ɂ���
    lblDDE.Visible = False

    ' �t�H�[���̏�����
    Call sInit

    ' Load�C�x���g����ToolBar1.Height��"555"��Ԃ��̂�"360"�ŏ������B
    ' ����Ȃ̂ł����̂�...
    picDraw(0).Height = SysInfo1.WorkAreaHeight _
                        - (GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY) _
                        - 360 _
                        - (GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY)

    ' �s�N�`���[�{�b�N�X�̕\��/��\���̐ݒ�
    picDraw(0).Visible = True
    picDraw(1).Visible = False

    If Command = "" Then
        Me.NCFileName = fGetInputFile()
        If Me.NCFileName = "" Then
            Unload Me
            End
        End If
    Else
        Me.NCFileName = Command
    End If

    Caption = conCaption & " - " & Me.NCFileName

End Sub

'*********************************************************
' �p  �r: frmMain��Resize�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Resize()

    With picFrame
        .Align = 3 ' ������
        .Align = 1 ' �㑵��
    End With

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw��MouseDown�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         Button: �����ꂽ�{�^�������������l
'         Shift: �{�^���������ꂽ����Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
'         X, Y: �}�E�X�|�C���^�̌��݈ʒu��\�����l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim udtRect As tagRECT

    If Button = 1 Then
        MousePointer = vbCustom '�}�E�X�J�[�\����ύX
        mblnPanMode = True
        mblnHMove = True
        mblnVMove = True
        msngDragDistX = X
        msngDragDistY = Y
        msngCurrentTop = picDraw(Index).top
        msngCurrentLeft = picDraw(Index).left

        ' �s�N�`���{�b�N�X�̋�`�̈���擾
        GetWindowRect picFrame.hwnd, udtRect
        ' �擾�����̈�Ƀ}�E�X�̈ړ��͈͂𐧌�
        ClipCursor udtRect
    ElseIf Button = 2 Then
        PopupMenu mnuView
    End If

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw��MouseMove�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         Button: �����ꂽ�{�^�������������l
'         Shift: Shift, Ctrl, Alt�L�[�̏�Ԃ����������l
'         X, Y: �}�E�X�|�C���^�̌��݈ʒu��\�����l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mblnPanMode = False Then Exit Sub

    Dim udtRect As tagRECT
    Dim dblFactor As Double

    If Index = 1 Then
        dblFactor = 56.7
    Else
        dblFactor = 1 / Me.ScaleFactor
    End If

    ' �s�N�`���[�{�b�N�X���R���e�i����͂ݏo���Ȃ��l�ɂ���
    Call sPicPosition(Index)

    ' left, top�v���p�e�B��twip�P�ʂł��鎖�ɒ���!
    If mblnHMove = True Then
        picDraw(Index).left = _
            -(msngDragDistX - X) * dblFactor + msngCurrentLeft
    End If
    If mblnVMove = True Then
        picDraw(Index).top = _
            (msngDragDistY - Y) * dblFactor + msngCurrentTop
    End If
    msngCurrentLeft = picDraw(Index).left
    msngCurrentTop = picDraw(Index).top

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw��MouseUp�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         Button: �����ꂽ�{�^�������������l
'         Shift: �����ꂽ����Shift, Ctrl, Alt�L�[�̏�Ԃ�
'                ���������l
'         X, Y: �}�E�X�|�C���^�̌��݈ʒu��\�����l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mblnPanMode = False
    mblnHMove = False
    mblnVMove = False

    If Button = 1 Then ' ���{�^��
        MousePointer = vbDefault ' �}�E�X�J�[�\�����f�t�H���g�ɖ߂�

        ' ������NULL���w�肷�邱�Ƃ�
        ' �}�E�X�J�[�\���̈ړ�����������
        ClipCursor ByVal 0
    End If

End Sub

'*********************************************************
' �p  �r: �s�N�`���{�b�N�X�̏�����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sInit()

    mblnDisp = False

    With picFrame
        .Align = 3 ' ������
        .Align = 1 ' �㑵��
    End With

    With picDraw(0) ' �S�̕\���p�s�N�`���[�{�b�N�X
        ' �w�i����
        .BackColor = RGB(250, 250, 250)
        .ForeColor = QBColor(0) ' ��
        picDraw(0).Width = SysInfo1.WorkAreaWidth
        picDraw(0).Height = SysInfo1.WorkAreaHeight _
                            - (GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY) _
                            - Toolbar1.Height _
                            - (GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY)
        .top = -24
        .left = -24
        .AutoRedraw = True
        .ScaleHeight = -Abs(.ScaleHeight)
        .Appearance = 0 ' �t���b�g
'        .Visible = True
    End With


    With picDraw(1) ' �����\���p�s�N�`���{�b�N�X
        ' �w�i����
        .BackColor = RGB(250, 250, 250)
        .ForeColor = QBColor(0) ' ��
        .top = -24
        .left = -24
        .AutoRedraw = True
        .Width = picFrame.Width
        .Height = picFrame.Height
        .ScaleHeight = -Abs(.ScaleHeight)
        .Appearance = 0 ' �t���b�g
'        .Visible = False ' �N�����͔�\��
    End With

End Sub

'*********************************************************
' �p  �r: �v���O���X�o�[�̏�����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sSetControl()

    ' �v���O���X�o�[�̃v���p�e�B�̐ݒ�
    With ProgressBar1
        .Width = 3000
        .Height = Toolbar1.Height - 36
        .top = 36
        .left = picFrame.Width - .Width
    End With

    ' �v���Z�X�ԒʐM�����݂�
    Call sDDElink(gudtWBInfo)

End Sub

Private Sub sPicPosition(Index As Integer)

'*********************************************************
' �p  �r: �s�N�`���[�{�b�N�X���R���e�i����͂ݏo���Ȃ��l�ɂ���
' ��  ��: ����
' �߂�l: ����
'*********************************************************

    With picDraw(Index)
        If .Width >= picFrame.Width Then
            If .left > -24 Then
                mblnHMove = False
                .left = -24
            ElseIf .left < picFrame.Width - .Width - 24 Then
                mblnHMove = False
                .left = picFrame.Width - .Width - 24
            End If
        ElseIf .Width < picFrame.Width Then
            If .left < -24 Then ' ����
                mblnHMove = False
                .left = -24
            ElseIf .left > picFrame.Width - .Width - 24 Then
                mblnHMove = False
                .left = picFrame.Width - .Width - 24
            End If
        End If
        If .Height >= picFrame.Height Then
            If .top > -24 Then
                mblnVMove = False
                .top = -24
            ElseIf .top < picFrame.Height - .Height - 24 Then
                mblnVMove = False
                .top = picFrame.Height - .Height - 24
            End If
        ElseIf .Height < picFrame.Height Then
            If .top < -24 Then
                mblnVMove = False
                .top = -24
            ElseIf .top > picFrame.Height - .Height - 24 Then
                mblnVMove = False
                .top = picFrame.Height - .Height - 24
            End If
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �t�@�C�����J���_�C�A���O��\������
' ��  ��: ����
' �߂�l: �I�������t�@�C����
'*********************************************************

Public Function fGetInputFile() As String

    With CommonDialog1
        ' CancelError�v���p�e�B��^(True)�ɐݒ肵�܂��B
        .CancelError = True
        On Error GoTo ErrHandler

        ' �t�@�C���̑I����@��ݒ肵�܂��B
        .Filter = "���ׂẴt�@�C�� (*.*)|*.*|" & _
                  "�f�[�^�t�@�C�� (*.dat)|*.dat|" & _
                  "NC�f�[�^�t�@�C�� (*.nc)|*.nc"

        ' ����̑I����@���w�肵�܂��B
        .FilterIndex = 1

        ' [�ǂݎ���p�t�@�C���Ƃ��ĊJ��]�`�F�b�N�{�b�N�X��\�����Ȃ�
        ' �����̃t�@�C�����������͂ł��Ȃ��悤�ɂ���
        .Flags = cdlOFNHideReadOnly Or _
                 cdlOFNFileMustExist

        ' [�t�@�C�����J��] �_�C�A���O �{�b�N�X��\�����܂��B
        .ShowOpen

        fGetInputFile = .FileName
        Exit Function
    End With

ErrHandler:
        '���[�U�[��[�L�����Z��] �{�^�����N���b�N���܂����B
'       If Err.Number = cdlCancel Then
'           If Mid(Caption, Len(conCaption) + 4) = "" Then
'               Caption = conCaption
'           End If
'       End If
    fGetInputFile = ""

End Function
