VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "��߼��"
   ClientHeight    =   2310
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.CommandButton cmdCansel 
      Caption         =   "��ݾ�"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin TabDlg.SSTab mstOption 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2778
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "ܰ��ް��"
      TabPicture(0)   =   "frmOption.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStart"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPitch"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblWBS"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSosu"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbStack"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPitch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtWBS(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtWBS(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtSosu"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "HPGL�ϊ���۸���"
      TabPicture(1)   =   "frmOption.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtHPGLCmd"
      Tab(1).Control(1)=   "lblHPGLCmd"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtHPGLCmd 
         Height          =   270
         Left            =   -74760
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtSosu 
         Height          =   264
         Left            =   720
         TabIndex        =   2
         Text            =   "999"
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox txtWBS 
         Height          =   264
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Text            =   "999999"
         Top             =   600
         Width           =   732
      End
      Begin VB.TextBox txtWBS 
         Height          =   264
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Text            =   "999999"
         Top             =   600
         Width           =   732
      End
      Begin VB.TextBox txtPitch 
         Height          =   264
         Left            =   720
         TabIndex        =   7
         Text            =   "999999"
         Top             =   960
         Width           =   732
      End
      Begin VB.ComboBox cmbStack 
         Height          =   300
         ItemData        =   "frmOption.frx":0044
         Left            =   2040
         List            =   "frmOption.frx":0051
         TabIndex        =   9
         Text            =   "��������/����"
         Top             =   960
         Width           =   1692
      End
      Begin VB.Label lblHPGLCmd 
         Caption         =   "HPGL�ϊ���۸��і�"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblSosu 
         Caption         =   "�w��"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblWBS 
         Caption         =   "WBS"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblPitch 
         Caption         =   "�߯�"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblStart 
         Caption         =   "����"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPitch As Long

'*********************************************************
' �p  �r: �L�����Z���{�^���̃N���b�N�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdCansel_Click()

    Unload Me

End Sub

'*********************************************************
' �p  �r: OK�{�^���̃N���b�N�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdOK_Click()

    With gudtWBInfo
        If txtWBS(X).Text <> "" Then .lngWBS(X) = CLng(txtWBS(X)) * int1mm
        If txtWBS(Y).Text <> "" Then .lngWBS(Y) = CLng(txtWBS(Y)) * int1mm
        If txtSosu.Text = "" Then
            ' �������Ȃ�
        ElseIf CInt(txtSosu.Text) > 2 Then ' ���w��
            mlngPitch = CLng(txtPitch.Text) * int1mm
            .lngStack(X) = (.lngWBS(X) - mlngPitch) / 2
            If cmbStack.Text = "��������/����" Then
                .lngStack(Y) = CSng(txtWBS(Y).Text) * int1mm / 2
            Else
                .lngStack(Y) = CSng(cmbStack.Text) * int1mm
            End If
            ' ���w��NC���_�ƃX�^�b�N�ʒu������
            .strStart = "STACK"
        Else ' ���ʔ�
            .lngStack(X) = 400&
            If cmbStack.Text = "��������/����" Then
                .lngStack(Y) = CSng(txtWBS(Y).Text) * int1mm / 2
                .strStart = "STACK"
            Else
                .lngStack(Y) = CSng(cmbStack.Text) * int1mm
                .strStart = "MACHINE"
            End If
        End If
    End With

    Unload Me

End Sub

'*********************************************************
' �p  �r: frmPlot��Load�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Load()

    Dim strHPGLCommand As String

    txtHPGLCmd.Text = GetSetting("NCView", _
                                 "Settings", _
                                 "HPGLCommand", _
                                 conDefaultHPGLCommand)

    ' �v���Z�X�ԒʐM�����݂�
    Call sDDElink(gudtWBInfo)

    ' �e�L�X�g�{�b�N�X�̏�����
    With txtSosu
        .Text = ""
        .MaxLength = 3
        .ToolTipText = "�w��"
    End With
    With txtWBS(X)
        .Text = ""
        .MaxLength = 6
        .ToolTipText = "X��WBS"
    End With
    With txtWBS(Y)
        .Text = ""
        .MaxLength = 6
        .ToolTipText = "Y��WBS"
    End With
    With txtPitch
        .Text = ""
        .MaxLength = 6
        .ToolTipText = "�U�O���ԃs�b�`"
    End With

    ' �R���{�{�b�N�X�̏�����
    With cmbStack
        .Text = ""
        .ToolTipText = "�X�^�b�N�ʒu"
    End With

    ' �ϐ��ɒl���Z�b�g����Ă�����e�L�X�g�{�b�N�X�ɃZ�b�g����
    Call sSetTextBox(gudtWBInfo)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' HPGL�ϊ��v���O�����������W�X�g���ɕۑ�
    If txtHPGLCmd.Text <> "" Then
        SaveSetting "NCView", _
                    "Settings", _
                    "HPGLCommand", _
                    txtHPGLCmd.Text
    End If

End Sub

'*********************************************************
' �p  �r: �s�b�`���͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtPitch_GotFocus()

    With txtPitch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'*********************************************************
' �p  �r: �s�b�`���͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
'         ���Ғʂ�̒l�����͂���Ă��邩�`�F�b�N����
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���ێ����邩���肷��
'                 True�ňێ�
' �߂�l: ����
'*********************************************************

Private Sub txtPitch_Validate(Cancel As Boolean)

    ' ���͂��ꂽ�l���`�F�b�N
    With txtPitch
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �w�����͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtSosu_GotFocus()

    With txtSosu
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'*********************************************************
' �p  �r: �w��, WBS�ɉ������X�^�b�N�ʒu�����肷��
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sSetStack()

    Dim lngWBSY As Long

    ' �ϐ��ɑw�����Z�b�g
    With gudtWBInfo
        If txtSosu.Text <> "" Then
            .intSosu = CInt(txtSosu)
            If .intSosu <= 2 Then
                With txtPitch
                    .Text = ""
                    .Enabled = False
                    .BackColor = &H8000000F
                End With
                lblPitch.Enabled = False
            Else
                If txtWBS(X).Text <> "" And txtPitch.Text = "" Then
                    txtPitch.Text = CLng(txtWBS(X).Text) - 10
                End If
                With txtPitch
                    .Enabled = True
                    .BackColor = &H80000005
                End With
                lblPitch.Enabled = True
            End If
        End If
    End With

    If txtWBS(Y).Text = "" Then
        lngWBSY = 0&
    Else
        lngWBSY = CLng(txtWBS(Y))
    End If

    With cmbStack
        If gudtWBInfo.lngStack(Y) <> 0 Then
            .Text = gudtWBInfo.lngStack(Y) / 100
        ElseIf gudtWBInfo.intSosu <= 2 Then
            If Dir(Command) Like "AMS*" Then
                .Text = "180" ' AMS�i��180�X�^�b�N
            ElseIf lngWBSY > 500 Then
                .Text = "��������/����"
            ElseIf lngWBSY >= 400 Then
                .Text = "205"
            ElseIf lngWBSY <> 0 Then
                .Text = "180"
            End If
        Else
            .Text = "��������/����"
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �w�����͗p�e�L�X�g�{�b�N�X��LostFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtSosu_LostFocus()

    ' �s�b�`/�X�^�b�N���ʒu�̐ݒ�
    Call sSetStack

End Sub

'*********************************************************
' �p  �r: �w�����͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
'         ���Ғʂ�̒l�����͂���Ă��邩�`�F�b�N
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���ێ����邩���肷��
'                 True�ňێ�
' �߂�l: ����
'*********************************************************

Private Sub txtSosu_Validate(Cancel As Boolean)

    ' ���͂��ꂽ�l���`�F�b�N
    With txtSosu
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: WBS���͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: Index: XY������0�܂���1�����ꂩ�̐��l
' �߂�l: ����
'*********************************************************

Private Sub txtWBS_GotFocus(Index As Integer)

    With txtWBS(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'*********************************************************
' �p  �r: WBS���͗p�e�L�X�g�{�b�N�X��LostFocus�C�x���g
' ��  ��: Index: XY������0�܂���1�����ꂩ�̐��l
' �߂�l: ����
'*********************************************************

Private Sub txtWBS_LostFocus(Index As Integer)

    ' �s�b�`/�X�^�b�N���ʒu�̐ݒ�
    Call sSetStack

End Sub

'*********************************************************
' �p  �r: WBS���͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
'         ���Ғʂ�̒l�����͂���Ă��邩�`�F�b�N
' ��  ��: Index: XY������0�܂���1�����ꂩ�̐��l
'         Cancel: �R���g���[�����t�H�[�J�X���ێ����邩���肷��
'                 True�ňێ�
' �߂�l: ����
'*********************************************************

Private Sub txtWBS_Validate(Index As Integer, Cancel As Boolean)

    ' ���͂��ꂽ�l���`�F�b�N
    With txtWBS(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �e�L�X�g�{�b�N�X�ɑw��, WBS etc...�̒l��������
' ��  ��: udtWBInfo: WB��񂪊i�[���ꂽ�\����WBInfo
' �߂�l: ����
'*********************************************************

Private Sub sSetTextBox(udtWBInfo As WBInfo)

    ' �ϐ��ɒl���Z�b�g����Ă�����e�L�X�g�{�b�N�X�ɃZ�b�g����
    With udtWBInfo
        If .intSosu <> 0 Then txtSosu.Text = .intSosu
        If .lngWBS(X) <> 0 Then txtWBS(X).Text = .lngWBS(X) / 100
        If .lngWBS(Y) <> 0 Then txtWBS(Y).Text = .lngWBS(Y) / 100
        If .lngStack(Y) <> 0 Then cmbStack.Text = .lngStack(Y) / 100
        If mlngPitch <> 0 Then txtPitch.Text = mlngPitch / 100
    End With

    Call sSetStack

End Sub
