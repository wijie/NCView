VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmToolInfo 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�H����"
   ClientHeight    =   5775
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmToolInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin TabDlg.SSTab SSTab1 
      Height          =   5052
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4812
      _ExtentX        =   8493
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "TH"
      TabPicture(0)   =   "frmToolInfo.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMaxY(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMinY(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblMinX(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTotal(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblMaxX(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "msgDrill(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtInput(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgcboColor(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "NT"
      TabPicture(1)   =   "frmToolInfo.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblTotal(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblMinX(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblMinY(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblMaxX(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblMaxY(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "msgDrill(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "imgcboColor(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtInput(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtInput 
         Height          =   264
         Index           =   1
         Left            =   -74760
         TabIndex        =   12
         Text            =   "000,000"
         Top             =   960
         Width           =   732
      End
      Begin MSComctlLib.ImageCombo imgcboColor 
         Height          =   315
         Index           =   1
         Left            =   -73920
         TabIndex        =   13
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "�K��"
         ImageList       =   "ImageList1"
      End
      Begin MSComctlLib.ImageCombo imgcboColor 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "�K��"
         ImageList       =   "ImageList1"
      End
      Begin MSFlexGridLib.MSFlexGrid msgDrill 
         Height          =   4560
         Index           =   1
         Left            =   -74880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   2868
         _ExtentX        =   5054
         _ExtentY        =   8043
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtInput 
         Height          =   264
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Text            =   "000,000"
         Top             =   960
         Width           =   672
      End
      Begin MSFlexGridLib.MSFlexGrid msgDrill 
         Height          =   4560
         Index           =   0
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   2868
         _ExtentX        =   5054
         _ExtentY        =   8043
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMaxY 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71040
         TabIndex        =   22
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lblMaxX 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71880
         TabIndex        =   21
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lblMinY 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71040
         TabIndex        =   19
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label lblMinX 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71880
         TabIndex        =   18
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  '����
         Caption         =   "999,999"
         Height          =   252
         Index           =   1
         Left            =   -71880
         TabIndex        =   16
         Top             =   720
         Width           =   732
      End
      Begin VB.Label Label15 
         Caption         =   "�ŏ��l"
         Height          =   252
         Left            =   -71880
         TabIndex        =   17
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label14 
         Caption         =   "�������v"
         Height          =   252
         Left            =   -71880
         TabIndex        =   15
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label12 
         Caption         =   "�ő�l"
         Height          =   252
         Left            =   -71880
         TabIndex        =   20
         Top             =   1920
         Width           =   612
      End
      Begin VB.Label lblMaxX 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3120
         TabIndex        =   10
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  '����
         Caption         =   "999,999"
         Height          =   252
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   732
      End
      Begin VB.Label lblMinX 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "�ő�l"
         Height          =   252
         Left            =   3120
         TabIndex        =   9
         Top             =   1920
         Width           =   612
      End
      Begin VB.Label lblMinY 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3960
         TabIndex        =   8
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label Label7 
         Caption         =   "�������v"
         Height          =   252
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "�ŏ��l"
         Height          =   252
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label lblMaxY 
         BorderStyle     =   1  '����
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3960
         TabIndex        =   11
         Top             =   2160
         Width           =   732
      End
   End
   Begin VB.CommandButton cmdCansel 
      Caption         =   "��ݾ�"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   5280
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5160
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0144
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0244
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0344
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0444
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolInfo.frx":0644
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmToolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnTHLoadFlag As Boolean ' ���߂ă��[�h����̂��ۂ��������t���O(TH)
Private mblnNTLoadFlag As Boolean ' ���߂ă��[�h����̂��ۂ��������t���O(NT)
Private mlngTotal(1) As Long
Private mudtToolInfo(1, 1 To intRow) As New frmToolInfo ' TH/NT�̃c�[�����
Private mblnKeyFlag As Boolean ' �e�L�X�g�{�b�N�X�Ńr�[�v������Ȃ��l�ɂ���t���O

'*********************************************************
' �p  �r: TH��NC�����[�h�ς݂��ۂ��������l
'         (THLoadFlag�v���p�e�B)�̎擾
' ��  ��: ����
' �߂�l: �v���p�e�B�̒l
'*********************************************************

Public Property Get THLoadFlag() As Boolean

    THLoadFlag = mblnTHLoadFlag

End Property

'*********************************************************
' �p  �r: THLoadFlag�v���p�e�B�Ƀ��[�h�ς݂������l���Z�b�g
' ��  ��: blnTHLoadFlag: ���[�h�ς�-True, ��-False
' �߂�l: ����
'*********************************************************

Public Property Let THLoadFlag(ByVal blnTHLoadFlag As Boolean)

    mblnTHLoadFlag = blnTHLoadFlag

End Property

'*********************************************************
' �p  �r: NT��NC�����[�h�ς݂��ۂ��������l
'         (NTLoadFlag�v���p�e�B)�̎擾
' ��  ��: ����
' �߂�l: �v���p�e�B�̒l
'*********************************************************

Public Property Get NTLoadFlag() As Boolean

    NTLoadFlag = mblnNTLoadFlag

End Property

'*********************************************************
' �p  �r: NTLoadFlag�v���p�e�B�Ƀ��[�h�ς݂������l���Z�b�g
' ��  ��: blnNTLoadFlag: ���[�h�ς�-True, ��-False
' �߂�l: ����
'*********************************************************

Public Property Let NTLoadFlag(ByVal blnNTLoadFlag As Boolean)

    mblnNTLoadFlag = blnNTLoadFlag

End Property

'*********************************************************
' �p  �r: OK�{�^���̃N���b�N�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdOK_Click()

    Dim strColor As String
    Dim i As Integer
    Dim intTHNT As Integer
    Dim intColor As Integer
    Dim varColorList() As Variant

    Me.Hide ' �t�H�[�����\���ɂ���

    ' �K���Ȏ���,���̃��X�g���珇�ԂɑI��
    varColorList = Array("RED", _
                         "BLUE", _
                         "MAGENTA", _
                         "CYAN")

    With frmMain
        If Me.SSTab1.Tab = TH Then
            frmMain.THNT = TH
            gudtNCInfo(TH).strFileName = .NCFileName
            Me.THLoadFlag = True
        Else
            frmMain.THNT = NT
            gudtNCInfo(NT).strFileName = .NCFileName
            Me.NTLoadFlag = True
        End If
    End With

    intColor = 0
    For intTHNT = 0 To 1
        With msgDrill(intTHNT)
            For i = 1 To intRow
                .Row = i
                .Col = 1 ' �h�����a
                If .Text <> "" Then
                    .Col = 3 ' �F
                    Select Case .Text
                        Case "���K��"
                            strColor = varColorList(intColor)
                            intColor = intColor + 1
                            If intColor > UBound(varColorList) Then
                                intColor = 0
                            End If
                        Case "����"
                            strColor = "BLACK"
                        Case "����"
                            strColor = "RED"
                        Case "����"
                            strColor = "GREEN"
                        Case "����"
                            strColor = "BLUE"
                        Case "��Ͼ���"
                            strColor = "MAGENTA"
                        Case "�����"
                            strColor = "CYAN"
                    End Select
                    .Col = 0 ' TNo
                    gudtToolInfo(intTHNT, i).intTNo = CInt(.Text)
                    .Col = 1 ' �h�����a
                    gudtToolInfo(intTHNT, i).sngDrill = CSng(.Text) / 2
                    .Col = 3 ' �F
                    gudtToolInfo(intTHNT, i).strColor = strColor
                End If
            Next
        End With
    Next

    Unload Me

End Sub

'*********************************************************
' �p  �r: �L�����Z���{�^���̃N���b�N�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdCansel_Click()

    Unload Me

End Sub

'*********************************************************
' �p  �r: frmToolInfo��Load�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Load()

    Dim strNC As String

    mblnKeyFlag = True
    mlngTotal(TH) = 0
    mlngTotal(NT) = 0

    Call sInit(NT) ' �t�H�[���̏�����
    Call sInit(TH) ' �t�H�[���̏�����

    If frmMain.THNT = TH Then
        SSTab1.Tab = TH
        If Me.THLoadFlag = False Then
            strNC = fReadNC(frmMain.NCFileName) ' �ǂނ͍̂ŏ���1�񂾂�
            Call sSetUsedTool(strNC) ' �g���Ă���c�[���𒲂ׂ�
        End If
    Else
        SSTab1.Tab = NT
        If Me.NTLoadFlag = False Then
            strNC = fReadNC(frmMain.NCFileName) ' �ǂނ͍̂ŏ���1�񂾂�
            Call sSetUsedTool(strNC) ' �g���Ă���c�[���𒲂ׂ�
        End If
    End If
    SSTab1.TabStop = False

End Sub

'*********************************************************
' �p  �r: �F�I��p�C���[�W�R���{��Click�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub imgcboColor_Click(Index As Integer)

    With msgDrill(Index)
        Select Case imgcboColor(Index).SelectedItem.Text
            Case "��"
                .CellForeColor = RGB(0, 220, 0)
                .Text = "����"
            Case "��"
                .CellForeColor = RGB(0, 0, 0)
                .Text = "����"
            Case "��"
                .CellForeColor = RGB(255, 0, 0)
                .Text = "����"
            Case "��"
                .CellForeColor = RGB(0, 0, 255)
                .Text = "����"
            Case "Ͼ���"
                .CellForeColor = RGB(255, 0, 255)
                .Text = "��Ͼ���"
            Case "���"
                .CellForeColor = RGB(0, 220, 220)
                .Text = "�����"
            Case Else
                .CellForeColor = RGB(0, 0, 0)
                .Text = "���K��"
        End Select
    End With

End Sub

'*********************************************************
' �p  �r: �F�I��p�C���[�W�R���{��GotFocus�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub imgcboColor_GotFocus(Index As Integer)

'    SendKeys "{F4}"

End Sub

'*********************************************************
' �p  �r: �t���L�V�u���O���b�h��Click�C�x���g
' ��  ��: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub msgDrill_Click(Index As Integer)

    With msgDrill(Index)
        If .Col = 0 Then
            txtInput(Index).MaxLength = 3
        ElseIf .Col = 1 Then
            txtInput(Index).MaxLength = 5
        End If
    End With

    Select Case msgDrill(Index).Col
        Case 3 ' �F�̌�
            imgcboColor(Index).Visible = True
            With msgDrill(Index)
                txtInput(Index).Visible = False
                imgcboColor(Index).top = .CellTop + .top
                imgcboColor(Index).SetFocus
            End With
            With imgcboColor(Index).ComboItems
                Select Case msgDrill(Index).Text
                    Case "����"
                        .Item(1).Selected = True
                    Case "����"
                        .Item(2).Selected = True
                    Case "����"
                        .Item(3).Selected = True
                    Case "����"
                        .Item(4).Selected = True
                    Case "��Ͼ���"
                        .Item(5).Selected = True
                    Case "�����"
                        .Item(6).Selected = True
                    Case Else
                        .Item(7).Selected = True
                End Select
            End With
        Case 2 ' �����̌�
            txtInput(Index).Visible = False
            imgcboColor(Index).Visible = False
        Case Else
            imgcboColor(Index).Visible = False
            With txtInput(Index)
                .Width = msgDrill(Index).CellWidth
                .Height = msgDrill(Index).CellHeight
                .Text = msgDrill(Index).Text
                .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
                      msgDrill(Index).CellTop + msgDrill(Index).top
                .SelStart = 0
                .SelLength = Len(.Text)
                .Visible = True
                .SetFocus
            End With
    End Select

End Sub

'*********************************************************
' �p  �r: �t���L�V�u���O���b�h��Scroll�C�x���g
' ��  ��: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub msgDrill_Scroll(Index As Integer)

    ' �R���g���[�����\�������O�ɃC�x���g����������ƃG���[�ɂȂ�̂Ńg���b�v����(-_-;
    On Error GoTo bye

    msgDrill(Index).SetFocus ' TextBox��Focus�����鎞��Scroll�����Focus���R�}���h�{�^���ɔ��ł��܂���
    txtInput(Index).Visible = False
    imgcboColor(Index).Visible = False

bye:

End Sub

'*********************************************************
' �p  �r: �^�u�R���g���[����Click�C�x���g
' ��  ��: PreviousTab: �؂�ւ��O�̃^�u��Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub SSTab1_Click(PreviousTab As Integer)

'    txtInput(SSTab1.Tab).SetFocus

End Sub

'*********************************************************
' �p  �r: ���͗p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtInput_Change(Index As Integer)

    With msgDrill(Index)
        .CellAlignment = 1
        .Text = txtInput(Index).Text
    End With

End Sub

'*********************************************************
' �p  �r: ���͗p�e�L�X�g�{�b�N�X��KeyDown�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
'         KeyCode: �L�[ �R�[�h�������萔
'         Shift: �C�x���g��������Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
' �߂�l: ����
'*********************************************************

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    ' �������͂���Ă��Ȃ����͎��ɐi�܂Ȃ�(�߂�̂͋�����)
    If txtInput(Index).Text = "" And KeyCode <> vbKeyUp Then Exit Sub

    ' Enter����, Ctrl-M����, �����L�[
    With msgDrill(Index)
        If KeyCode = vbKeyReturn Or _
           (Shift = 2 And KeyCode = vbKeyM) Or _
           KeyCode = vbKeyDown Then
                mblnKeyFlag = False
                .Text = txtInput(Index).Text
                If .Row < intRow - 1 Then
                    .Row = .Row + 1
                End If
        ElseIf KeyCode = vbKeyUp Then
            With msgDrill(Index)
                If .Row > 1 Then
                    .Row = .Row - 1
                End If
            End With
        Else
            Exit Sub
        End If

        If .Col = 0 Then ' TNo.
            txtInput(Index).MaxLength = 3
        Else
            txtInput(Index).MaxLength = 5
        End If

        With txtInput(Index)
            With txtInput(Index)
                .Width = msgDrill(Index).CellWidth
                .Height = msgDrill(Index).CellHeight
                .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
                      msgDrill(Index).CellTop + msgDrill(Index).top
                    .Visible = True
            End With
            .SetFocus
            .Text = msgDrill(Index).Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End With

End Sub

'*********************************************************
' �p  �r: �g�p����Ă���T�R�[�h�𒲂ׂăO���b�h�R���g���[����
'         �Z�b�g����
' ��  ��: strNC: NC�f�[�^
' �߂�l: ����
'*********************************************************

Private Sub sSetUsedTool(ByVal strNC As String)

    Dim i As Integer
    Dim objReg As New RegExp
    Dim objMatches As Object
    Dim objMatch As Object

    objReg.Global = True
    objReg.IgnoreCase = False ' �啶������������ʂ���
    objReg.Pattern = "T[0-9]+"
    Set objMatches = objReg.Execute(strNC)

    ' T�R�[�h���H����ɃZ�b�g����
    i = 1
    With Me.msgDrill(frmMain.THNT)
        For Each objMatch In objMatches
            .Row = i
            .Col = 0
            .Text = Mid(objMatch.Value, 2)
            ' T50�̃f�t�H���g��ݒ肷��
            If objMatch = "T50" Then
                .Col = 1 ' �h�����a
                With frmToolInfo.txtInput(frmMain.THNT)
                    .Text = "1.999"
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
                .Col = 3 ' �F�̌�
                .CellForeColor = RGB(0, 0, 0)
                .Text = "����"
            End If
            i = i + 1
        Next
        ' �f�t�H���g�̈ʒu�ɃZ�b�g
        .Row = 1
        .Col = 1
    End With

End Sub

'*********************************************************
' �p  �r: �R���g���[���̏�����
' ��  ��: Index: TH/NT�������l(TH = 0, NT = 1)
' �߂�l: ����
'*********************************************************

Private Sub sInit(Index As Integer)

    Dim i As Integer

    With msgDrill(Index) ' �O���b�h�̏�����
        .Cols = 4
        .Rows = intRow + 1 ' +1���Ă���̂͌Œ�s�������
        .FixedCols = 0 ' �Œ��Ȃ�
        .FixedRows = 1 ' �Œ�s1
        .Width = 2880
        .Height = 4560
        .RowHeight(-1) = 288 ' �S��̍���
        .RowHeight(0) = 240 ' �Œ��̍���
        .ColWidth(0) = 456 ' TNo.�̌���
        .ColWidth(1) = 624 ' �h�����a�̌���
        .ColWidth(2) = 672 ' �����̌���
        .ColWidth(3) = 780 ' �F�̌���
'        .FillStyle = flexFillRepeat
        .FocusRect = flexFocusNone ' �t�H�[�J�X����������\�����Ȃ�
        .HighLight = flexHighlightNever ' �����\�����Ȃ�
        .Row = 0 ' ���
        .Col = 0
        .Text = "TNo."
        .Col = 1
        .Text = "���ٌa"
        .Col = 2
        .Text = "����"
        .Col = 3
        .Text = "�F"
        For i = 1 To intRow
            .Row = i
            .Col = 0 ' TNo�̌�
            .CellAlignment = 1 ' �����̒���
            If gudtToolInfo(Index, i).intTNo > -1 Then
                .Text = gudtToolInfo(Index, i).intTNo
            End If
            .Col = 1 ' �h�����a�̌�
            .CellAlignment = 1 ' �����̒���
            If gudtToolInfo(Index, i).sngDrill > -1 Then
                .Text = Format(gudtToolInfo(Index, i).sngDrill * 2, "##0.000")
            End If
            .Col = 2 ' �����̌�
            If gudtToolInfo(Index, i).intTNo = 50 Then ' �t�Z�b�g�`�F�b�N�p�c�[��
                ' �����͕\�������ō��v�����ɂ͐����Ȃ�
                .Text = "(" & gudtToolInfo(Index, i).lngCount & ")"
            ElseIf gudtToolInfo(Index, i).lngCount > -1 Then
                .Text = Format(gudtToolInfo(Index, i).lngCount, "##,##0")
                mlngTotal(Index) = mlngTotal(Index) + gudtToolInfo(Index, i).lngCount
            End If
            .Col = 3 ' �F�̌�
            Select Case gudtToolInfo(Index, i).strColor
                Case "GREEN" ' �C���[�W�R���{��1�Ԗ�
                    .CellForeColor = RGB(0, 220, 0)
                    .Text = "����"
                Case "BLACK" ' 2�Ԗ�
                    .CellForeColor = RGB(0, 0, 0)
                    .Text = "����"
                Case "RED" ' 3�Ԗ�
                    .CellForeColor = RGB(255, 0, 0)
                    .Text = "����"
                Case "BLUE" ' 4�Ԗ�
                    .CellForeColor = RGB(0, 0, 255)
                    .Text = "����"
                Case "MAGENTA" ' 5�Ԗ�
                    .CellForeColor = RGB(255, 0, 255)
                    .Text = "��Ͼ���"
                Case "CYAN" ' 6�Ԗ�
                    .CellForeColor = RGB(0, 220, 220)
                    .Text = "�����"
                Case Else
                    .CellForeColor = RGB(0, 0, 0)
                    .Text = "���K��"
            End Select
        Next
    End With
    msgDrill(Index).Row = 1

    With imgcboColor(Index) ' �C���[�W�R���{�̏�����
        .ZOrder 0 ' �őO�ʂֈړ�
        .TabStop = False
        .Locked = True ' �ҏW�s��
        ' �C���[�W�R���{�ɍ��ڂ�ǉ�
        With .ComboItems.Add '1�Ԗ�
            .Image = 3
            .Text = "��"
        End With
        With .ComboItems.Add ' 2�Ԗ�
            .Image = 1
            .Text = "��"
        End With
        With .ComboItems.Add ' 3�Ԗ�
            .Image = 2
            .Text = "��"
        End With
        With .ComboItems.Add ' 4�Ԗ�
            .Image = 4
            .Text = "��"
        End With
        With .ComboItems.Add ' 5�Ԗ�
            .Image = 5
            .Text = "Ͼ���"
        End With
        With .ComboItems.Add ' 6�Ԗ�
            .Image = 6
            .Text = "���"
        End With
        With .ComboItems.Add ' 7�Ԗ�
            .Image = 7
            .Text = "�K��"
        End With
        .ComboItems.Item(7).Selected = True
        .Height = msgDrill(Index).CellHeight
        .Width = msgDrill(Index).CellWidth
        .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
              msgDrill(Index).CellTop + msgDrill(Index).top
        .Visible = False
    End With
    msgDrill(Index).Col = 1

    With txtInput(Index) ' �e�L�X�g�{�b�N�X�̏�����
        .ZOrder 0 ' �őO�ʂֈړ�
        .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
              msgDrill(Index).CellTop + msgDrill(Index).top
        .Width = msgDrill(Index).CellWidth
        .Height = msgDrill(Index).CellHeight
        .Appearance = 0 ' �t���b�g
        .Alignment = 0 ' ����
        If msgDrill(Index).Text <> "" Then
            .Text = msgDrill(Index).Text
        Else
            .Text = ""
        End If
        .MaxLength = 5
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    ' ���x���̏�����
    With lblTotal(Index)
        .Alignment = 1 ' �E��
        .Caption = Format(mlngTotal(Index), "##,##0")
    End With
    With lblMinX(Index)
        .Alignment = 1 ' �E��
        .Caption = Format(gudtNCInfo(Index).dblMin(X), "##0.00")
    End With
    With lblMinY(Index)
        .Alignment = 1 ' �E��
        .Caption = Format(gudtNCInfo(Index).dblMin(Y), "##0.00")
    End With
    With lblMaxX(Index)
        .Alignment = 1 ' �E��
        .Caption = Format(gudtNCInfo(Index).dblMax(X), "##0.00")
    End With
    With lblMaxY(Index)
        .Alignment = 1 ' �E��
        .Caption = Format(gudtNCInfo(Index).dblMax(Y), "##0.00")
    End With

End Sub

'*********************************************************
' �p  �r: ���͗p�e�L�X�g�{�b�N�X��KeyPress�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         KeyAscii: ANSI�����R�[�h��\�������l
' �߂�l: ����
'*********************************************************

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)

    ' �r�[�v����ق点��
    If mblnKeyFlag = False Then KeyAscii = 0

End Sub

'*********************************************************
' �p  �r: ���͗p�e�L�X�g�{�b�N�X��KeyUp�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         KeyCode: �L�[ �R�[�h�������萔
'         Shift: Shift, Ctrl, Alt�L�[�̏�Ԃ����������l
' �߂�l: ����
'*********************************************************

Private Sub txtInput_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    mblnKeyFlag = True

End Sub
