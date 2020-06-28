VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ｵﾌﾟｼｮﾝ"
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
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdCansel 
      Caption         =   "ｷｬﾝｾﾙ"
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
      TabCaption(0)   =   "ﾜｰｸﾎﾞｰﾄﾞ"
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
      TabCaption(1)   =   "HPGL変換ﾌﾟﾛｸﾞﾗﾑ"
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
         Text            =   "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
         Top             =   960
         Width           =   1692
      End
      Begin VB.Label lblHPGLCmd 
         Caption         =   "HPGL変換ﾌﾟﾛｸﾞﾗﾑ名"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblSosu 
         Caption         =   "層数"
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
         Caption         =   "ﾋﾟｯﾁ"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblStart 
         Caption         =   "ｽﾀｯｸ"
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
' 用  途: キャンセルボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdCansel_Click()

    Unload Me

End Sub

'*********************************************************
' 用  途: OKボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdOK_Click()

    With gudtWBInfo
        If txtWBS(X).Text <> "" Then .lngWBS(X) = CLng(txtWBS(X)) * int1mm
        If txtWBS(Y).Text <> "" Then .lngWBS(Y) = CLng(txtWBS(Y)) * int1mm
        If txtSosu.Text = "" Then
            ' 何もしない
        ElseIf CInt(txtSosu.Text) > 2 Then ' 多層板
            mlngPitch = CLng(txtPitch.Text) * int1mm
            .lngStack(X) = (.lngWBS(X) - mlngPitch) / 2
            If cmbStack.Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Then
                .lngStack(Y) = CSng(txtWBS(Y).Text) * int1mm / 2
            Else
                .lngStack(Y) = CSng(cmbStack.Text) * int1mm
            End If
            ' 多層板はNC原点とスタック位置が同じ
            .strStart = "STACK"
        Else ' 両面板
            .lngStack(X) = 400&
            If cmbStack.Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Then
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
' 用  途: frmPlotのLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Load()

    Dim strHPGLCommand As String

    txtHPGLCmd.Text = GetSetting("NCView", _
                                 "Settings", _
                                 "HPGLCommand", _
                                 conDefaultHPGLCommand)

    ' プロセス間通信を試みる
    Call sDDElink(gudtWBInfo)

    ' テキストボックスの初期化
    With txtSosu
        .Text = ""
        .MaxLength = 3
        .ToolTipText = "層数"
    End With
    With txtWBS(X)
        .Text = ""
        .MaxLength = 6
        .ToolTipText = "XのWBS"
    End With
    With txtWBS(Y)
        .Text = ""
        .MaxLength = 6
        .ToolTipText = "YのWBS"
    End With
    With txtPitch
        .Text = ""
        .MaxLength = 6
        .ToolTipText = "ザグリ間ピッチ"
    End With

    ' コンボボックスの初期化
    With cmbStack
        .Text = ""
        .ToolTipText = "スタック位置"
    End With

    ' 変数に値がセットされていたらテキストボックスにセットする
    Call sSetTextBox(gudtWBInfo)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' HPGL変換プログラム名をレジストリに保存
    If txtHPGLCmd.Text <> "" Then
        SaveSetting "NCView", _
                    "Settings", _
                    "HPGLCommand", _
                    txtHPGLCmd.Text
    End If

End Sub

'*********************************************************
' 用  途: ピッチ入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtPitch_GotFocus()

    With txtPitch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'*********************************************************
' 用  途: ピッチ入力用テキストボックスのValidateイベント
'         期待通りの値が入力されているかチェックする
' 引  数: Cancel: コントロールがフォーカスを維持するか決定する
'                 Trueで維持
' 戻り値: 無し
'*********************************************************

Private Sub txtPitch_Validate(Cancel As Boolean)

    ' 入力された値をチェック
    With txtPitch
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: 層数入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtSosu_GotFocus()

    With txtSosu
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'*********************************************************
' 用  途: 層数, WBSに応じたスタック位置を決定する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub sSetStack()

    Dim lngWBSY As Long

    ' 変数に層数をセット
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
                .Text = "180" ' AMS品は180スタック
            ElseIf lngWBSY > 500 Then
                .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
            ElseIf lngWBSY >= 400 Then
                .Text = "205"
            ElseIf lngWBSY <> 0 Then
                .Text = "180"
            End If
        Else
            .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
        End If
    End With

End Sub

'*********************************************************
' 用  途: 層数入力用テキストボックスのLostFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtSosu_LostFocus()

    ' ピッチ/スタックを位置の設定
    Call sSetStack

End Sub

'*********************************************************
' 用  途: 層数入力用テキストボックスのValidateイベント
'         期待通りの値が入力されているかチェック
' 引  数: Cancel: コントロールがフォーカスを維持するか決定する
'                 Trueで維持
' 戻り値: 無し
'*********************************************************

Private Sub txtSosu_Validate(Cancel As Boolean)

    ' 入力された値をチェック
    With txtSosu
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: WBS入力用テキストボックスのGotFocusイベント
' 引  数: Index: XYを示す0または1いずれかの数値
' 戻り値: 無し
'*********************************************************

Private Sub txtWBS_GotFocus(Index As Integer)

    With txtWBS(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'*********************************************************
' 用  途: WBS入力用テキストボックスのLostFocusイベント
' 引  数: Index: XYを示す0または1いずれかの数値
' 戻り値: 無し
'*********************************************************

Private Sub txtWBS_LostFocus(Index As Integer)

    ' ピッチ/スタックを位置の設定
    Call sSetStack

End Sub

'*********************************************************
' 用  途: WBS入力用テキストボックスのValidateイベント
'         期待通りの値が入力されているかチェック
' 引  数: Index: XYを示す0または1いずれかの数値
'         Cancel: コントロールがフォーカスを維持するか決定する
'                 Trueで維持
' 戻り値: 無し
'*********************************************************

Private Sub txtWBS_Validate(Index As Integer, Cancel As Boolean)

    ' 入力された値をチェック
    With txtWBS(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: テキストボックスに層数, WBS etc...の値を代入する
' 引  数: udtWBInfo: WB情報が格納された構造体WBInfo
' 戻り値: 無し
'*********************************************************

Private Sub sSetTextBox(udtWBInfo As WBInfo)

    ' 変数に値がセットされていたらテキストボックスにセットする
    With udtWBInfo
        If .intSosu <> 0 Then txtSosu.Text = .intSosu
        If .lngWBS(X) <> 0 Then txtWBS(X).Text = .lngWBS(X) / 100
        If .lngWBS(Y) <> 0 Then txtWBS(Y).Text = .lngWBS(Y) / 100
        If .lngStack(Y) <> 0 Then cmbStack.Text = .lngStack(Y) / 100
        If mlngPitch <> 0 Then txtPitch.Text = mlngPitch / 100
    End With

    Call sSetStack

End Sub
