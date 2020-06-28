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
      Align           =   1  '上揃え
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
            Key             =   "開く"
            Object.ToolTipText     =   "開く"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "印刷"
            Object.ToolTipText     =   "印刷"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "閲覧"
            Object.ToolTipText     =   "閲覧"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDDE 
      BorderStyle     =   1  '実線
      Caption         =   "DDE用ラベル"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "開く(&O)"
      End
      Begin VB.Menu mnuNTIn 
         Caption         =   "NTの読込み(&I)"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "印刷(&P)"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "NCViewの終了(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "表示(&V)"
      Begin VB.Menu mnuLook 
         Caption         =   "閲覧(&G)"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeisun 
         Caption         =   "正寸表示(&L)"
      End
      Begin VB.Menu mnuStandard 
         Caption         =   "全体表示(&S)"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "ﾂｰﾙ(&T)"
      Begin VB.Menu mnuInfo 
         Caption         =   "工具情報(&I)"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "ｵﾌﾟｼｮﾝ(&O)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ情報(&A)..."
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

' ウィンドウの矩形サイズを取得
Private Declare Function GetWindowRect Lib "user32.dll" _
    (ByVal hwnd As Long, _
     lpRect As tagRECT) As Long

' マウスカーソルの移動範囲を指定する関数
Private Declare Function ClipCursor Lib "user32.dll" _
    (lpRect As Any) As Long

' システムの設定やシステムメトリックの値を取得する関数
Private Declare Function GetSystemMetrics Lib "user32.dll" _
    (ByVal nIntex As Long) As Long

Private Const SM_CYCAPTION = 4 ' タイトルバーの高さを取得
Private Const SM_CYMENU = 15 ' クライアントウィンドウのメニューの高さを取得

Private mblnDisp As Boolean ' 画面表示後か否かを示すFlag
Private mblnPanMode As Boolean ' picDrawの移動可/不可を示すFlag
Private msngDragDistX As Single ' MouseDown時のマウスポインターのX座標
Private msngDragDistY As Single ' MouseDown時のマウスポインターのX座標
Private msngCurrentTop As Single ' MouseMove時のpicDrawのTopプロパティ
Private msngCurrentLeft As Single ' MouseMove時のpicDrawのLeftプロパティ
Private mblnHMove As Boolean ' picDrawが横方向へ移動可/不可を示すFlag
Private mblnVMove As Boolean ' picDrawが縦方向へ移動可/不可を示すFlag
Private mstrFileName As String ' 現在編集中のNCファイル名(プロパティ用)
Private mintTHNT As Integer ' 現在編集中のNCはTH/NTかを示す(プロパティ用)
Private mdblScaleFactor As Double ' ピクチャボックスに表示する時のファクタ(プロパティ用)

'*********************************************************
' 用  途: 現在作業中のNCファイル名(NCFileNameプロパティ)の取得
' 引  数: 無し
' 戻り値: NCFileNameプロパティの値
'*********************************************************

Public Property Get NCFileName() As String

    NCFileName = mstrFileName

End Property

'*********************************************************
' 用  途: NCFileNameプロパティに現在作業中のNCファイル名をセット
' 引  数: strFileName: 現在作業中のNCファイル名
' 戻り値: 無し
'*********************************************************

Public Property Let NCFileName(ByVal strFileName As String)

    mstrFileName = strFileName

End Property

'*********************************************************
' 用  途: 現在作業中のNCがTH/NTかを示す値(THNTプロパティ)の取得
' 引  数: 無し
' 戻り値: THNTプロパティの値
'*********************************************************

Public Property Get THNT() As Integer

    THNT = mintTHNT

End Property

'*********************************************************
' 用  途: THNTプロパティに現在作業中のNCがTH/NTかを示す値をセット
' 引  数: intTHNT: TH/NTを示す値(0 or 1)
' 戻り値: 無し
'*********************************************************

Public Property Let THNT(ByVal intTHNT As Integer)

    mintTHNT = intTHNT

End Property

'*********************************************************
' 用  途: ピクチャボックスに表示する為のスケールファクタ
'         (ScaleFactorプロパティ)の取得
' 引  数: 無し
' 戻り値: ScaleFactorプロパティの値
'*********************************************************

Public Property Get ScaleFactor() As Double

    ScaleFactor = mdblScaleFactor

End Property

'*********************************************************
' 用  途: ScaleFactorプロパティにピクチャボックスに表示する
'         為のスケールファクタをセット
' 引  数: dblScaleFactor: スケールファクタ
' 戻り値: 無し
'*********************************************************

Public Property Let ScaleFactor(ByVal dblScaleFactor As Double)

    mdblScaleFactor = dblScaleFactor

End Property

'*********************************************************
' 用  途: frmMainのKeyDownイベント
' 引  数: KeyCode: キー コードを示す定数
'         Shift: イベント発生時のShift, Ctrl, Altキーの
'                状態を示す整数値
' 戻り値: 無し
'*********************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then gblnCancel = False

End Sub

'*********************************************************
' 用  途: frmMainのUnloadイベント
' 引  数: Cancel: フォームを画面から消去するかどうかを指定する
'                 整数値(0で消去, その他は消去しない)
' 戻り値: 無し
'*********************************************************

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Unload frmToolInfo

    If Dir(fTempPath & conTempFileName) <> "" Then
        Kill fTempPath & conTempFileName ' テンポラリファイルを削除
    End If

    If WindowState = vbNormal Then
        ' Formの位置と大きさをレジストリに保存
        SaveSetting "NCView", "Position", "Top", top
        SaveSetting "NCView", "Position", "Left", left
        SaveSetting "NCView", "Position", "Height", Height
        SaveSetting "NCView", "Position", "Width", Width
    End If

    ' 最小化の時はWindowStateを保存しない
    If WindowState <> vbMinimized Then
        SaveSetting "NCView", "Position", "WindowState", WindowState
    End If

End Sub

'*********************************************************
' 用  途: frmMain.mnuAboutのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuAbout_Click()

    Load frmAbout
    frmAbout.Show vbModal

End Sub

'*********************************************************
' 用  途: frmMain.mnuInfoのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuInfo_Click()

    Load frmToolInfo
    frmToolInfo.Show vbModal

End Sub

'*********************************************************
' 用  途: frmMain.mnuLookのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuLook_Click()

    Dim strNC As String
    Dim blnRet As Boolean
    Dim i As Integer

    gblnCancel = True

    ' NCを読み込む
    strNC = fReadNC(Me.NCFileName)

    ' コントロールの初期値の設定
    Call sSetControl

    If Me.THNT = TH Then
        picDraw(0).Cls ' 全体表示用ピクチャーボックス
        picDraw(1).Cls ' 正寸表示用ピクチャーボックス
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
' 用  途: frmMain.mnuNTInのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuNTIn_Click()

    Dim strNCFileName As String
    Dim strInput As String
    Dim strRet As String
    Dim strNC As String
    Dim strXY() As String
    Dim strStack As String

    On Error GoTo Trap

    Me.THNT = NT ' NTのTabをアクティブにする為にここで設定しておく
    
    ' 変数がセットされていなければプロセス間通信を試みる
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
        MsgBox "THを設定しないとNTは合成出来ません。"
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
            strRet = InputBox("層数は?", "層数の入力")
            If strRet = "" Then Exit Sub
            .intSosu = CInt(strRet)
            If .intSosu > 2 Then
'                .lngStack(X) = 5& * int1mm
                strNC = fReadNC(strNCFileName) ' NTの読み込み
                strRet = fGetNTIdou(strNC)
                If strRet <> "" Then
                    strXY = Split(Mid(strRet, 2), "Y", -1, vbTextCompare)
                    strStack = CLng(strXY(Y)) / int1mm
                End If
            Else ' 両面板の場合
'                .lngStack(X) = 4& * int1mm
                With gudtNCInfo(TH)
                    strNC = fReadNC(.strFileName) ' THの読み込み
                    strRet = fGetNTIdou(strNC)
                    If strRet <> "" Then
                        strXY = Split(Mid(strRet, 2), "Y", -1, vbTextCompare)
                        strStack = 180 - (CLng(strXY(Y)) / int1mm)
                    Else
                        strStack = "180"
                    End If
                End With
            End If
            strInput = InputBox("何mmスタックですか?", "スタックの入力", strStack)
            .lngStack(Y) = CLng(CSng(strInput) * int1mm)
        End If
    End With

    Exit Sub

Trap:
    MsgBox "入力エラーです"

End Sub

'*********************************************************
' 用  途: frmMain.mnuOpenのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuOpen_Click()

    Dim strNCFileName As String
    Dim intRet As Integer ' MsgBoxの戻り値

    intRet = MsgBox("現在の工具情報は破棄されます。", 49, "確認")
    If intRet = vbCancel Then Exit Sub

    ' 変数を初期化する
    Call sInitialize

    ' frmToolInfoのプロパティを初期化する
    With frmToolInfo
        .THLoadFlag = False
        .NTLoadFlag = False
    End With

    Me.THNT = TH ' THのTabをアクティブにする為にここで設定しておく
    strNCFileName = fGetInputFile()
    If strNCFileName = "" Then Exit Sub

    Me.NCFileName = strNCFileName
    Caption = conCaption & " - " & Me.NCFileName

    Load frmToolInfo
    frmToolInfo.Show vbModal

End Sub

'*********************************************************
' 用  途: frmMain.mnuOptionのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuOption_Click()

    Load frmOption
    frmOption.Show vbModal

End Sub

'*********************************************************
' 用  途: frmMain.mnuPrintのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuPrint_Click()

    Dim i As Integer

    ' コントロールの初期値の設定
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
' 用  途: frmMain.mnuQuitのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuQuit_Click()

    ' NCViewの終了
    Unload Me
    End

End Sub

'*********************************************************
' 用  途: frmMain.mnuSeisunのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuSeisun_Click()

    picDraw(0).Visible = False ' 全体表示
    picDraw(1).Visible = True ' 正寸表示

End Sub

'*********************************************************
' 用  途: frmMain.mnuStandardのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuStandard_Click()

    picDraw(0).Visible = True ' 全体表示
    picDraw(1).Visible = False ' 正寸表示

End Sub

'*********************************************************
' 用  途: frmMain.picDraw()のKeyDownイベント
' 引  数: Index: コントロール配列のIndex
'         KeyCode: キー コードを示す定数
'         Shift: イベント発生時のShift, Ctrl, Altキーの
'                状態を示す整数値
' 戻り値: 無し
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

    ' ピクチャーボックスがコンテナからはみ出さない様にする
    Call sPicPosition(Index)

End Sub

'*********************************************************
' 用  途: frmMain.Toolbar1のButtonClickイベント
' 引  数: Button: クリックされた Button オブジェクトへの参照
' 戻り値: 無し
'*********************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

    Select Case Button.Key
        Case "開く"
            Call mnuOpen_Click
        Case "印刷"
            Call mnuPrint_Click
        Case "閲覧"
            Call mnuLook_Click
    End Select

End Sub

'*********************************************************
' 用  途: frmMainのLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Load()

    ' 前回終了時のFormの位置と大きさを復元
    top = GetSetting("NCView", "Position", "Top", "0")
    left = GetSetting("NCView", "Position", "Left", "0")
    Height = GetSetting("NCView", "Position", "Height", Height)
    Width = GetSetting("NCView", "Position", "Width", Width)
    WindowState = GetSetting("NCView", "Position", "WindowState", vbNormal)

    KeyPreview = True
    gblnCancel = True

    ' プログレスバーを非表示にする
    ProgressBar1.Visible = False

    ' タイトル
    Caption = conCaption

    ' DDE通信用ラベルコントロールを非表示にする
    lblDDE.Visible = False

    ' フォームの初期化
    Call sInit

    ' Loadイベント中はToolBar1.Heightが"555"を返すので"360"で初期化。
    ' こんなのでいいのか...
    picDraw(0).Height = SysInfo1.WorkAreaHeight _
                        - (GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY) _
                        - 360 _
                        - (GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY)

    ' ピクチャーボックスの表示/非表示の設定
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
' 用  途: frmMainのResizeイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Resize()

    With picFrame
        .Align = 3 ' 左揃え
        .Align = 1 ' 上揃え
    End With

End Sub

'*********************************************************
' 用  途: frmMain.picDrawのMouseDownイベント
' 引  数: Index: コントロール配列のIndex
'         Button: 押されたボタンを示す製数値
'         Shift: ボタンが押された時のShift, Ctrl, Altキーの
'                状態を示す製数値
'         X, Y: マウスポインタの現在位置を表す数値
' 戻り値: 無し
'*********************************************************

Private Sub picDraw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim udtRect As tagRECT

    If Button = 1 Then
        MousePointer = vbCustom 'マウスカーソルを変更
        mblnPanMode = True
        mblnHMove = True
        mblnVMove = True
        msngDragDistX = X
        msngDragDistY = Y
        msngCurrentTop = picDraw(Index).top
        msngCurrentLeft = picDraw(Index).left

        ' ピクチャボックスの矩形領域を取得
        GetWindowRect picFrame.hwnd, udtRect
        ' 取得した領域にマウスの移動範囲を制限
        ClipCursor udtRect
    ElseIf Button = 2 Then
        PopupMenu mnuView
    End If

End Sub

'*********************************************************
' 用  途: frmMain.picDrawのMouseMoveイベント
' 引  数: Index: コントロール配列のIndex
'         Button: 押されたボタンを示す製数値
'         Shift: Shift, Ctrl, Altキーの状態を示す製数値
'         X, Y: マウスポインタの現在位置を表す数値
' 戻り値: 無し
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

    ' ピクチャーボックスがコンテナからはみ出さない様にする
    Call sPicPosition(Index)

    ' left, topプロパティはtwip単位である事に注意!
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
' 用  途: frmMain.picDrawのMouseUpイベント
' 引  数: Index: コントロール配列のIndex
'         Button: 離されたボタンを示す製数値
'         Shift: 離された時のShift, Ctrl, Altキーの状態を
'                示す製数値
'         X, Y: マウスポインタの現在位置を表す数値
' 戻り値: 無し
'*********************************************************

Private Sub picDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mblnPanMode = False
    mblnHMove = False
    mblnVMove = False

    If Button = 1 Then ' 左ボタン
        MousePointer = vbDefault ' マウスカーソルをデフォルトに戻す

        ' 引数にNULLを指定することで
        ' マウスカーソルの移動制限を解除
        ClipCursor ByVal 0
    End If

End Sub

'*********************************************************
' 用  途: ピクチャボックスの初期化
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub sInit()

    mblnDisp = False

    With picFrame
        .Align = 3 ' 左揃え
        .Align = 1 ' 上揃え
    End With

    With picDraw(0) ' 全体表示用ピクチャーボックス
        ' 背景が白
        .BackColor = RGB(250, 250, 250)
        .ForeColor = QBColor(0) ' 黒
        picDraw(0).Width = SysInfo1.WorkAreaWidth
        picDraw(0).Height = SysInfo1.WorkAreaHeight _
                            - (GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY) _
                            - Toolbar1.Height _
                            - (GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY)
        .top = -24
        .left = -24
        .AutoRedraw = True
        .ScaleHeight = -Abs(.ScaleHeight)
        .Appearance = 0 ' フラット
'        .Visible = True
    End With


    With picDraw(1) ' 正寸表示用ピクチャボックス
        ' 背景が白
        .BackColor = RGB(250, 250, 250)
        .ForeColor = QBColor(0) ' 黒
        .top = -24
        .left = -24
        .AutoRedraw = True
        .Width = picFrame.Width
        .Height = picFrame.Height
        .ScaleHeight = -Abs(.ScaleHeight)
        .Appearance = 0 ' フラット
'        .Visible = False ' 起動時は非表示
    End With

End Sub

'*********************************************************
' 用  途: プログレスバーの初期化
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub sSetControl()

    ' プログレスバーのプロパティの設定
    With ProgressBar1
        .Width = 3000
        .Height = Toolbar1.Height - 36
        .top = 36
        .left = picFrame.Width - .Width
    End With

    ' プロセス間通信を試みる
    Call sDDElink(gudtWBInfo)

End Sub

Private Sub sPicPosition(Index As Integer)

'*********************************************************
' 用  途: ピクチャーボックスがコンテナからはみ出さない様にする
' 引  数: 無し
' 戻り値: 無し
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
            If .left < -24 Then ' 左側
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
' 用  途: ファイルを開くダイアログを表示する
' 引  数: 無し
' 戻り値: 選択したファイル名
'*********************************************************

Public Function fGetInputFile() As String

    With CommonDialog1
        ' CancelErrorプロパティを真(True)に設定します。
        .CancelError = True
        On Error GoTo ErrHandler

        ' ファイルの選択方法を設定します。
        .Filter = "すべてのファイル (*.*)|*.*|" & _
                  "データファイル (*.dat)|*.dat|" & _
                  "NCデータファイル (*.nc)|*.nc"

        ' 既定の選択方法を指定します。
        .FilterIndex = 1

        ' [読み取り専用ファイルとして開く]チェックボックスを表示しない
        ' 既存のファイル名しか入力できないようにする
        .Flags = cdlOFNHideReadOnly Or _
                 cdlOFNFileMustExist

        ' [ファイルを開く] ダイアログ ボックスを表示します。
        .ShowOpen

        fGetInputFile = .FileName
        Exit Function
    End With

ErrHandler:
        'ユーザーが[キャンセル] ボタンをクリックしました。
'       If Err.Number = cdlCancel Then
'           If Mid(Caption, Len(conCaption) + 4) = "" Then
'               Caption = conCaption
'           End If
'       End If
    fGetInputFile = ""

End Function
