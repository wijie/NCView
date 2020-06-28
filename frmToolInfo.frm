VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmToolInfo 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "工具情報"
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
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
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
         Text            =   "適当"
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
         Text            =   "適当"
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
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71040
         TabIndex        =   22
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lblMaxX 
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71880
         TabIndex        =   21
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lblMinY 
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71040
         TabIndex        =   19
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label lblMinX 
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   1
         Left            =   -71880
         TabIndex        =   18
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  '実線
         Caption         =   "999,999"
         Height          =   252
         Index           =   1
         Left            =   -71880
         TabIndex        =   16
         Top             =   720
         Width           =   732
      End
      Begin VB.Label Label15 
         Caption         =   "最小値"
         Height          =   252
         Left            =   -71880
         TabIndex        =   17
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label14 
         Caption         =   "穴数合計"
         Height          =   252
         Left            =   -71880
         TabIndex        =   15
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label12 
         Caption         =   "最大値"
         Height          =   252
         Left            =   -71880
         TabIndex        =   20
         Top             =   1920
         Width           =   612
      End
      Begin VB.Label lblMaxX 
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3120
         TabIndex        =   10
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  '実線
         Caption         =   "999,999"
         Height          =   252
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   732
      End
      Begin VB.Label lblMinX 
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "最大値"
         Height          =   252
         Left            =   3120
         TabIndex        =   9
         Top             =   1920
         Width           =   612
      End
      Begin VB.Label lblMinY 
         BorderStyle     =   1  '実線
         Caption         =   "-999.99"
         Height          =   252
         Index           =   0
         Left            =   3960
         TabIndex        =   8
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label Label7 
         Caption         =   "穴数合計"
         Height          =   252
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "最小値"
         Height          =   252
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label lblMaxY 
         BorderStyle     =   1  '実線
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
      Caption         =   "ｷｬﾝｾﾙ"
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

Private mblnTHLoadFlag As Boolean ' 初めてロードするのか否かを示すフラグ(TH)
Private mblnNTLoadFlag As Boolean ' 初めてロードするのか否かを示すフラグ(NT)
Private mlngTotal(1) As Long
Private mudtToolInfo(1, 1 To intRow) As New frmToolInfo ' TH/NTのツール情報
Private mblnKeyFlag As Boolean ' テキストボックスでビープ音が鳴らない様にするフラグ

'*********************************************************
' 用  途: THのNCをロード済みか否かを示す値
'         (THLoadFlagプロパティ)の取得
' 引  数: 無し
' 戻り値: プロパティの値
'*********************************************************

Public Property Get THLoadFlag() As Boolean

    THLoadFlag = mblnTHLoadFlag

End Property

'*********************************************************
' 用  途: THLoadFlagプロパティにロード済みを示す値をセット
' 引  数: blnTHLoadFlag: ロード済み-True, 未-False
' 戻り値: 無し
'*********************************************************

Public Property Let THLoadFlag(ByVal blnTHLoadFlag As Boolean)

    mblnTHLoadFlag = blnTHLoadFlag

End Property

'*********************************************************
' 用  途: NTのNCをロード済みか否かを示す値
'         (NTLoadFlagプロパティ)の取得
' 引  数: 無し
' 戻り値: プロパティの値
'*********************************************************

Public Property Get NTLoadFlag() As Boolean

    NTLoadFlag = mblnNTLoadFlag

End Property

'*********************************************************
' 用  途: NTLoadFlagプロパティにロード済みを示す値をセット
' 引  数: blnNTLoadFlag: ロード済み-True, 未-False
' 戻り値: 無し
'*********************************************************

Public Property Let NTLoadFlag(ByVal blnNTLoadFlag As Boolean)

    mblnNTLoadFlag = blnNTLoadFlag

End Property

'*********************************************************
' 用  途: OKボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdOK_Click()

    Dim strColor As String
    Dim i As Integer
    Dim intTHNT As Integer
    Dim intColor As Integer
    Dim varColorList() As Variant

    Me.Hide ' フォームを非表示にする

    ' 適当な時は,このリストから順番に選ぶ
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
                .Col = 1 ' ドリル径
                If .Text <> "" Then
                    .Col = 3 ' 色
                    Select Case .Text
                        Case "□適当"
                            strColor = varColorList(intColor)
                            intColor = intColor + 1
                            If intColor > UBound(varColorList) Then
                                intColor = 0
                            End If
                        Case "■黒"
                            strColor = "BLACK"
                        Case "■赤"
                            strColor = "RED"
                        Case "■緑"
                            strColor = "GREEN"
                        Case "■青"
                            strColor = "BLUE"
                        Case "■ﾏｾﾞﾝﾀ"
                            strColor = "MAGENTA"
                        Case "■ｼｱﾝ"
                            strColor = "CYAN"
                    End Select
                    .Col = 0 ' TNo
                    gudtToolInfo(intTHNT, i).intTNo = CInt(.Text)
                    .Col = 1 ' ドリル径
                    gudtToolInfo(intTHNT, i).sngDrill = CSng(.Text) / 2
                    .Col = 3 ' 色
                    gudtToolInfo(intTHNT, i).strColor = strColor
                End If
            Next
        End With
    Next

    Unload Me

End Sub

'*********************************************************
' 用  途: キャンセルボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdCansel_Click()

    Unload Me

End Sub

'*********************************************************
' 用  途: frmToolInfoのLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Load()

    Dim strNC As String

    mblnKeyFlag = True
    mlngTotal(TH) = 0
    mlngTotal(NT) = 0

    Call sInit(NT) ' フォームの初期化
    Call sInit(TH) ' フォームの初期化

    If frmMain.THNT = TH Then
        SSTab1.Tab = TH
        If Me.THLoadFlag = False Then
            strNC = fReadNC(frmMain.NCFileName) ' 読むのは最初の1回だけ
            Call sSetUsedTool(strNC) ' 使われているツールを調べる
        End If
    Else
        SSTab1.Tab = NT
        If Me.NTLoadFlag = False Then
            strNC = fReadNC(frmMain.NCFileName) ' 読むのは最初の1回だけ
            Call sSetUsedTool(strNC) ' 使われているツールを調べる
        End If
    End If
    SSTab1.TabStop = False

End Sub

'*********************************************************
' 用  途: 色選択用イメージコンボのClickイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub imgcboColor_Click(Index As Integer)

    With msgDrill(Index)
        Select Case imgcboColor(Index).SelectedItem.Text
            Case "緑"
                .CellForeColor = RGB(0, 220, 0)
                .Text = "■緑"
            Case "黒"
                .CellForeColor = RGB(0, 0, 0)
                .Text = "■黒"
            Case "赤"
                .CellForeColor = RGB(255, 0, 0)
                .Text = "■赤"
            Case "青"
                .CellForeColor = RGB(0, 0, 255)
                .Text = "■青"
            Case "ﾏｾﾞﾝﾀ"
                .CellForeColor = RGB(255, 0, 255)
                .Text = "■ﾏｾﾞﾝﾀ"
            Case "ｼｱﾝ"
                .CellForeColor = RGB(0, 220, 220)
                .Text = "■ｼｱﾝ"
            Case Else
                .CellForeColor = RGB(0, 0, 0)
                .Text = "□適当"
        End Select
    End With

End Sub

'*********************************************************
' 用  途: 色選択用イメージコンボのGotFocusイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub imgcboColor_GotFocus(Index As Integer)

'    SendKeys "{F4}"

End Sub

'*********************************************************
' 用  途: フレキシブルグリッドのClickイベント
' 引  数: コントロール配列のIndexプロパティ
' 戻り値: 無し
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
        Case 3 ' 色の桁
            imgcboColor(Index).Visible = True
            With msgDrill(Index)
                txtInput(Index).Visible = False
                imgcboColor(Index).top = .CellTop + .top
                imgcboColor(Index).SetFocus
            End With
            With imgcboColor(Index).ComboItems
                Select Case msgDrill(Index).Text
                    Case "■緑"
                        .Item(1).Selected = True
                    Case "■黒"
                        .Item(2).Selected = True
                    Case "■赤"
                        .Item(3).Selected = True
                    Case "■青"
                        .Item(4).Selected = True
                    Case "■ﾏｾﾞﾝﾀ"
                        .Item(5).Selected = True
                    Case "■ｼｱﾝ"
                        .Item(6).Selected = True
                    Case Else
                        .Item(7).Selected = True
                End Select
            End With
        Case 2 ' 穴数の桁
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
' 用  途: フレキシブルグリッドのScrollイベント
' 引  数: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub msgDrill_Scroll(Index As Integer)

    ' コントロールが表示される前にイベントが発生するとエラーになるのでトラップする(-_-;
    On Error GoTo bye

    msgDrill(Index).SetFocus ' TextBoxにFocusがある時にScrollするとFocusがコマンドボタンに飛んでしまう為
    txtInput(Index).Visible = False
    imgcboColor(Index).Visible = False

bye:

End Sub

'*********************************************************
' 用  途: タブコントロールのClickイベント
' 引  数: PreviousTab: 切り替え前のタブのIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub SSTab1_Click(PreviousTab As Integer)

'    txtInput(SSTab1.Tab).SetFocus

End Sub

'*********************************************************
' 用  途: 入力用テキストボックスのChangeイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtInput_Change(Index As Integer)

    With msgDrill(Index)
        .CellAlignment = 1
        .Text = txtInput(Index).Text
    End With

End Sub

'*********************************************************
' 用  途: 入力用テキストボックスのKeyDownイベント
' 引  数: Index: コントロール配列のIndexプロパティ
'         KeyCode: キー コードを示す定数
'         Shift: イベント発生時のShift, Ctrl, Altキーの
'                状態を示す整数値
' 戻り値: 無し
'*********************************************************

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    ' 何も入力されていない時は次に進まない(戻るのは許可する)
    If txtInput(Index).Text = "" And KeyCode <> vbKeyUp Then Exit Sub

    ' Enter又は, Ctrl-M又は, 下矢印キー
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
' 用  途: 使用されているTコードを調べてグリッドコントロールに
'         セットする
' 引  数: strNC: NCデータ
' 戻り値: 無し
'*********************************************************

Private Sub sSetUsedTool(ByVal strNC As String)

    Dim i As Integer
    Dim objReg As New RegExp
    Dim objMatches As Object
    Dim objMatch As Object

    objReg.Global = True
    objReg.IgnoreCase = False ' 大文字小文字を区別する
    objReg.Pattern = "T[0-9]+"
    Set objMatches = objReg.Execute(strNC)

    ' Tコードを工具情報にセットする
    i = 1
    With Me.msgDrill(frmMain.THNT)
        For Each objMatch In objMatches
            .Row = i
            .Col = 0
            .Text = Mid(objMatch.Value, 2)
            ' T50のデフォルトを設定する
            If objMatch = "T50" Then
                .Col = 1 ' ドリル径
                With frmToolInfo.txtInput(frmMain.THNT)
                    .Text = "1.999"
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
                .Col = 3 ' 色の桁
                .CellForeColor = RGB(0, 0, 0)
                .Text = "■黒"
            End If
            i = i + 1
        Next
        ' デフォルトの位置にセット
        .Row = 1
        .Col = 1
    End With

End Sub

'*********************************************************
' 用  途: コントロールの初期化
' 引  数: Index: TH/NTを示す値(TH = 0, NT = 1)
' 戻り値: 無し
'*********************************************************

Private Sub sInit(Index As Integer)

    Dim i As Integer

    With msgDrill(Index) ' グリッドの初期化
        .Cols = 4
        .Rows = intRow + 1 ' +1しているのは固定行がある為
        .FixedCols = 0 ' 固定列なし
        .FixedRows = 1 ' 固定行1
        .Width = 2880
        .Height = 4560
        .RowHeight(-1) = 288 ' 全列の高さ
        .RowHeight(0) = 240 ' 固定列の高さ
        .ColWidth(0) = 456 ' TNo.の桁幅
        .ColWidth(1) = 624 ' ドリル径の桁幅
        .ColWidth(2) = 672 ' 穴数の桁幅
        .ColWidth(3) = 780 ' 色の桁幅
'        .FillStyle = flexFillRepeat
        .FocusRect = flexFocusNone ' フォーカスを示す線を表示しない
        .HighLight = flexHighlightNever ' 強調表示しない
        .Row = 0 ' 個定列
        .Col = 0
        .Text = "TNo."
        .Col = 1
        .Text = "ﾄﾞﾘﾙ径"
        .Col = 2
        .Text = "穴数"
        .Col = 3
        .Text = "色"
        For i = 1 To intRow
            .Row = i
            .Col = 0 ' TNoの桁
            .CellAlignment = 1 ' 左側の中央
            If gudtToolInfo(Index, i).intTNo > -1 Then
                .Text = gudtToolInfo(Index, i).intTNo
            End If
            .Col = 1 ' ドリル径の桁
            .CellAlignment = 1 ' 左側の中央
            If gudtToolInfo(Index, i).sngDrill > -1 Then
                .Text = Format(gudtToolInfo(Index, i).sngDrill * 2, "##0.000")
            End If
            .Col = 2 ' 穴数の桁
            If gudtToolInfo(Index, i).intTNo = 50 Then ' 逆セットチェック用ツール
                ' 穴数は表示だけで合計穴数には数えない
                .Text = "(" & gudtToolInfo(Index, i).lngCount & ")"
            ElseIf gudtToolInfo(Index, i).lngCount > -1 Then
                .Text = Format(gudtToolInfo(Index, i).lngCount, "##,##0")
                mlngTotal(Index) = mlngTotal(Index) + gudtToolInfo(Index, i).lngCount
            End If
            .Col = 3 ' 色の桁
            Select Case gudtToolInfo(Index, i).strColor
                Case "GREEN" ' イメージコンボの1番目
                    .CellForeColor = RGB(0, 220, 0)
                    .Text = "■緑"
                Case "BLACK" ' 2番目
                    .CellForeColor = RGB(0, 0, 0)
                    .Text = "■黒"
                Case "RED" ' 3番目
                    .CellForeColor = RGB(255, 0, 0)
                    .Text = "■赤"
                Case "BLUE" ' 4番目
                    .CellForeColor = RGB(0, 0, 255)
                    .Text = "■青"
                Case "MAGENTA" ' 5番目
                    .CellForeColor = RGB(255, 0, 255)
                    .Text = "■ﾏｾﾞﾝﾀ"
                Case "CYAN" ' 6番目
                    .CellForeColor = RGB(0, 220, 220)
                    .Text = "■ｼｱﾝ"
                Case Else
                    .CellForeColor = RGB(0, 0, 0)
                    .Text = "□適当"
            End Select
        Next
    End With
    msgDrill(Index).Row = 1

    With imgcboColor(Index) ' イメージコンボの初期化
        .ZOrder 0 ' 最前面へ移動
        .TabStop = False
        .Locked = True ' 編集不可
        ' イメージコンボに項目を追加
        With .ComboItems.Add '1番目
            .Image = 3
            .Text = "緑"
        End With
        With .ComboItems.Add ' 2番目
            .Image = 1
            .Text = "黒"
        End With
        With .ComboItems.Add ' 3番目
            .Image = 2
            .Text = "赤"
        End With
        With .ComboItems.Add ' 4番目
            .Image = 4
            .Text = "青"
        End With
        With .ComboItems.Add ' 5番目
            .Image = 5
            .Text = "ﾏｾﾞﾝﾀ"
        End With
        With .ComboItems.Add ' 6番目
            .Image = 6
            .Text = "ｼｱﾝ"
        End With
        With .ComboItems.Add ' 7番目
            .Image = 7
            .Text = "適当"
        End With
        .ComboItems.Item(7).Selected = True
        .Height = msgDrill(Index).CellHeight
        .Width = msgDrill(Index).CellWidth
        .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
              msgDrill(Index).CellTop + msgDrill(Index).top
        .Visible = False
    End With
    msgDrill(Index).Col = 1

    With txtInput(Index) ' テキストボックスの初期化
        .ZOrder 0 ' 最前面へ移動
        .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
              msgDrill(Index).CellTop + msgDrill(Index).top
        .Width = msgDrill(Index).CellWidth
        .Height = msgDrill(Index).CellHeight
        .Appearance = 0 ' フラット
        .Alignment = 0 ' 左寄せ
        If msgDrill(Index).Text <> "" Then
            .Text = msgDrill(Index).Text
        Else
            .Text = ""
        End If
        .MaxLength = 5
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    ' ラベルの初期化
    With lblTotal(Index)
        .Alignment = 1 ' 右寄せ
        .Caption = Format(mlngTotal(Index), "##,##0")
    End With
    With lblMinX(Index)
        .Alignment = 1 ' 右寄せ
        .Caption = Format(gudtNCInfo(Index).dblMin(X), "##0.00")
    End With
    With lblMinY(Index)
        .Alignment = 1 ' 右寄せ
        .Caption = Format(gudtNCInfo(Index).dblMin(Y), "##0.00")
    End With
    With lblMaxX(Index)
        .Alignment = 1 ' 右寄せ
        .Caption = Format(gudtNCInfo(Index).dblMax(X), "##0.00")
    End With
    With lblMaxY(Index)
        .Alignment = 1 ' 右寄せ
        .Caption = Format(gudtNCInfo(Index).dblMax(Y), "##0.00")
    End With

End Sub

'*********************************************************
' 用  途: 入力用テキストボックスのKeyPressイベント
' 引  数: Index: コントロール配列のIndex
'         KeyAscii: ANSI文字コードを表す整数値
' 戻り値: 無し
'*********************************************************

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)

    ' ビープ音を黙らせる
    If mblnKeyFlag = False Then KeyAscii = 0

End Sub

'*********************************************************
' 用  途: 入力用テキストボックスのKeyUpイベント
' 引  数: Index: コントロール配列のIndex
'         KeyCode: キー コードを示す定数
'         Shift: Shift, Ctrl, Altキーの状態を示す整数値
' 戻り値: 無し
'*********************************************************

Private Sub txtInput_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    mblnKeyFlag = True

End Sub
