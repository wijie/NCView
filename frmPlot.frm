VERSION 5.00
Begin VB.Form frmPlot 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "印刷"
   ClientHeight    =   1455
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmPlot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdCansel 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox cmbStack 
      Height          =   276
      ItemData        =   "frmPlot.frx":000C
      Left            =   1920
      List            =   "frmPlot.frx":0019
      TabIndex        =   7
      Text            =   "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
      Top             =   480
      Width           =   1692
   End
   Begin VB.TextBox txtPitch 
      Height          =   264
      Left            =   600
      TabIndex        =   5
      Text            =   "999999"
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox txtWBS 
      Height          =   264
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Text            =   "999999"
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtWBS 
      Height          =   264
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Text            =   "999999"
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtSosu 
      Height          =   264
      Left            =   600
      TabIndex        =   0
      Text            =   "999"
      Top             =   120
      Width           =   372
   End
   Begin VB.Label lblStart 
      Caption         =   "ｽﾀｯｸ"
      Height          =   252
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   492
   End
   Begin VB.Label lblPitch 
      Caption         =   "ﾋﾟｯﾁ"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   492
   End
   Begin VB.Label lblWBS 
      Caption         =   "WBS"
      Height          =   252
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   372
   End
   Begin VB.Label lblSosu 
      Caption         =   "層数"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "frmPlot"
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

    Dim intF0 As Integer ' ファイルNo.
    Dim i As Integer
    Dim intColor As Integer
    Dim lngStart As Long
    Dim strCommand As String
    Dim RetVal As Variant
    Dim strHPGLCommand As String
    Dim lngNCOrigin(1) As Long

    On Error GoTo FileWriteError

    With gudtWBInfo
        .lngWBS(X) = CLng(txtWBS(X)) * int1mm
        .lngWBS(Y) = CLng(txtWBS(Y)) * int1mm
        If CInt(txtSosu) > 2 Then ' 多層板
            mlngPitch = CLng(txtPitch.Text) * int1mm
            .lngStack(X) = (.lngWBS(X) - mlngPitch) / 2
            If cmbStack.Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Then
                .lngStack(Y) = CSng(txtWBS(Y).Text) * int1mm / 2
            Else
                .lngStack(Y) = CSng(cmbStack.Text) * int1mm
            End If
            ' 多層板はNC原点とスタック位置が同じ
            lngNCOrigin(X) = .lngStack(X)
            lngNCOrigin(Y) = .lngStack(Y)
        Else ' 両面板
            If cmbStack.Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Then
                .lngStack(Y) = CSng(txtWBS(Y).Text) * int1mm / 2
                lngNCOrigin(Y) = .lngStack(Y)
            ElseIf .strStart = "STACK" Then
                .lngStack(Y) = CSng(cmbStack.Text) * int1mm
                lngNCOrigin(Y) = .lngStack(Y)
            Else
                .lngStack(Y) = CSng(cmbStack.Text) * int1mm
                lngNCOrigin(Y) = .lngStack(Y) - 18000
            End If
            lngNCOrigin(X) = .lngStack(X)
        End If
    End With

    ' データの書き込み
    intF0 = FreeFile
    Open fTempPath & "NC2HPGL.TBL" For Output As #intF0
    Print #intF0, "NC"
    Print #intF0, Dir(gudtNCInfo(TH).strFileName)
    For i = 1 To intRow
        With gudtToolInfo(TH, i)
            If .intTNo > -1 Then
                Select Case .strColor
                    Case "BLACK"
                        intColor = 1
                    Case "RED"
                        intColor = 2
                    Case "GREEN"
                        intColor = 3
                    Case "BLUE"
                        intColor = 5
                    Case "MAGENTA"
                        intColor = 6
                    Case "CYAN"
                        intColor = 7
                End Select
                If i > 1 Then Print #intF0, " ";
                Print #intF0, "T_" & Format(CStr(.intTNo), "0#");
                Print #intF0, ":";
                Print #intF0, CStr(intColor);
                Print #intF0, ":";
                Print #intF0, CStr(Format(.sngDrill * 2, "#0.000"));
            End If
        End With
    Next
    Print #intF0, ""
    Print #intF0, txtWBS(X).Text;
    Print #intF0, ":";
    Print #intF0, txtWBS(Y).Text;
    Print #intF0, ":";
    With gudtNCInfo(frmMain.THNT)
        If CInt(txtSosu) > 2 Then
            Print #intF0, CStr(lngNCOrigin(X) / int1mm);
            Print #intF0, ":";
            Print #intF0, CStr(lngNCOrigin(Y) / int1mm)
            Print #intF0, "Multi"
        Else
            Print #intF0, CStr(lngNCOrigin(X) / int1mm);
            Print #intF0, ":";
            Print #intF0, CStr(lngNCOrigin(Y) / int1mm)
            Print #intF0, "Dual"
        End If
    End With
    If gudtToolInfo(NT, 1).intTNo > -1 Then 'NT
        Print #intF0, Dir(gudtNCInfo(NT).strFileName)
        For i = 1 To intRow
            With gudtToolInfo(NT, i)
                If .intTNo > -1 Then
                    Select Case .strColor
                        Case "BLACK"
                            intColor = 1
                        Case "RED"
                            intColor = 2
                        Case "GREEN"
                            intColor = 3
                        Case "BLUE"
                            intColor = 5
                        Case "MAGENTA"
                            intColor = 6
                        Case "CYAN"
                            intColor = 7
                    End Select
                    If i > 1 Then Print #intF0, " ";
                    Print #intF0, "T_" & Format(CStr(.intTNo), "0#");
                    Print #intF0, ":";
                    Print #intF0, CStr(intColor);
                    Print #intF0, ":";
                    Print #intF0, CStr(Format(.sngDrill * 2, "#0.000"));
                End If
            End With
        Next
        Print #intF0, ""
    Else
        Print #intF0, "null"
        Print #intF0, "null"
    End If
    Print #intF0, ""
    Close #intF0

    strHPGLCommand = GetSetting("NCView", _
                                "HPGLCommand", _
                                "Name", _
                                conDefaultHPGLCommand)

    If gudtNCInfo(NT).strFileName <> "" Then
        strCommand = strHPGLCommand & " " & _
                     gudtNCInfo(TH).strFileName & " " & _
                     gudtNCInfo(NT).strFileName
    Else
        strCommand = strHPGLCommand & " " & _
                     gudtNCInfo(TH).strFileName
    End If
    RetVal = Shell(strCommand, vbNormalNoFocus)
    Unload Me
    Exit Sub

FileWriteError:
    Close #intF0
    MsgBox "書き込みエラーです。", , "艦長、エラーです。"

End Sub

'*********************************************************
' 用  途: frmPlotのLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Load()

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
