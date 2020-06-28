Attribute VB_Name = "MainModule"
Option Explicit

' 定数の設定
Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const R As Integer = 2
Public Const TH As Integer = 0
Public Const NT As Integer = 1
Public Const conTempFileName = "NCView._$$" ' テンポラリファイル名
Public Const conCaption As String = "NCView"
Public Const intRow As Integer = 60 ' 列数
Public Const int1mm As Integer = 100 ' 1mmの値
' HPGL変換プログラムのデフォルト
Public Const conDefaultHPGLCommand As String = "C:\usr\local\CygHPGL\NC2HPGL.EXE"

Public Type NCInfo
    strFileName As String ' NCのファイル名
    dblMin(1) As Double ' 最小値 X/Y
    dblMax(1) As Double ' 最大値 X/Y
End Type

Public Type WBInfo
    intSosu As Integer ' 層数
    lngWBS(1) As Long ' WBS X/Y
    lngStack(1) As Long ' Stack X/Y
    strStart As String
End Type

Public Type ToolInfo
    intTNo As Integer ' Tコード
    sngDrill As Single ' ドリル径
    lngCount As Long ' 穴数
    strColor As String ' 色の名前
    lngColor As Long ' 色番号
End Type

Public gudtToolInfo(1, 1 To intRow) As ToolInfo ' TH/NTのツール情報
Public gudtNCInfo(1) As NCInfo ' NC情報
Public gudtWBInfo As WBInfo ' WB情報
Public gblnCancel As Boolean

'*********************************************************
' 用  途: NCViewのスタートアップ
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Sub Main()

    Dim strNC As String
    Dim strNCFileName As String
'    Dim frmNewMain As Form
'    Dim frmNewToolInfo As Form

    ' 2重起動をチェック
'    If App.PrevInstance Then
'        MsgBox "すでに起動されています！"
'        End
'    End If

    ' 初期化する
    Call sInitialize

'    Set frmNewMain = New frmMain
'    Set frmNewToolInfo = New frmToolInfo
'    frmNewMain.Show
'    frmNewToolInfo.Show vbModal

    Load frmMain
    Load frmToolInfo
    frmMain.Show
    frmToolInfo.Show vbModal

End Sub

'*********************************************************
' 用  途: 変数を初期化する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Public Sub sInitialize()

    Dim i As Integer

    frmMain.ScaleFactor = 1

    ' 値が入力されているか否かを判断する為に有り得ない値で初期化する
    For i = 1 To intRow
        With gudtToolInfo(0, i) 'TH
            .intTNo = -1
            .lngColor = -1
            .lngCount = -1
            .sngDrill = -1
            .strColor = ""
        End With
        With gudtToolInfo(1, i) 'NT
            .intTNo = -1
            .lngColor = -1
            .lngCount = -1
            .sngDrill = -1
            .strColor = "GREEN"
        End With
    Next

    With gudtNCInfo(frmMain.THNT)
        .strFileName = ""
        .dblMax(X) = 0
        .dblMax(Y) = 0
        .dblMin(X) = 0
        .dblMin(Y) = 0
    End With

    With gudtWBInfo
        .intSosu = 0
        .lngWBS(X) = 0
        .lngWBS(Y) = 0
        .lngStack(X) = 0
        .lngStack(Y) = 0
        .strStart = ""
    End With

End Sub

'*********************************************************
' 用  途: NCファイルを変数に一気読みする
' 引  数: strNCFileName: NCファイル名
' 戻り値: NCデータを丸ごと返す
'*********************************************************

Public Function fReadNC(ByVal strNCFileName As String) As String

    Dim intF0 As Integer
    Dim bytBuf() As Byte
    Dim strNC As String

    ' NCを読み込む
    intF0 = FreeFile
    Open strNCFileName For Binary As #intF0
    ReDim bytBuf(LOF(intF0))
    Get #intF0, , bytBuf
    Close #intF0
    strNC = StrConv(bytBuf, vbUnicode)
    Erase bytBuf ' 配列のメモリを開放する

    fReadNC = strNC

End Function

'*********************************************************
' 用  途: NCデータを全面追いに展開, 穴径, 色情報を追加する
' 引  数: strNC: NCデータ
'         udtNCInfo: NCファイル名, 最大値/最小値を格納する構造体
'         udtToolInfo(): ドリル径情報を格納する構造体の配列
'         udtWBInfo: 層数, WBS etc...を格納する構造体
'         objBar: プログレスバーのオブジェクト変数
' 戻り値: 正常終了すればTrue
'*********************************************************

Public Function fConvertNC(ByVal strNC As String, _
                           ByRef udtNCInfo As NCInfo, _
                           ByRef udtToolInfo() As ToolInfo, _
                           ByRef udtWBInfo As WBInfo, _
                           ByRef objBar As Object) As Boolean

    Dim lngABS(1) As Long
    Dim lngMax(1) As Long
    Dim lngMin(1) As Long
    Dim intF1 As Integer
    Dim strMainSub() As String
    Dim varSub(44 To 97) As Variant
    Dim strMain() As String
    Dim strSubTmp() As String
    Dim strEnter As String
    Dim intN As Integer
    Dim blnDrillHit As Boolean
    Dim i As Long
    Dim j As Long
    Dim intIndex As Integer
    Dim sngDrl As Single
    Dim strXY() As String
    Dim strOutFile As String
    Dim lngColor As Long
    Dim intSubNo As Integer
    Dim intTool As Integer
    Dim lngCount As Long ' プログレスバーのカウンタ用変数
    Dim lngNTIdou(1) As Long
    Dim blnEventFlag As Boolean
'    Dim objReg As New RegExp
'    Dim objMatches As Object
'    Dim objMatch As Object
'
'    With objReg
'        .Global = True '文字列全体を処理
'        .IgnoreCase = True '大文字小文字を区別しない
'        .Pattern = "X(-?[0-9]+)Y(-?[0-9]+)"
'    End With

    If frmMain.THNT = TH Then
            lngABS(X) = 0&
            lngABS(Y) = 0&
    Else ' NTの場合
        With udtWBInfo
            If .intSosu > 2 Then
                lngNTIdou(X) = int1mm
                lngNTIdou(Y) = .lngStack(Y)
            Else
                lngNTIdou(X) = 0
                lngNTIdou(Y) = .lngStack(Y) - 180& * int1mm
            End If
        End With
        lngABS(X) = 0& - lngNTIdou(X)
        lngABS(Y) = 0& - lngNTIdou(Y)
    End If
    lngMax(X) = -2147483647
    lngMax(Y) = -2147483647
    lngMin(X) = 2147483647
    lngMin(Y) = 2147483647
    blnDrillHit = False
    intTool = -32767
    With objBar
        .Max = 100
        .Min = 0
        .Value = .Min
    End With

    ' 改行コードを調べる
    If InStr(strNC, vbCrLf) > 0 Then
        strEnter = vbCrLf
    ElseIf InStr(strNC, vbLf) > 0 Then
        strEnter = vbLf
    ElseIf InStr(strNC, vbCr) > 0 Then
        strEnter = vbCr
    End If

    ' 削除する文字列を処理する
    strNC = Replace(strNC, " ", "")
    ' メイン,サブに分割する
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" ' 変数のメモリを開放する
    If UBound(strMainSub) = 1 Then
        strSubTmp = Split(strMainSub(0), "N", -1, vbTextCompare)
        For i = 1 To UBound(strSubTmp)
            intN = left(strSubTmp(i), 2) 'サブメモリの番号を取得
            varSub(intN) = Split(strSubTmp(i), strEnter, -1, vbBinaryCompare)
        Next
        strMain = Split(strMainSub(1), strEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), strEnter, -1, vbBinaryCompare)
    End If
    ' 配列のメモリを開放する
    Erase strMainSub
    Erase strSubTmp

    strOutFile = fTempPath & conTempFileName
    ' 出力する
    blnEventFlag = False
    objBar.Visible = True
    intF1 = FreeFile
    Open strOutFile For Output As #intF1
    lngCount = UBound(strMain)
    For i = 0 To lngCount
        If strMain(i) Like "X*Y*" = True Then
            strXY = Split(Mid(strMain(i), 2), "Y", -1, vbTextCompare)
            lngABS(X) = lngABS(X) + CLng(strXY(X)) ' 現在のX座標
            lngABS(Y) = lngABS(Y) + CLng(strXY(Y)) ' 現在のY座標
            If blnDrillHit = True Then
                With udtToolInfo(frmMain.THNT, intIndex)
                    .lngCount = .lngCount + 1
                End With
                Write #intF1, lngABS(X) / int1mm, lngABS(Y) / int1mm, sngDrl, lngColor
                ' 最小値/最大値をセットする
                Call sSetMinMax(lngMin, lngMax, lngABS)
            End If
        ElseIf strMain(i) Like "M89" = True Then '逆セットチェック用コード
            Write #intF1, lngABS(X) / int1mm, lngABS(Y) / int1mm, sngDrl, lngColor
            With udtToolInfo(frmMain.THNT, intIndex)
                .lngCount = .lngCount + 1
            End With
            ' 最小値/最大値をセットする
            Call sSetMinMax(lngMin, lngMax, lngABS)
        ElseIf strMain(i) Like "G81" = True Then
            blnDrillHit = True
        ElseIf strMain(i) Like "G80" = True Then
            blnDrillHit = False
        ElseIf strMain(i) Like "M##" = True Then
            intSubNo = CInt(Mid(strMain(i), 2))
            ' サブメモリーの範囲はN44～N97である
            If intSubNo >= 44 And intSubNo <= 97 Then
                For j = 0 To UBound(varSub(intSubNo))
                    If varSub(intSubNo)(j) Like "X*Y*" = True Then
                        strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, vbTextCompare)
                        lngABS(X) = lngABS(X) + CLng(strXY(X)) '現在のX座標
                        lngABS(Y) = lngABS(Y) + CLng(strXY(Y)) '現在のY座標
                        If blnDrillHit = True Then
                            With udtToolInfo(frmMain.THNT, intIndex)
                                .lngCount = .lngCount + 1
                            End With
                            Write #intF1, lngABS(X) / int1mm, lngABS(Y) / int1mm, sngDrl, lngColor
                            ' 最小値/最大値をセットする
                            Call sSetMinMax(lngMin, lngMax, lngABS)
                        End If
                    ElseIf varSub(intSubNo)(j) Like "G81" = True Then
                        blnDrillHit = True
                    ElseIf varSub(intSubNo)(j) Like "G80" = True Then
                        blnDrillHit = False
                    End If
                Next
            End If
        ElseIf strMain(i) Like "T*" = True Then
            intTool = CInt(Mid(strMain(i), 2))
            For intIndex = 1 To intRow
                With udtToolInfo(frmMain.THNT, intIndex)
                    If intTool = CInt(.intTNo) Then
                        sngDrl = .sngDrill
                        .lngCount = 0
                        lngColor = .lngColor
                        Exit For
                    End If
                End With
            Next
            If intIndex > intRow Then ' 一致するツールが見つからなかった時
                MsgBox "工具情報を見直して下さい"
                GoTo Quit
            End If
        End If
        ' プログレスバーの範囲は0～50%
        If objBar.Value < Int(i / lngCount * 50) Then
            objBar.Value = objBar.Value + 1
            ' 速度低下を防ぐ為, DoEventsの回数を半分にする
            blnEventFlag = Not blnEventFlag
            If blnEventFlag = True Then DoEvents
            If gblnCancel = False Then GoTo Quit
        End If
    Next
    Close #intF1
    Erase strMain ' 配列のメモリを開放する
'    objBar.Visible = False
    ' NCデータの最大/最小値をセット
    With udtNCInfo
        .dblMin(X) = lngMin(X) / int1mm
        .dblMin(Y) = lngMin(Y) / int1mm
        .dblMax(X) = lngMax(X) / int1mm
        .dblMax(Y) = lngMax(Y) / int1mm
    End With

    fConvertNC = True ' 正常終了時はTrueを返す
    Exit Function

Quit:
    objBar.Visible = False
    Close #intF1
    fConvertNC = False

End Function

'*********************************************************
' 用  途: NTのデータの最小値/最大値を設定する
' 引  数: lngMIN(): 現在までの最小値X/Yの配列
'         lngMAX(): 現在までの最大値X/Yの配列
'         lngABS(): 現在の座標X/Yの配列
' 戻り値: 無し
'*********************************************************

Private Sub sSetMinMax(ByRef lngMin() As Long, _
                       ByRef lngMax() As Long, _
                       ByRef lngABS() As Long)

    If lngMax(X) < lngABS(X) Then lngMax(X) = lngABS(X)
    If lngMin(X) > lngABS(X) Then lngMin(X) = lngABS(X)
    If lngMax(Y) < lngABS(Y) Then lngMax(Y) = lngABS(Y)
    If lngMin(Y) > lngABS(Y) Then lngMin(Y) = lngABS(Y)

End Sub

'*********************************************************
' 用  途: NTのデータから移動量を取得する
' 引  数: strNC: NCデータ
' 戻り値: 移動量を"X～Y～"の形式で返す
'*********************************************************

Public Function fGetNTIdou(ByVal strNC As String) As String

    Dim strMainSub() As String
    Dim strMain() As String
    Dim strEnter As String
    Dim i As Long

    ' 改行コードを調べる
    If InStr(strNC, vbCrLf) > 0 Then
        strEnter = vbCrLf
    ElseIf InStr(strNC, vbLf) > 0 Then
        strEnter = vbLf
    ElseIf InStr(strNC, vbCr) > 0 Then
        strEnter = vbCr
    End If

    ' メイン,サブに分割する
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" ' 変数のメモリを開放する
    If UBound(strMainSub) = 1 Then
        strMain = Split(strMainSub(1), strEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), strEnter, -1, vbBinaryCompare)
    End If
    ' 配列のメモリを開放する
    Erase strMainSub

    ' NTの移動量を調べる
    For i = 0 To UBound(strMain)
        If strMain(i) Like "X*Y*" = True Then
            fGetNTIdou = strMain(i) ' 移動量を返す
            Exit For
        ElseIf strMain(i) Like "T*" = True Then
            Exit For
        ElseIf strMain(i) Like "G81" = True Then
            Exit For
        End If
    Next
    Erase strMain ' 配列のメモリを開放する

End Function

'*********************************************************
' 用  途: 環境変数TEMPの値を取得する
' 引  数: 無し
' 戻り値: 環境変数TEMPの値を返す
'*********************************************************

Public Function fTempPath() As String

    ' プログラム終了までTempPathの内容を保持
    Static TempPath As String

    ' 途中でディレクトリ-が変更されてもTempディレクトリ-を確保
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP") ' ディレクトリ-を取得
        ' ルートディレクトリーかの判断
        If right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath

End Function

'*********************************************************
' 用  途: NCArray.exeからDDEを用いて層数, WBS等を取得する
' 引  数: udtWBInfo: 取得した値を格納する構造体
' 戻り値: 無し
'*********************************************************

Public Sub sDDElink(ByRef udtWBInfo As WBInfo)

    On Error GoTo Trap

    With frmMain.lblDDE
        If .LinkMode = vbLinkNone Then
            .LinkTopic = "NCArray|frmMain"
            .LinkItem = "lblDDE"
            .LinkMode = vbLinkManual
        End If

        .LinkExecute "Sosu"
        .LinkRequest
        If .Caption <> 0 Then
            udtWBInfo.intSosu = CInt(.Caption)
        End If

        .LinkExecute "WBSX"
        .LinkRequest
        If .Caption <> 0 Then
            udtWBInfo.lngWBS(X) = CLng(.Caption)
        End If

        .LinkExecute "WBSY"
        .LinkRequest
        If .Caption <> 0 Then
            udtWBInfo.lngWBS(Y) = CLng(.Caption)
        End If

        .LinkExecute "Stack"
        .LinkRequest
        If .Caption <> 0 Then
            udtWBInfo.lngStack(Y) = CLng(.Caption)
            With udtWBInfo
                If .lngStack(X) = 0 Then ' 設定されていない時だけセットする
                    If .intSosu > 2 Then
                        .lngStack(X) = 5 * int1mm
                    Else
                        .lngStack(X) = 4 * int1mm
                    End If
                End If
            End With
        End If

        .LinkExecute "Start"
        .LinkRequest
        If .Caption <> "" Then
            udtWBInfo.strStart = UCase(.Caption)
        End If
    End With

Trap:

End Sub
