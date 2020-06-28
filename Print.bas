Attribute VB_Name = "Print"
Option Explicit

Sub sPrint(objDisp As Object, _
           objBar As Object)

    Dim i As Long, j As Long
    Dim lngColor As Long
    Dim intDigit As Integer
    Dim dblmm As Double
    Dim dblNCWidth As Double
    Dim dblNCHeight As Double
    Dim sngDrl As Single
    Dim strXY() As String
    Dim blnDrillHit As Boolean
    Dim intF0 As Integer
    Dim lngSize As Long
    Dim sngX As Single
    Dim sngY As Single
    Dim sngR As Single
    Dim dblSpace(1) As Double

    On Error GoTo PrinterError

    blnDrillHit = False
    intDigit = 7 '少数点何桁で丸めるか
    dblmm = 0.01 'NCは何ミリ単位か
                 '(Long型にSingle型を掛けると端数が出るのでDouble型にしている)
    sngDrl = 0.15 'ドリルの半径の初期値

With objDisp
    .ScaleMode = 6 'ScaleMode をミリに設定します
    .ScaleHeight = Abs(.ScaleHeight) * -1 '下向きをY+にする
    .DrawWidth = 4 '1Pixel 0.13mm

    'NCデータの作画エリア
    With gudtNCInfo(frmMain.THNT)
        dblNCWidth = Abs(.dblMax(X) - .dblMin(X))
        dblNCHeight = Abs(.dblMax(Y) - .dblMin(Y))
    End With

    '表示位置の設定
    '横向き
'    If .Orientation = vbPRORPortrait Then
'        .ScaleLeft = (dblNCWidth / 2) * -1
'        .ScaleTop = dblNCHeight / 2
'    Else
'        '縦向き
'        .ScaleLeft = (dblNCHeight / 2) * -1
'        .ScaleTop = dblNCWidth / 2
'    End If

    '余白の設定
'    dblSpace(X) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)
'    dblSpace(Y) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)
    dblSpace(X) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)
    dblSpace(Y) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)

    '表示位置の設定
'    .ScaleLeft = Round(gudtNCInfo.dblMin(X) - dblSpace(X), intDigit)
'    .ScaleTop = Round(gudtNCInfo.dblMax(Y) + dblSpace(Y), intDigit)
    .ScaleLeft = gudtNCInfo(frmMain.THNT).dblMin(X)
    .ScaleTop = gudtNCInfo(frmMain.THNT).dblMax(Y)

'    Debug.Print .ScaleLeft
'    Debug.Print .ScaleTop

'    .ScaleLeft = (.ScaleWidth / 2) * -1
'    .ScaleTop = (Abs(.ScaleHeight / 2) - dblNCHeight) / 2 + .ScaleHeight / 2
'    .CurrentX = .ScaleWidth / 2
'    .CurrentY = .ScaleHeight / 2

    'プリンタに出力
    App.Title = Dir(Command) 'Doc名の設定
    objBar.Value = objBar.Min
    intF0 = FreeFile
    Open fTempPath & "NCView._$$" For Input As #intF0
    lngSize = LOF(intF0)
    Do While Not EOF(1)
        Input #intF0, sngX, sngY, sngR, lngColor
        objDisp.Circle (sngX, sngY), sngR, lngColor
        objBar.Value = Int(Seek(intF0) / lngSize * 100)
    Loop
    With gudtNCInfo(frmMain.THNT)
        objDisp.Line (.lngStack(X) * -1, .lngStack(Y) * -1)- _
                     Step(.lngWBS(X), .lngWBS(Y)), , B
    End With
    Close #intF0
    objBar.Visible = False
    objDisp.EndDoc
End With

PrinterError:
    Close #intF0

End Sub

