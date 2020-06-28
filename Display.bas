Attribute VB_Name = "Display"
Option Explicit

'*********************************************************
' 用  途: ピクチャボックスに穴の絵をサークルを描く
' 引  数: objDisp: ピクチャボックスの名前
'         objBar: プログレスバーの名前
' 戻り値: 無し
'*********************************************************

Sub sDispNC(objDisp As Object, _
            objBar As Object)

    Dim dblX As Double
    Dim dblY As Double
    Dim dblR As Double
    Dim lngColor As Long
    Dim intDigit As Integer
    Dim dblSpace(1) As Double
    Dim dblNCWidth As Double
    Dim dblNCHeight As Double
    Dim dblFactor(1) As Double
    Dim intF0 As Integer
    Dim lngSize As Long ' ファイルサイズ
    Dim dblMin() As Double
    Dim dblMax() As Double
    Dim dblPixelPerMM As Double
    Dim lngWBOrigin(1) As Long ' WB左下の座標
    Dim dblNTOriginY As Double ' NTのY側原点の座標
    Dim blnEventFlag As Boolean

    intDigit = 7

    If frmMain.THNT = NT Then GoTo Display

    With gudtWBInfo
        If .lngWBS(X) > 0 And .lngWBS(Y) > 0 Then
            If .intSosu > 2 Then
                lngWBOrigin(X) = .lngStack(X) * -1&
                lngWBOrigin(Y) = .lngStack(Y) * -1&
            Else
                lngWBOrigin(X) = .lngStack(X) * -1&
                If .strStart = "MACHINE" Then
                    lngWBOrigin(Y) = 18000& - .lngStack(Y)
                Else
                    lngWBOrigin(Y) = .lngStack(Y) * -1&
                End If
            End If
        End If
    End With

    With objDisp(0)
        .DrawWidth = 1 ' 線幅は1Pixel

        ' NCデータの作画エリア
        With gudtNCInfo(frmMain.THNT)
            dblMin = .dblMin
            dblMax = .dblMax
        End With
        If dblMin(X) > 0 Then dblMin(X) = 0
        If dblMax(X) < 0 Then dblMax(X) = 0
        If dblMin(Y) > 0 Then dblMin(Y) = 0
        If dblMax(Y) < 0 Then dblMax(Y) = 0

        ' WBを作画エリアに含める
        With gudtWBInfo
            If lngWBOrigin(X) < dblMin(X) Then
                dblMin(X) = lngWBOrigin(X) / int1mm
            End If
            If lngWBOrigin(Y) < dblMin(Y) Then
                dblMin(Y) = lngWBOrigin(Y) / int1mm
            End If
            If (.lngWBS(X) + lngWBOrigin(X)) / int1mm > dblMax(X) Then
                dblMax(X) = (.lngWBS(X) + lngWBOrigin(X)) / int1mm
            End If
            If (.lngWBS(Y) + lngWBOrigin(Y)) / int1mm > dblMax(Y) Then
                dblMax(Y) = (.lngWBS(Y) + lngWBOrigin(Y)) / int1mm
            End If
        End With
        dblNCWidth = Abs(dblMax(X) - dblMin(X))
        dblNCHeight = Abs(dblMax(Y) - dblMin(Y))

        ' 縮尺の決定
        dblFactor(X) = Round((dblNCWidth + 20) / Abs(.Width), intDigit)
        dblFactor(Y) = Round((dblNCHeight + 20) / Abs(.Height), intDigit)
        If dblFactor(X) > dblFactor(Y) Then
            frmMain.ScaleFactor = dblFactor(X)
        Else
            frmMain.ScaleFactor = dblFactor(Y)
        End If

        ' 座標系の設定
        .ScaleHeight = .Height * frmMain.ScaleFactor * -1 ' 下から上をY+方向にする
        .ScaleWidth = .Width * frmMain.ScaleFactor

        ' 余白の設定
        dblSpace(X) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)
        dblSpace(Y) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)

        ' 表示位置の設定
        .ScaleLeft = Round(dblMin(X) - dblSpace(X), intDigit)
        .ScaleTop = Round(dblMax(Y) + dblSpace(Y), intDigit)

        ' 正寸表示用ピクチャーボックスの設定
        With objDisp(1)
            .Width = (dblNCWidth + 20) * 56.7 ' twip単位
            .Height = (dblNCHeight + 20) * 56.7 ' twip単位
            .ScaleWidth = dblNCWidth + 20
            .ScaleHeight = (dblNCHeight + 20) * -1
            .ScaleLeft = Round(dblMin(X) - 10, intDigit)
            .ScaleTop = Round(dblMax(Y) + 10, intDigit)
        End With

        ' 1ピクセルの大きさ(mm単位)
        dblPixelPerMM = Round(Screen.TwipsPerPixelX / 56.7, intDigit)

Display: ' 画面に出力
        blnEventFlag = False
        objBar.Value = 50
        intF0 = FreeFile
        Open fTempPath & "NCView._$$" For Input As #intF0
        lngSize = LOF(intF0)
        Do While Not EOF(intF0)
            Input #intF0, dblX, dblY, dblR, lngColor

            ' 全体表示
            If dblR / frmMain.ScaleFactor < dblPixelPerMM / 2 Then
                ' 画面上で1ピクセル以下は1ピクセルの点で描く
                objDisp(0).PSet (dblX, dblY), lngColor
            Else
                ' その他はサークルで描く
                objDisp(0).Circle (dblX, dblY), dblR, lngColor
'                objDisp(0).Circle (dblX, dblY), dblR - (dblPixelPerMM / 2), lngColor
            End If

            ' 正寸表示
            If (dblR * 2) - dblPixelPerMM < dblPixelPerMM Then
                ' ラインのセンタ～センタで1ピクセル以下は1ピクセルの点で描く
                objDisp(1).PSet (dblX, dblY), lngColor
            Else
                ' ラインの外側が穴径と一致する様に描く
                objDisp(1).Circle (dblX, dblY), dblR - (dblPixelPerMM / 2), lngColor
            End If

            ' プログレスバーの範囲は50～100%
            If objBar.Value < Int(Seek(intF0) / lngSize * 50) + 50 Then
                objBar.Value = objBar.Value + 1
                ' 速度低下を防ぐ為, DoEventsの回数を半分にする
                blnEventFlag = Not blnEventFlag
                If blnEventFlag = True Then DoEvents
                If gblnCancel = False Then GoTo Quit
            End If
        Loop
        If frmMain.THNT = TH Then
            ' 全体表示用原点マーク
            objDisp(0).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
            objDisp(0).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

            ' 全体表示用WB
            With gudtWBInfo
                objDisp(0).Line _
                (lngWBOrigin(X) / int1mm, lngWBOrigin(Y) / int1mm) _
                -Step(.lngWBS(X) / int1mm, .lngWBS(Y) / int1mm), RGB(0, 0, 0), B
            End With

            ' 正寸表示用原点マーク
            objDisp(1).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
            objDisp(1).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

            ' 正寸表示用WB
            With gudtWBInfo
                objDisp(1).Line _
                (lngWBOrigin(X) / int1mm, lngWBOrigin(Y) / int1mm) _
                -Step(.lngWBS(X) / int1mm, .lngWBS(Y) / int1mm), RGB(0, 0, 0), B
            End With
        Else ' NTの場合
            With gudtWBInfo
                If .intSosu > 2 Then
                    dblNTOriginY = 0&
                Else
                    dblNTOriginY = 180 - (.lngStack(Y) / int1mm)
                End If
            End With

            ' 全体表示用原点マーク
            With objDisp(0)
                objDisp(0).Line (-3.5, dblNTOriginY)-Step(7, 0), RGB(255, 0, 0)
                objDisp(0).Line (0, dblNTOriginY - 3.5)-Step(0, 7), RGB(255, 0, 0)
            End With

            ' 正寸表示用原点マーク
            With objDisp(1)
                objDisp(1).Line (-3.5, dblNTOriginY)-Step(7, 0), RGB(255, 0, 0)
                objDisp(1).Line (0, dblNTOriginY - 3.5)-Step(0, 7), RGB(255, 0, 0)
            End With
        End If
Quit:
        Close #intF0
        objBar.Visible = False
    End With

End Sub
