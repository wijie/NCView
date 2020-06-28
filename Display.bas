Attribute VB_Name = "Display"
Option Explicit

'*********************************************************
' �p  �r: �s�N�`���{�b�N�X�Ɍ��̊G���T�[�N����`��
' ��  ��: objDisp: �s�N�`���{�b�N�X�̖��O
'         objBar: �v���O���X�o�[�̖��O
' �߂�l: ����
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
    Dim lngSize As Long ' �t�@�C���T�C�Y
    Dim dblMin() As Double
    Dim dblMax() As Double
    Dim dblPixelPerMM As Double
    Dim lngWBOrigin(1) As Long ' WB�����̍��W
    Dim dblNTOriginY As Double ' NT��Y�����_�̍��W
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
        .DrawWidth = 1 ' ������1Pixel

        ' NC�f�[�^�̍��G���A
        With gudtNCInfo(frmMain.THNT)
            dblMin = .dblMin
            dblMax = .dblMax
        End With
        If dblMin(X) > 0 Then dblMin(X) = 0
        If dblMax(X) < 0 Then dblMax(X) = 0
        If dblMin(Y) > 0 Then dblMin(Y) = 0
        If dblMax(Y) < 0 Then dblMax(Y) = 0

        ' WB�����G���A�Ɋ܂߂�
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

        ' �k�ڂ̌���
        dblFactor(X) = Round((dblNCWidth + 20) / Abs(.Width), intDigit)
        dblFactor(Y) = Round((dblNCHeight + 20) / Abs(.Height), intDigit)
        If dblFactor(X) > dblFactor(Y) Then
            frmMain.ScaleFactor = dblFactor(X)
        Else
            frmMain.ScaleFactor = dblFactor(Y)
        End If

        ' ���W�n�̐ݒ�
        .ScaleHeight = .Height * frmMain.ScaleFactor * -1 ' ��������Y+�����ɂ���
        .ScaleWidth = .Width * frmMain.ScaleFactor

        ' �]���̐ݒ�
        dblSpace(X) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)
        dblSpace(Y) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)

        ' �\���ʒu�̐ݒ�
        .ScaleLeft = Round(dblMin(X) - dblSpace(X), intDigit)
        .ScaleTop = Round(dblMax(Y) + dblSpace(Y), intDigit)

        ' �����\���p�s�N�`���[�{�b�N�X�̐ݒ�
        With objDisp(1)
            .Width = (dblNCWidth + 20) * 56.7 ' twip�P��
            .Height = (dblNCHeight + 20) * 56.7 ' twip�P��
            .ScaleWidth = dblNCWidth + 20
            .ScaleHeight = (dblNCHeight + 20) * -1
            .ScaleLeft = Round(dblMin(X) - 10, intDigit)
            .ScaleTop = Round(dblMax(Y) + 10, intDigit)
        End With

        ' 1�s�N�Z���̑傫��(mm�P��)
        dblPixelPerMM = Round(Screen.TwipsPerPixelX / 56.7, intDigit)

Display: ' ��ʂɏo��
        blnEventFlag = False
        objBar.Value = 50
        intF0 = FreeFile
        Open fTempPath & "NCView._$$" For Input As #intF0
        lngSize = LOF(intF0)
        Do While Not EOF(intF0)
            Input #intF0, dblX, dblY, dblR, lngColor

            ' �S�̕\��
            If dblR / frmMain.ScaleFactor < dblPixelPerMM / 2 Then
                ' ��ʏ��1�s�N�Z���ȉ���1�s�N�Z���̓_�ŕ`��
                objDisp(0).PSet (dblX, dblY), lngColor
            Else
                ' ���̑��̓T�[�N���ŕ`��
                objDisp(0).Circle (dblX, dblY), dblR, lngColor
'                objDisp(0).Circle (dblX, dblY), dblR - (dblPixelPerMM / 2), lngColor
            End If

            ' �����\��
            If (dblR * 2) - dblPixelPerMM < dblPixelPerMM Then
                ' ���C���̃Z���^�`�Z���^��1�s�N�Z���ȉ���1�s�N�Z���̓_�ŕ`��
                objDisp(1).PSet (dblX, dblY), lngColor
            Else
                ' ���C���̊O�������a�ƈ�v����l�ɕ`��
                objDisp(1).Circle (dblX, dblY), dblR - (dblPixelPerMM / 2), lngColor
            End If

            ' �v���O���X�o�[�͈̔͂�50�`100%
            If objBar.Value < Int(Seek(intF0) / lngSize * 50) + 50 Then
                objBar.Value = objBar.Value + 1
                ' ���x�ቺ��h����, DoEvents�̉񐔂𔼕��ɂ���
                blnEventFlag = Not blnEventFlag
                If blnEventFlag = True Then DoEvents
                If gblnCancel = False Then GoTo Quit
            End If
        Loop
        If frmMain.THNT = TH Then
            ' �S�̕\���p���_�}�[�N
            objDisp(0).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
            objDisp(0).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

            ' �S�̕\���pWB
            With gudtWBInfo
                objDisp(0).Line _
                (lngWBOrigin(X) / int1mm, lngWBOrigin(Y) / int1mm) _
                -Step(.lngWBS(X) / int1mm, .lngWBS(Y) / int1mm), RGB(0, 0, 0), B
            End With

            ' �����\���p���_�}�[�N
            objDisp(1).Line (-2.5, -2.5)-Step(5, 5), RGB(0, 0, 0)
            objDisp(1).Line (-2.5, 2.5)-Step(5, -5), RGB(0, 0, 0)

            ' �����\���pWB
            With gudtWBInfo
                objDisp(1).Line _
                (lngWBOrigin(X) / int1mm, lngWBOrigin(Y) / int1mm) _
                -Step(.lngWBS(X) / int1mm, .lngWBS(Y) / int1mm), RGB(0, 0, 0), B
            End With
        Else ' NT�̏ꍇ
            With gudtWBInfo
                If .intSosu > 2 Then
                    dblNTOriginY = 0&
                Else
                    dblNTOriginY = 180 - (.lngStack(Y) / int1mm)
                End If
            End With

            ' �S�̕\���p���_�}�[�N
            With objDisp(0)
                objDisp(0).Line (-3.5, dblNTOriginY)-Step(7, 0), RGB(255, 0, 0)
                objDisp(0).Line (0, dblNTOriginY - 3.5)-Step(0, 7), RGB(255, 0, 0)
            End With

            ' �����\���p���_�}�[�N
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
