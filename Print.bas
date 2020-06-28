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
    intDigit = 7 '�����_�����Ŋۂ߂邩
    dblmm = 0.01 'NC�͉��~���P�ʂ�
                 '(Long�^��Single�^���|����ƒ[�����o��̂�Double�^�ɂ��Ă���)
    sngDrl = 0.15 '�h�����̔��a�̏����l

With objDisp
    .ScaleMode = 6 'ScaleMode ���~���ɐݒ肵�܂�
    .ScaleHeight = Abs(.ScaleHeight) * -1 '��������Y+�ɂ���
    .DrawWidth = 4 '1Pixel 0.13mm

    'NC�f�[�^�̍��G���A
    With gudtNCInfo(frmMain.THNT)
        dblNCWidth = Abs(.dblMax(X) - .dblMin(X))
        dblNCHeight = Abs(.dblMax(Y) - .dblMin(Y))
    End With

    '�\���ʒu�̐ݒ�
    '������
'    If .Orientation = vbPRORPortrait Then
'        .ScaleLeft = (dblNCWidth / 2) * -1
'        .ScaleTop = dblNCHeight / 2
'    Else
'        '�c����
'        .ScaleLeft = (dblNCHeight / 2) * -1
'        .ScaleTop = dblNCWidth / 2
'    End If

    '�]���̐ݒ�
'    dblSpace(X) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)
'    dblSpace(Y) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)
    dblSpace(X) = Round((Abs(.ScaleHeight) - dblNCHeight) / 2, intDigit)
    dblSpace(Y) = Round((Abs(.ScaleWidth) - dblNCWidth) / 2, intDigit)

    '�\���ʒu�̐ݒ�
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

    '�v�����^�ɏo��
    App.Title = Dir(Command) 'Doc���̐ݒ�
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

