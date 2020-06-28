Attribute VB_Name = "MainModule"
Option Explicit

' �萔�̐ݒ�
Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const R As Integer = 2
Public Const TH As Integer = 0
Public Const NT As Integer = 1
Public Const conTempFileName = "NCView._$$" ' �e���|�����t�@�C����
Public Const conCaption As String = "NCView"
Public Const intRow As Integer = 60 ' ��
Public Const int1mm As Integer = 100 ' 1mm�̒l
' HPGL�ϊ��v���O�����̃f�t�H���g
Public Const conDefaultHPGLCommand As String = "C:\usr\local\CygHPGL\NC2HPGL.EXE"

Public Type NCInfo
    strFileName As String ' NC�̃t�@�C����
    dblMin(1) As Double ' �ŏ��l X/Y
    dblMax(1) As Double ' �ő�l X/Y
End Type

Public Type WBInfo
    intSosu As Integer ' �w��
    lngWBS(1) As Long ' WBS X/Y
    lngStack(1) As Long ' Stack X/Y
    strStart As String
End Type

Public Type ToolInfo
    intTNo As Integer ' T�R�[�h
    sngDrill As Single ' �h�����a
    lngCount As Long ' ����
    strColor As String ' �F�̖��O
    lngColor As Long ' �F�ԍ�
End Type

Public gudtToolInfo(1, 1 To intRow) As ToolInfo ' TH/NT�̃c�[�����
Public gudtNCInfo(1) As NCInfo ' NC���
Public gudtWBInfo As WBInfo ' WB���
Public gblnCancel As Boolean

'*********************************************************
' �p  �r: NCView�̃X�^�[�g�A�b�v
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Sub Main()

    Dim strNC As String
    Dim strNCFileName As String
'    Dim frmNewMain As Form
'    Dim frmNewToolInfo As Form

    ' 2�d�N�����`�F�b�N
'    If App.PrevInstance Then
'        MsgBox "���łɋN������Ă��܂��I"
'        End
'    End If

    ' ����������
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
' �p  �r: �ϐ�������������
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Public Sub sInitialize()

    Dim i As Integer

    frmMain.ScaleFactor = 1

    ' �l�����͂���Ă��邩�ۂ��𔻒f����ׂɗL�蓾�Ȃ��l�ŏ���������
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
' �p  �r: NC�t�@�C����ϐ��Ɉ�C�ǂ݂���
' ��  ��: strNCFileName: NC�t�@�C����
' �߂�l: NC�f�[�^���ۂ��ƕԂ�
'*********************************************************

Public Function fReadNC(ByVal strNCFileName As String) As String

    Dim intF0 As Integer
    Dim bytBuf() As Byte
    Dim strNC As String

    ' NC��ǂݍ���
    intF0 = FreeFile
    Open strNCFileName For Binary As #intF0
    ReDim bytBuf(LOF(intF0))
    Get #intF0, , bytBuf
    Close #intF0
    strNC = StrConv(bytBuf, vbUnicode)
    Erase bytBuf ' �z��̃��������J������

    fReadNC = strNC

End Function

'*********************************************************
' �p  �r: NC�f�[�^��S�ʒǂ��ɓW�J, ���a, �F����ǉ�����
' ��  ��: strNC: NC�f�[�^
'         udtNCInfo: NC�t�@�C����, �ő�l/�ŏ��l���i�[����\����
'         udtToolInfo(): �h�����a�����i�[����\���̂̔z��
'         udtWBInfo: �w��, WBS etc...���i�[����\����
'         objBar: �v���O���X�o�[�̃I�u�W�F�N�g�ϐ�
' �߂�l: ����I�������True
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
    Dim lngCount As Long ' �v���O���X�o�[�̃J�E���^�p�ϐ�
    Dim lngNTIdou(1) As Long
    Dim blnEventFlag As Boolean
'    Dim objReg As New RegExp
'    Dim objMatches As Object
'    Dim objMatch As Object
'
'    With objReg
'        .Global = True '������S�̂�����
'        .IgnoreCase = True '�啶������������ʂ��Ȃ�
'        .Pattern = "X(-?[0-9]+)Y(-?[0-9]+)"
'    End With

    If frmMain.THNT = TH Then
            lngABS(X) = 0&
            lngABS(Y) = 0&
    Else ' NT�̏ꍇ
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

    ' ���s�R�[�h�𒲂ׂ�
    If InStr(strNC, vbCrLf) > 0 Then
        strEnter = vbCrLf
    ElseIf InStr(strNC, vbLf) > 0 Then
        strEnter = vbLf
    ElseIf InStr(strNC, vbCr) > 0 Then
        strEnter = vbCr
    End If

    ' �폜���镶�������������
    strNC = Replace(strNC, " ", "")
    ' ���C��,�T�u�ɕ�������
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" ' �ϐ��̃��������J������
    If UBound(strMainSub) = 1 Then
        strSubTmp = Split(strMainSub(0), "N", -1, vbTextCompare)
        For i = 1 To UBound(strSubTmp)
            intN = left(strSubTmp(i), 2) '�T�u�������̔ԍ����擾
            varSub(intN) = Split(strSubTmp(i), strEnter, -1, vbBinaryCompare)
        Next
        strMain = Split(strMainSub(1), strEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), strEnter, -1, vbBinaryCompare)
    End If
    ' �z��̃��������J������
    Erase strMainSub
    Erase strSubTmp

    strOutFile = fTempPath & conTempFileName
    ' �o�͂���
    blnEventFlag = False
    objBar.Visible = True
    intF1 = FreeFile
    Open strOutFile For Output As #intF1
    lngCount = UBound(strMain)
    For i = 0 To lngCount
        If strMain(i) Like "X*Y*" = True Then
            strXY = Split(Mid(strMain(i), 2), "Y", -1, vbTextCompare)
            lngABS(X) = lngABS(X) + CLng(strXY(X)) ' ���݂�X���W
            lngABS(Y) = lngABS(Y) + CLng(strXY(Y)) ' ���݂�Y���W
            If blnDrillHit = True Then
                With udtToolInfo(frmMain.THNT, intIndex)
                    .lngCount = .lngCount + 1
                End With
                Write #intF1, lngABS(X) / int1mm, lngABS(Y) / int1mm, sngDrl, lngColor
                ' �ŏ��l/�ő�l���Z�b�g����
                Call sSetMinMax(lngMin, lngMax, lngABS)
            End If
        ElseIf strMain(i) Like "M89" = True Then '�t�Z�b�g�`�F�b�N�p�R�[�h
            Write #intF1, lngABS(X) / int1mm, lngABS(Y) / int1mm, sngDrl, lngColor
            With udtToolInfo(frmMain.THNT, intIndex)
                .lngCount = .lngCount + 1
            End With
            ' �ŏ��l/�ő�l���Z�b�g����
            Call sSetMinMax(lngMin, lngMax, lngABS)
        ElseIf strMain(i) Like "G81" = True Then
            blnDrillHit = True
        ElseIf strMain(i) Like "G80" = True Then
            blnDrillHit = False
        ElseIf strMain(i) Like "M##" = True Then
            intSubNo = CInt(Mid(strMain(i), 2))
            ' �T�u�������[�͈̔͂�N44�`N97�ł���
            If intSubNo >= 44 And intSubNo <= 97 Then
                For j = 0 To UBound(varSub(intSubNo))
                    If varSub(intSubNo)(j) Like "X*Y*" = True Then
                        strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, vbTextCompare)
                        lngABS(X) = lngABS(X) + CLng(strXY(X)) '���݂�X���W
                        lngABS(Y) = lngABS(Y) + CLng(strXY(Y)) '���݂�Y���W
                        If blnDrillHit = True Then
                            With udtToolInfo(frmMain.THNT, intIndex)
                                .lngCount = .lngCount + 1
                            End With
                            Write #intF1, lngABS(X) / int1mm, lngABS(Y) / int1mm, sngDrl, lngColor
                            ' �ŏ��l/�ő�l���Z�b�g����
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
            If intIndex > intRow Then ' ��v����c�[����������Ȃ�������
                MsgBox "�H������������ĉ�����"
                GoTo Quit
            End If
        End If
        ' �v���O���X�o�[�͈̔͂�0�`50%
        If objBar.Value < Int(i / lngCount * 50) Then
            objBar.Value = objBar.Value + 1
            ' ���x�ቺ��h����, DoEvents�̉񐔂𔼕��ɂ���
            blnEventFlag = Not blnEventFlag
            If blnEventFlag = True Then DoEvents
            If gblnCancel = False Then GoTo Quit
        End If
    Next
    Close #intF1
    Erase strMain ' �z��̃��������J������
'    objBar.Visible = False
    ' NC�f�[�^�̍ő�/�ŏ��l���Z�b�g
    With udtNCInfo
        .dblMin(X) = lngMin(X) / int1mm
        .dblMin(Y) = lngMin(Y) / int1mm
        .dblMax(X) = lngMax(X) / int1mm
        .dblMax(Y) = lngMax(Y) / int1mm
    End With

    fConvertNC = True ' ����I������True��Ԃ�
    Exit Function

Quit:
    objBar.Visible = False
    Close #intF1
    fConvertNC = False

End Function

'*********************************************************
' �p  �r: NT�̃f�[�^�̍ŏ��l/�ő�l��ݒ肷��
' ��  ��: lngMIN(): ���݂܂ł̍ŏ��lX/Y�̔z��
'         lngMAX(): ���݂܂ł̍ő�lX/Y�̔z��
'         lngABS(): ���݂̍��WX/Y�̔z��
' �߂�l: ����
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
' �p  �r: NT�̃f�[�^����ړ��ʂ��擾����
' ��  ��: strNC: NC�f�[�^
' �߂�l: �ړ��ʂ�"X�`Y�`"�̌`���ŕԂ�
'*********************************************************

Public Function fGetNTIdou(ByVal strNC As String) As String

    Dim strMainSub() As String
    Dim strMain() As String
    Dim strEnter As String
    Dim i As Long

    ' ���s�R�[�h�𒲂ׂ�
    If InStr(strNC, vbCrLf) > 0 Then
        strEnter = vbCrLf
    ElseIf InStr(strNC, vbLf) > 0 Then
        strEnter = vbLf
    ElseIf InStr(strNC, vbCr) > 0 Then
        strEnter = vbCr
    End If

    ' ���C��,�T�u�ɕ�������
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" ' �ϐ��̃��������J������
    If UBound(strMainSub) = 1 Then
        strMain = Split(strMainSub(1), strEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), strEnter, -1, vbBinaryCompare)
    End If
    ' �z��̃��������J������
    Erase strMainSub

    ' NT�̈ړ��ʂ𒲂ׂ�
    For i = 0 To UBound(strMain)
        If strMain(i) Like "X*Y*" = True Then
            fGetNTIdou = strMain(i) ' �ړ��ʂ�Ԃ�
            Exit For
        ElseIf strMain(i) Like "T*" = True Then
            Exit For
        ElseIf strMain(i) Like "G81" = True Then
            Exit For
        End If
    Next
    Erase strMain ' �z��̃��������J������

End Function

'*********************************************************
' �p  �r: ���ϐ�TEMP�̒l���擾����
' ��  ��: ����
' �߂�l: ���ϐ�TEMP�̒l��Ԃ�
'*********************************************************

Public Function fTempPath() As String

    ' �v���O�����I���܂�TempPath�̓��e��ێ�
    Static TempPath As String

    ' �r���Ńf�B���N�g��-���ύX����Ă�Temp�f�B���N�g��-���m��
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP") ' �f�B���N�g��-���擾
        ' ���[�g�f�B���N�g���[���̔��f
        If right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath

End Function

'*********************************************************
' �p  �r: NCArray.exe����DDE��p���đw��, WBS�����擾����
' ��  ��: udtWBInfo: �擾�����l���i�[����\����
' �߂�l: ����
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
                If .lngStack(X) = 0 Then ' �ݒ肳��Ă��Ȃ��������Z�b�g����
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
