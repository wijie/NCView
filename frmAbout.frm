VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾊﾞｰｼﾞｮﾝ情報"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdAbout 
      Caption         =   "OK"
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1092
   End
   Begin VB.Line linAbout2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      Top             =   120
      Width           =   1500
   End
   Begin VB.Line linAbout1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   240
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblAbout2 
      Caption         =   "Copyright (C) 2001 WATABE Eiji"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblAbout1 
      Caption         =   "ﾊﾞｰｼﾞｮﾝ情報"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblAbout3 
      Caption         =   "自画像(嘘)"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************************
' 用  途: OKボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdAbout_Click()

    Unload Me

End Sub

'*********************************************************
' 用  途: frmAboutのLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Load()

    ' フォームの表示位置をfrmMainの中央に設定する
    Me.left = _
        frmMain.left + (frmMain.Width - Me.Width) / 2
    Me.top = _
        frmMain.top + (frmMain.Height - Me.Height) / 2

    With lblAbout1
        .AutoSize = True
        .Caption = _
            "NCデータPreview && Plotシステム" & vbCrLf & _
            "Ver." & _
            App.Major & "." & _
            App.Minor & "." & _
            App.Revision
    End With

    With lblAbout2
        .AutoSize = True
        .Caption = _
            "Copyright (C) 2001" & vbCrLf & _
            "WATABE Eiji"
    End With

    With lblAbout3
        .AutoSize = True
        .top = Image1.top + Image1.Height + 60
        .left = Image1.left + (Image1.Width - .Width) / 2
    End With

End Sub

