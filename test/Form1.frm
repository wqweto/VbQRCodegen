VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7812
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7812
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   348
      Left            =   252
      TabIndex        =   0
      Text            =   "http://www.vbforums.com"
      Top             =   168
      Width           =   6480
   End
   Begin VB.Image Image1 
      Height          =   6480
      Left            =   252
      Stretch         =   -1  'True
      Top             =   756
      Width           =   6480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' QR Code generator library (VB6/VBA)
'
' Copyright (c) Project Nayuki. (MIT License)
' https://www.nayuki.io/page/qr-code-generator-library
'
' Copyright (c) wqweto@gmail.com (MIT License)
'
'=========================================================================
Option Explicit
DefObj A-Z

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetData Image1.Picture
    End If
End Sub

Private Sub Form_Load()
    Text1_Change
End Sub

Private Sub Text1_Change()
    Set Image1.Picture = QRCodegenBarcode(Text1.Text)
End Sub
