VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8304
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8736
   ScaleWidth      =   8304
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Square"
      Height          =   192
      Left            =   6300
      TabIndex        =   1
      Top             =   252
      Width           =   936
   End
   Begin VB.TextBox Text1 
      Height          =   348
      Left            =   252
      TabIndex        =   0
      Text            =   "http://www.vbforums.com"
      Top             =   168
      Width           =   5976
   End
   Begin VB.Image Image1 
      Height          =   7656
      Left            =   252
      Stretch         =   -1  'True
      Top             =   756
      Width           =   7656
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

Private Sub Form_Resize()
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    If WindowState <> vbMinimized Then
        lWidth = ScaleWidth - Image1.Left - Image1.Left
        lHeight = ScaleHeight - Image1.Top - Image1.Left
        If lWidth > lHeight Then
            lWidth = lHeight
        End If
        Image1.Width = lWidth
        Image1.Height = lWidth
    End If
End Sub

Private Sub Text1_Change()
    Set Image1.Picture = QRCodegenBarcode(Text1.Text, SquareModules:=(Check1.Value = vbChecked))
End Sub

Private Sub Check1_Click()
    Text1_Change
End Sub
