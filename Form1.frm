VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   9195
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton btnCommonLine 
      Caption         =   "Draw a common line"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton btnSmoothLine 
      Caption         =   "Draw a smooth line"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gdi As New clsGdip

Private Sub btnCommonLine_Click()
    gdi.SetHDC Me.hDC
    gdi.SetSmoothMode SmoothingModeDefault
    gdi.DrawLine 10, 10, 400, 60
End Sub

Private Sub btnSmoothLine_Click()
    gdi.SetHDC Me.hDC
    gdi.SetSmoothMode SmoothingModeAntiAlias8x8
    gdi.DrawLine 10, 30, 400, 90
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gdi = Nothing
End Sub
