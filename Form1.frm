VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Visualisasi Mengklik Tombol dari Coding"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()  'Command1 kelihatan masuk
                              'ke dalam (ditekan)
  Call SendMessage(Command1.hwnd, BM_SETSTATE, 1, _
                   ByVal 0&)
End Sub

Private Sub Command3_Click()  'Command1 normal kembali.
  Call SendMessage(Command1.hwnd, BM_SETSTATE, 0, _
                  ByVal 0&)
End Sub


