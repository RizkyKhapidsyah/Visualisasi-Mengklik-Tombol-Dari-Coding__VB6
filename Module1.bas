Attribute VB_Name = "Module1"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const BM_SETSTATE = &HF3
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

