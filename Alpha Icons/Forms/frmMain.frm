VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to add Alpha Icons in VB"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' Sets an Alpha Icon as the Project Icon, this is not visible at Design Time
    ' Once compiled the Icon will be displayed on the exe file
    ' This is only required in the startup form!
    Call SetIcon(Me.hwnd, "APPICON", True)
    
    ' This icon is the form icon
    ' Note: Add this line to all forms that you want to display Alpha Icons on, remember
    '       you can change the FORMICON to be any name & icon you want in the AlphaIcon.rc file.
    Call SetIcon(Me.hwnd, "FORMICON", False)
End Sub
