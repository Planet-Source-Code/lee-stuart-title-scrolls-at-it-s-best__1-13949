VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   -15
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   ScaleHeight     =   -15
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuFast 
         Caption         =   "Fast"
      End
      Begin VB.Menu MnuNormal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuSlow 
         Caption         =   "Slow"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MnuFast_Click()
MnuFast.Checked = True      'Check Please!
MnuNormal.Checked = False   'No check
MnuSlow.Checked = False     'No Check
Form1.Timer1.Interval = 50  '50 Millisecond Scroll
End Sub
Private Sub MnuNormal_Click()
MnuFast.Checked = False     'You Should Get The Idea...
MnuNormal.Checked = True
MnuSlow.Checked = False
Form1.Timer1.Interval = 150
End Sub
Private Sub MnuSlow_Click()
MnuFast.Checked = False
MnuNormal.Checked = False
MnuSlow.Checked = True
Form1.Timer1.Interval = 300
End Sub
