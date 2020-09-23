VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Right"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Scroll Right"
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Left"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Scroll Left"
      Top             =   120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Scroll"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Start/Stop Scroll"
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Right Click The Form For Speed Options Or Double Click To Change Scroll Text"
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'We'll Use These Later
Dim ScrollText, Scroll1, Scroll2 As String
Dim i As Integer
Private Sub Form_Load()
 'What Do We Want The Scroll To Say?
ScrollText = "Scrolling Title! ******* "
'Make "i" One Less Than The Length Of What We Want Scrolling
i = Len(ScrollText) - 1
'Set The Form Caption
Form1.Caption = ScrollText
End Sub
Private Sub Command1_Click()
'Just So We Can Use One Comand Button
If Command1.Caption = "&Start Scroll" Then
'Make It Scroll
Timer1.Enabled = True
Command1.Caption = "&Stop Scroll"
Else
'Now Stop The Bugger
Timer1.Enabled = False
Command1.Caption = "&Start Scroll"
End If
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Right Click...
If Button = 2 Then
'...Then Show The Popup Menu From Form2
Form1.PopupMenu Form2.MnuFile, 0, X, Y
End If
End Sub
Private Sub Label1_DblClick()
'Use An InputBox Upon Double Click
ScrollText = InputBox("Change Scrolling Text:", "Change Scroll")
'If The Dumbass Clicks Cancel, Or Types Nothing...
'...Give It Something To Scroll
If ScrollText = "" Then ScrollText = "Scrolling Title! ******* "
'Reset i Otherwise An Error Occurs
i = Len(ScrollText) - 1
End Sub
Private Sub Timer1_Timer()
'If Left Scroll Is Selected
If Option1.Value = True Then
'Get The Character On The Left
Scroll1 = Left(ScrollText, 1)
'Then The Rest
Scroll2 = Right(ScrollText, i)
'Put That Character On The End
ScrollText = Scroll2 & Scroll1
'Otherwise
Else
'Get The Character On The Right
Scroll1 = Right(ScrollText, 1)
'And The Rest
Scroll2 = Left(ScrollText, i)
'Put That Character On The Beginning
ScrollText = Scroll1 & Scroll2
End If
'Now Set It As The Caption
Form1.Caption = ScrollText
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Get Rid Of Form1 From Memory
Unload Me
'The Popup Menu From Form2 Was Called, So Close Form2 Also
Unload Form2
'Bye bye!
End Sub
