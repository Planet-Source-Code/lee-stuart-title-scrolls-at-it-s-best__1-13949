VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Is Packaged With The Form Title Scroll
'It Has Less Commentation To Let The User
'Figure Out How It Was Done After
'Reading The Form Title Scroll
'Commentation
Dim title As String
Dim i, i2 As Integer
Private Sub Command1_Click()
With Command1
If .Caption = "&Start" Then
Timer1.Enabled = True
.Caption = "&Stop"
Else
.Caption = "&Stop"
Timer1.Enabled = False
.Caption = "&Start"
End If
End With
End Sub
Private Sub Form_Load()
title = " BACKWARDS AND FORWARDS"
i = 1
i2 = 0
End Sub
Private Sub Timer1_Timer()
If i2 = 1 Then
i = i + 1
Form1.Caption = Right(title, i) 'Try Combinations Of
'Form1.Caption = Left(title, i) 'One Of These Two...
Else
i = i - 1
Form1.Caption = Left(title, i)  '...And
'Form1.Caption = Right(title, i)'One Of These Two
End If
If i = Len(title) Then i2 = 0
If i = 0 Then i2 = 1
End Sub
