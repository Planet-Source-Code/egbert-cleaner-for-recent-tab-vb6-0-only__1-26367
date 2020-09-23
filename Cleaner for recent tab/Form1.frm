VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recent cleaner"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Note :"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6015
      Begin VB.Label Label3 
         Caption         =   "Functions above only work if vb is inactive."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clea&r recent"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clea&n recent"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Use this one to clear the hole recent tab."
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Use this command for clean the recent on projects what doesnt exist eny more."
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Ret As VbMsgBoxResult
Ret = MsgBox("Are you sure you want to clean your recent tab?", vbQuestion + vbYesNo, "Confirm")
If Ret = vbYes Then CleanRecent
MsgBox "Completed!", vbInformation + vbOKOnly, "Complete"
End Sub

Private Sub Command2_Click()
Dim Ret As VbMsgBoxResult
Ret = MsgBox("Are you sure you want to clear your recent tab?", vbQuestion + vbYesNo, "Confirm")
If Ret = vbYes Then DeleteRecent
MsgBox "Completed!", vbInformation + vbOKOnly, "Complete"
End Sub

