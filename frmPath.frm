VERSION 5.00
Begin VB.Form frmPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirList v1.23"
   ClientHeight    =   3840
   ClientLeft      =   6120
   ClientTop       =   3540
   ClientWidth     =   2970
   Icon            =   "frmPath.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2970
   Begin VB.CommandButton Command2 
      Height          =   285
      Left            =   1125
      Picture         =   "frmPath.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3510
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Height          =   285
      Left            =   2115
      Picture         =   "frmPath.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3510
      Width           =   780
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   180
      Width           =   2700
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   135
      TabIndex        =   0
      Top             =   540
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   2850
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    frmPath.Hide
    frmDirList123.Text2.Text = Dir1.Path

End Sub

Private Sub Command2_Click()
    
    frmPath.Hide

End Sub

Private Sub Drive1_Change()
    
    Dir1.Path = Drive1.Drive

End Sub

