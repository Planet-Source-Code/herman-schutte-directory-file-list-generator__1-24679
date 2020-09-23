VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirList v1.23"
   ClientHeight    =   2175
   ClientLeft      =   6120
   ClientTop       =   4680
   ClientWidth     =   3225
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3225
   Begin VB.Frame Frame1 
      Height          =   2130
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   3120
      Begin VB.Shape Shape1 
         Height          =   600
         Left            =   495
         Top             =   225
         Width           =   2130
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Created by Herman Schutte"
         Height          =   285
         Left            =   45
         TabIndex        =   6
         Top             =   585
         Width           =   3030
      End
      Begin VB.Label Label5 
         Caption         =   "herman@qmuzik.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   765
         MouseIcon       =   "frmAbout.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1845
         Width           =   1590
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Send any suggestions or comments to:"
         Height          =   240
         Left            =   45
         TabIndex        =   4
         Top             =   1620
         Width           =   3000
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Download the source code for this program at:"
         Height          =   465
         Left            =   45
         TabIndex        =   3
         Top             =   900
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "www.geocities.com\open_source_2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   135
         MouseIcon       =   "frmAbout.frx":130C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1350
         Width           =   2850
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "DirList v1.23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   45
         TabIndex        =   1
         Top             =   225
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()

Dim lReturn As Long

On Error GoTo ErrHandler

'API call to open Internet Explorer

    lReturn = ShellExecute(hwnd, _
                           "open", _
                           "http://www.geocities.com\open_source_2001", _
                           vbNull, vbNull, _
                           SW_SHOWNORMAL)
Exit Sub

ErrHandler:
    
    MsgBox ("Please check your internet connection and try again")

End Sub

Private Sub Label5_Click()

Dim lReturn As Long

On Error GoTo ErrHandler

'API call to open Internet Explorer

    lReturn = ShellExecute(hwnd, _
                           "open", _
                           "mailto:herman@qmuzik.com", _
                           vbNull, vbNull, _
                           SW_SHOWNORMAL)

Exit Sub

ErrHandler:
    
    MsgBox ("Please check your internet connection and try again")

End Sub
