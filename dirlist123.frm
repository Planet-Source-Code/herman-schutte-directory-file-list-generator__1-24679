VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDirList123 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirList  v1.23"
   ClientHeight    =   4605
   ClientLeft      =   4215
   ClientTop       =   3165
   ClientWidth     =   7005
   Icon            =   "dirlist123.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7005
   Begin VB.CheckBox Check1 
      Caption         =   "Include Extension"
      Height          =   225
      Left            =   135
      TabIndex        =   3
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Height          =   285
      Left            =   5940
      Picture         =   "dirlist123.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4185
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Height          =   285
      Left            =   4950
      Picture         =   "dirlist123.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4185
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   270
      Left            =   5580
      TabIndex        =   2
      Top             =   3240
      Width           =   345
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   540
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   3240
      Width           =   4995
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   45
      TabIndex        =   12
      Top             =   3060
      Width           =   6900
      Begin VB.CommandButton Command6 
         Height          =   285
         Left            =   90
         Picture         =   "dirlist123.frx":205E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1125
         Width           =   780
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   3060
         Picture         =   "dirlist123.frx":2928
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1125
         Width           =   780
      End
      Begin VB.CommandButton Command4 
         Height          =   285
         Left            =   3960
         Picture         =   "dirlist123.frx":31F2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1125
         Width           =   780
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Include Full Path"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   810
         Width           =   1545
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6255
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "DirList v1.23"
         FileName        =   "*.txt"
         Filter          =   "*.txt"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag n Drop Enabled!"
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   1125
         TabIndex        =   16
         Top             =   1170
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Path:"
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   435
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2685
      Left            =   135
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   225
      Width           =   6705
   End
   Begin VB.Frame fmFiles 
      Caption         =   "List of Files in Folder:"
      Height          =   3030
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   6900
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "No Folder Selected"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1620
         TabIndex        =   14
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   1350
      Visible         =   0   'False
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   825
      Left            =   2385
      TabIndex        =   15
      Top             =   675
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1455
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"dirlist123.frx":3ABC
   End
End
Attribute VB_Name = "frmDirList123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function FullPath(sText As String) As String

'Check to see if a \ is needed at the end of the path or not

Dim sMyFormat As String

    sMyFormat = Text2.Text
    
    If Right(sMyFormat, 1) <> "\" Then _
        sMyFormat = sMyFormat & "\"
        
    If Check2.Value Then
        
        FullPath = sMyFormat & sText
    
    Else
        
        FullPath = sText
    
    End If
    
End Function

Private Sub Command1_Click()
    
    frmPath.Show

End Sub

Private Sub Command2_Click()
    
    Text1.Text = ""
    File1.Path = Text2.Text
    myint = 0
    
    'Check label caption lenght
    
    lblFile.Caption = " " & File1.Path & " "
    If Len(lblFile.Caption) >= 60 Then _
        lblFile.Caption = Left(lblFile.Caption, 60) & "... "
    
    'Populate the text box with files from the hidden file listbox
    
    If File1.ListCount > 0 Then
        
        Do
            
            If Check1.Value Then
                Text1.Text = Text1.Text & FullPath(File1.List(myint)) & vbCrLf
            Else
                Text1.Text = Text1.Text & FullPath(Left(File1.List(myint), (Len(File1.List(myint)) - 4))) & vbCrLf
            End If
                myint = myint + 1
        
        Loop Until myint = File1.ListCount
    
    Else
        
        Text1.Text = ""
    
    End If

End Sub

Private Sub Command3_Click()
    
    End

End Sub

Private Sub Command4_Click()

    RichTextBox1.Text = Text1.Text
    RichTextBox1.SelPrint Printer.hDC
        
End Sub

Private Sub Command5_Click()

Dim cFileName As String

    RichTextBox1.Text = Text1.Text
    CommonDialog1.Action = 2
    cFileName = CommonDialog1.FileName
    
    'Format the save to filename
    
    If InStr(1, cFileName, ".") = 0 Then _
        cFileName = cFileName & ".txt"
    
    If cFileName <> "" And cFileName <> "*.txt" Then _
        RichTextBox1.SaveFile cFileName, rtfText

End Sub

Private Sub Command6_Click()
    
    frmAbout.Show
    
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim txt As String
Dim fname As Variant
Dim MyPos As Integer

    MyPos = 0
    
    'Go through the list of file names in the list
    'of selected files. This will cause just the last
    'selected one to be displayed in the path text
    'box.
    
    For Each fname In Data.Files
        txt = txt & fname & vbCrLf
    Next fname
        
    If InStr(1, txt, vbCrLf) Then
        Text2.Text = Left(txt, Len(txt) - 2)
    Else
        Text2.Text = txt
    End If
    
    Effect = vbDropEffectNone
    
    'Simulate command button press
    Command2_Click

End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim txt As String
Dim fname As Variant
Dim MyPos As Integer

    MyPos = 0
    
    'Go through the list of file names in the list
    'of selected files. This will cause just the last
    'selected one to be displayed in the path text
    'box.
    
    For Each fname In Data.Files
        txt = txt & fname & vbCrLf
    Next fname
        
    If InStr(1, txt, vbCrLf) Then
        Text2.Text = Left(txt, Len(txt) - 2)
    Else
        Text2.Text = txt
    End If
    
    Effect = vbDropEffectNone

End Sub
