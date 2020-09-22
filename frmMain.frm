VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Light Speed HTML Remover"
   ClientHeight    =   1125
   ClientLeft      =   2790
   ClientTop       =   3645
   ClientWidth     =   6165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6165
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Text"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.CheckBox avoidfreezing 
      Caption         =   "Avoid freezing"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label thetime 
      Caption         =   "Select a file to remove it's HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Holds the text filtered from the HTML file
Dim TextFile() As Byte

Private Sub Command1_Click()
Dim FileContent() As Byte
Dim f As Long
With CommonDialog1
.Filter = "All files (*.*)|*.*;"
.ShowOpen
If .FileName <> "" Then
f = FreeFile
Open .FileName For Binary As f
ReDim FileContent(0 To LOF(f) - 1)
Get f, 1, FileContent()
Close f
Dim time1 As Long
Dim time2 As Long
Dim timespent As Double
Dim filesize As Long
thetime.Caption = "Removing HTML from file..."
allowevents = CBool(avoidfreezing.Value)
time1 = GetTickCount
TextFile() = RemoveHTML(FileContent)
time2 = GetTickCount
timespent = Round(((time2 - time1) / 1000), 2)
filesize = Round((FileLen(.FileName) / 1024), 2)
thetime.Caption = "File Size: " & filesize & " KB, HTML Removed in " & timespent & " seconds"
.FileName = ""
End If
End With
End Sub

Private Sub Command2_Click()
Dim f As Long
With CommonDialog1
.Filter = "Text Files (*.txt)|*.txt;"
.ShowSave
If .FileName <> "" Then
f = FreeFile
Open .FileName For Binary As f
Put f, 1, TextFile()
Close f
MsgBox "File was saved succesfully as " & .FileName
.FileName = ""
End If
End With
End Sub

Private Sub Command3_Click()
End
End Sub
