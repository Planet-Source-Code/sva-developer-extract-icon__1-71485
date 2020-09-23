VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Icon Extractor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3480
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Browse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox File 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Extract 
      Caption         =   "Extract"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.PictureBox Picdraw 
      Height          =   615
      Left            =   1920
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ICON EXTRACTOR Created By SVA DEVELOPER
'This Code is Used to Extract Icon From Any File
'Don't Just Copy And Paste The Code Try To Write It Yourself
' To Understand It






'For This Code We Need 2 API

'To extract icon
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long

'And to draw the extracted icon
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Sub Browse_Click()
'For error handling
On Error Resume Next
'To Browse the file
cd1.Filter = "All Files(*.*)|*.*"
cd1.ShowOpen
File.Text = cd1.FileName
End Sub

Private Sub Extract_Click()
Dim ico As Long
'To Clear the picturebox before drawing icon
Picdraw.Cls
'This will extract the icon and save it to ico variable
ico = ExtractAssociatedIcon(App.hInstance, File.Text, 0)
'To draw icon in the picturebox
DrawIcon Picdraw.hdc, 0, 0, ico
DestroyIcon ico 'to destroy icon
End Sub
