VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Mencetak Sebagian Rich TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coding berikut akan mencetak isi teks berformat pada 'RichTextBox, tapi tidak akan mencetak gambar apapun 'yang ada di dalam RichTextBox.

Private Sub Command1_Click()
    Call RichTextBox1.SelPrint(Printer.hDC)
End Sub

