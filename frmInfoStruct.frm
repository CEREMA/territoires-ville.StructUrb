VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInfoStruct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informations sur "
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmInfoStruct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextInfo 
      Height          =   3000
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   5292
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmInfoStruct.frx":030A
   End
   Begin VB.CommandButton BtnFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer les informations"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   6375
   End
End
Attribute VB_Name = "frmInfoStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnFermer_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.RichTextInfo.BackColor = vbInfoBackground
End Sub
