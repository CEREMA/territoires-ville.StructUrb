VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDocument"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "frmDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   11250
   Begin VB.PictureBox PictureCarotte 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   10800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame FrameR�sultat 
      Height          =   6450
      Left            =   60
      TabIndex        =   108
      Top             =   0
      Width           =   4695
      Begin VB.Frame FrameGel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   90
         TabIndex        =   148
         Top             =   5520
         Width           =   2870
         Begin VB.CommandButton BtnOKGel 
            Height          =   460
            Left            =   1200
            Picture         =   "frmDocument.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   164
            Top             =   240
            Visible         =   0   'False
            Width           =   460
         End
         Begin VB.CommandButton BtnGelQ2 
            Height          =   460
            Left            =   3960
            Picture         =   "frmDocument.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   300
            Visible         =   0   'False
            Width           =   460
         End
         Begin VB.CommandButton BtnGelQ1 
            Height          =   460
            Left            =   120
            Picture         =   "frmDocument.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   300
            Visible         =   0   'False
            Width           =   460
         End
         Begin VB.Label LabelIndiceGel 
            AutoSize        =   -1  'True
            Caption         =   "Q1 <---- Indice de Gel en �C.J ----> Q2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   157
            Top             =   0
            Width           =   3240
         End
         Begin VB.Label LabelIGelRef 
            AutoSize        =   -1  'True
            Caption         =   " <---- R�f�rence Corrig� ----> "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   156
            Top             =   280
            Width           =   2490
         End
         Begin VB.Label LabelIGelAdmin 
            AutoSize        =   -1  'True
            Caption         =   " <----       Admissible      ----> "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   155
            Top             =   540
            Width           =   2490
         End
         Begin VB.Label LabelIGRefQ1 
            AutoSize        =   -1  'True
            Caption         =   "(Inconnu)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   154
            Top             =   280
            Width           =   825
         End
         Begin VB.Label LabelIGRefQ2 
            AutoSize        =   -1  'True
            Caption         =   "(Inconnu)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   153
            Top             =   300
            Width           =   825
         End
         Begin VB.Label LabelIGAdmQ1 
            AutoSize        =   -1  'True
            Caption         =   "(Inconnu)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   152
            Top             =   540
            Width           =   825
         End
         Begin VB.Label LabelIGAdmQ2 
            AutoSize        =   -1  'True
            Caption         =   "(Inconnu)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   151
            Top             =   540
            Width           =   825
         End
      End
      Begin VB.Frame FrameCarotte 
         Height          =   4395
         Left            =   90
         TabIndex        =   114
         Top             =   1020
         Width           =   4515
         Begin VB.CommandButton LabelInfoMaxPQ2 
            Height          =   735
            Left            =   3260
            Picture         =   "frmDocument.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   2760
            Width           =   1170
         End
         Begin VB.CommandButton LabelInfoMinTechnoQ2 
            Height          =   450
            Left            =   3260
            Picture         =   "frmDocument.frx":2F02
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   3810
            Width           =   1170
         End
         Begin VB.CommandButton LabelInfoMinTechnoQ1 
            Height          =   450
            Left            =   60
            Picture         =   "frmDocument.frx":4AE0
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   3810
            Width           =   1170
         End
         Begin VB.CommandButton LabelInfoMaxPQ1 
            Height          =   735
            Left            =   60
            Picture         =   "frmDocument.frx":66BE
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   2760
            Width           =   1170
         End
         Begin VB.Label LabelPFQ2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PF?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2055
            TabIndex        =   147
            Top             =   3960
            Width           =   345
         End
         Begin VB.Label LabelPFQ1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PF?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   146
            Top             =   3975
            Width           =   345
         End
         Begin VB.Label LabelBase2Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   1200
            TabIndex        =   145
            Top             =   1440
            Width           =   345
         End
         Begin VB.Label LabelBase2Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   2820
            TabIndex        =   144
            Top             =   1440
            Width           =   345
         End
         Begin VB.Label LabelCBase2Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base2 Q1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   143
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label LabelCBase2Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base1 Q2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   142
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label LabelBase1Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   1200
            TabIndex        =   141
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label LabelBase1Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   2820
            TabIndex        =   140
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label LabelCBase1Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base1 Q1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   139
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label LabelCBase1Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base1 Q2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   138
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label LabelSurf2Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BBSG 5 cm +"
            Height          =   195
            Left            =   660
            TabIndex        =   137
            Top             =   840
            Width           =   960
         End
         Begin VB.Label LabelSurf2Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BBSG 5 cm +"
            Height          =   195
            Left            =   2820
            TabIndex        =   136
            Top             =   840
            Width           =   960
         End
         Begin VB.Label LabelCSurfQ1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CSurface"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   135
            Top             =   720
            Width           =   795
         End
         Begin VB.Label LabelSurf1Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BBSG 5 cm +"
            Height          =   195
            Left            =   660
            TabIndex        =   134
            Top             =   600
            Width           =   960
         End
         Begin VB.Label LabelSurf1Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BBSG 5 cm +"
            Height          =   195
            Left            =   2820
            TabIndex        =   133
            Top             =   600
            Width           =   960
         End
         Begin VB.Label LabelCSurfQ2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CSurface"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   132
            Top             =   720
            Width           =   795
         End
         Begin VB.Label LabelQualit� 
            AutoSize        =   -1  'True
            Caption         =   "Q1 <----- Qualit� -----> Q2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1165
            TabIndex        =   131
            Top             =   0
            Width           =   2145
         End
         Begin VB.Label LabelFond2Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   1200
            TabIndex        =   130
            Top             =   2160
            Width           =   345
         End
         Begin VB.Label LabelFond2Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   2820
            TabIndex        =   129
            Top             =   2160
            Width           =   345
         End
         Begin VB.Label LabelCFond2Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fond2 Q1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   128
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label LabelCFond2Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fond2 Q2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   127
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label LabelFond1Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   1200
            TabIndex        =   126
            Top             =   1800
            Width           =   345
         End
         Begin VB.Label LabelFond1Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8 cm"
            Height          =   195
            Left            =   2820
            TabIndex        =   125
            Top             =   1800
            Width           =   345
         End
         Begin VB.Label LabelCFond1Q1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fond1 Q1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   124
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label LabelCFond1Q2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fond1 Q2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   123
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label LabelEpTotQ1 
            AutoSize        =   -1  'True
            Caption         =   "Tot = 50 cm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   140
            Width           =   1035
         End
         Begin VB.Label LabelEpTotQ2 
            AutoSize        =   -1  'True
            Caption         =   "Tot = 50 cm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3360
            TabIndex        =   121
            Top             =   140
            Width           =   1035
         End
         Begin VB.Label LabelLitPoseQ1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lit de pose"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   120
            Top             =   3120
            Width           =   900
         End
         Begin VB.Label LabelLitPoseQ2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lit de pose"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1755
            TabIndex        =   119
            Top             =   3600
            Width           =   900
         End
         Begin VB.Shape ShapeBase2Q2 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   1320
            Width           =   975
         End
         Begin VB.Shape ShapeFond2Q2 
            FillColor       =   &H00FFFF00&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   2040
            Width           =   975
         End
         Begin VB.Shape ShapeSurfQ2 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   600
            Width           =   975
         End
         Begin VB.Shape ShapeBase1Q2 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   960
            Width           =   975
         End
         Begin VB.Shape ShapeBase2Q1 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   1320
            Width           =   975
         End
         Begin VB.Shape ShapeFond2Q1 
            FillColor       =   &H00FFFF00&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   2040
            Width           =   975
         End
         Begin VB.Shape ShapeFond1Q1 
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   1680
            Width           =   975
         End
         Begin VB.Shape ShapeBase1Q1 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   960
            Width           =   975
         End
         Begin VB.Shape ShapeFond1Q2 
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   1680
            Width           =   975
         End
         Begin VB.Shape ShapeSurfQ1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   1740
            Top             =   600
            Width           =   975
         End
         Begin VB.Shape ShapeLitPoseQ1 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00C0C0C0&
            FillStyle       =   7  'Diagonal Cross
            Height          =   345
            Left            =   1740
            Top             =   3000
            Width           =   975
         End
         Begin VB.Shape ShapeLitPoseQ2 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00C0C0C0&
            FillStyle       =   7  'Diagonal Cross
            Height          =   345
            Left            =   1740
            Top             =   3480
            Width           =   975
         End
         Begin VB.Shape ShapePFQ2 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00C0C0C0&
            FillStyle       =   4  'Upward Diagonal
            Height          =   345
            Left            =   1740
            Top             =   3900
            Width           =   975
         End
         Begin VB.Shape ShapePFQ1 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00C0C0C0&
            FillStyle       =   4  'Upward Diagonal
            Height          =   345
            Left            =   1740
            Top             =   3900
            Width           =   975
         End
      End
      Begin VB.Frame FrameChoixQ1Q2 
         Caption         =   "Chantier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3005
         TabIndex        =   109
         Top             =   5520
         Width           =   1600
         Begin VB.CommandButton BtnPlusInfo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   720
            Picture         =   "frmDocument.frx":8728
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   600
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.OptionButton OptionChoixQ1 
            Caption         =   "Standard (Q1)"
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   280
            Width           =   1320
         End
         Begin VB.OptionButton OptionChoixQ2 
            Caption         =   "Difficile (Q2)"
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label LabelInfo1 
            AutoSize        =   -1  'True
            Caption         =   "Info 1"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   840
            TabIndex        =   113
            Top             =   120
            Width           =   405
         End
         Begin VB.Label LabelInfo2 
            AutoSize        =   -1  'True
            Caption         =   "Info 2"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   840
            TabIndex        =   112
            Top             =   480
            Width           =   405
         End
      End
      Begin VB.Label LabelTypeVoie 
         AutoSize        =   -1  'True
         Caption         =   "Type de voie : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   161
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label LabelTraficCum 
         AutoSize        =   -1  'True
         Caption         =   "Trafic Cumul� : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   160
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label LabelCAM 
         AutoSize        =   -1  'True
         Caption         =   "CAM : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   159
         Top             =   450
         Width           =   585
      End
      Begin VB.Label LabelNEequiv 
         AutoSize        =   -1  'True
         Caption         =   "Nombre d'Essieux Equivalents : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   158
         Top             =   735
         Width           =   2745
      End
   End
   Begin RichTextLib.RichTextBox RichTextAide 
      Height          =   3000
      Left            =   4800
      TabIndex        =   66
      Top             =   3480
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   5292
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDocument.frx":89DA
   End
   Begin TabDlg.SSTab TabData 
      Height          =   3300
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5821
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Voie"
      TabPicture(0)   =   "frmDocument.frx":8A8A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelDate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameTypeEtude"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameVoie"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextTitre"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TextVar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameGiratoire"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Trafic"
      TabPicture(1)   =   "frmDocument.frx":8AA6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OptionCA(5)"
      Tab(1).Control(1)=   "OptionCA(4)"
      Tab(1).Control(2)=   "OptionCA(3)"
      Tab(1).Control(3)=   "OptionCA(2)"
      Tab(1).Control(4)=   "OptionCA(1)"
      Tab(1).Control(5)=   "OptionCA(0)"
      Tab(1).Control(6)=   "TextTrafCUM"
      Tab(1).Control(7)=   "TextTrafIni"
      Tab(1).Control(8)=   "TextDur�eS"
      Tab(1).Control(9)=   "LabelTrafCUM"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "LabelTrafIni"
      Tab(1).Control(12)=   "LabelCroissA"
      Tab(1).Control(13)=   "Label5"
      Tab(1).Control(14)=   "LabelDur�eS"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Plateforme"
      TabPicture(2)   =   "frmDocument.frx":8AC2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameChoixPF"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Structure"
      TabPicture(3)   =   "frmDocument.frx":8ADE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LabelFichPerso"
      Tab(3).Control(1)=   "CheckFichPerso"
      Tab(3).Control(2)=   "CmdInfoMB"
      Tab(3).Control(3)=   "CmdInfoMF"
      Tab(3).Control(4)=   "cmdInfoMS"
      Tab(3).Control(5)=   "FrameTypeStruct"
      Tab(3).Control(6)=   "FrameChoixStruct"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "CAM"
      TabPicture(4)   =   "frmDocument.frx":8AFA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LabelCoefCAM"
      Tab(4).Control(1)=   "MaskCAM"
      Tab(4).Control(2)=   "FrameInfoValCAM"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Couche de surface"
      TabPicture(5)   =   "frmDocument.frx":8B16
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ComboCompQ2"
      Tab(5).Control(1)=   "ListViewMat"
      Tab(5).Control(2)=   "ComboCompQ1"
      Tab(5).Control(3)=   "LabelCompChoisieQ2"
      Tab(5).Control(4)=   "LabelEpPrecQ2"
      Tab(5).Control(5)=   "LabelValEpPrecQ2"
      Tab(5).Control(6)=   "LabelQ2cm"
      Tab(5).Control(7)=   "LabelInfo"
      Tab(5).Control(8)=   "LabelListMatSurfComp"
      Tab(5).Control(9)=   "ImageListMat"
      Tab(5).Control(10)=   "LabelQ1cm"
      Tab(5).Control(11)=   "LabelValEpPrecQ1"
      Tab(5).Control(12)=   "LabelEpPrecQ1"
      Tab(5).Control(13)=   "LabelCompChoisieQ1"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "Gel"
      TabPicture(6)   =   "frmDocument.frx":8B32
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "FrameStationGel"
      Tab(6).Control(1)=   "FrameHiver"
      Tab(6).Control(2)=   "FrameSolSup"
      Tab(6).Control(3)=   "FrameCFgel"
      Tab(6).ControlCount=   4
      Begin VB.Frame FrameChoixStruct 
         Caption         =   "Structure choisie :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -72980
         TabIndex        =   30
         Top             =   600
         Width           =   4200
         Begin VB.CommandButton BtnInfoStruct 
            Caption         =   "Informations sur la structure choisie"
            Enabled         =   0   'False
            Height          =   735
            Left            =   2400
            TabIndex        =   33
            Top             =   240
            Width           =   1635
         End
         Begin VB.ComboBox ComboStruct 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Tag             =   "-1"
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label LabelRisk 
            AutoSize        =   -1  'True
            Caption         =   "Risque de calcul :"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   780
            Visible         =   0   'False
            Width           =   1275
         End
      End
      Begin VB.Frame FrameTypeStruct 
         Caption         =   "Type de structure :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74940
         TabIndex        =   22
         Top             =   600
         Width           =   1900
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "Bitumineuse"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "Pav�e ou Dall�e"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   1575
         End
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "Mixte"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   855
         End
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "B�ton"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   855
         End
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "GTLH"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "Souple"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   23
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton OptionTypeStruct 
            Caption         =   "Souple"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame FrameGiratoire 
         Height          =   2415
         Left            =   4440
         TabIndex        =   9
         Top             =   720
         Width           =   2895
         Begin VB.OptionButton OptionGirDis 
            Caption         =   "Giratoire voie de distribution"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   320
            Width           =   2535
         End
         Begin VB.OptionButton OptionGirPL 
            Caption         =   "Giratoire voie principale avec PL"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   745
            Width           =   2655
         End
      End
      Begin VB.OptionButton OptionCA 
         Caption         =   "5%"
         Height          =   255
         Index           =   5
         Left            =   -69310
         TabIndex        =   20
         Top             =   1480
         Width           =   550
      End
      Begin VB.Frame FrameCFgel 
         Caption         =   "Couche de forme non g�live : "
         Height          =   1120
         Left            =   -72960
         TabIndex        =   62
         Top             =   2075
         Width           =   4215
         Begin VB.TextBox TextEpaisseur 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   65
            Top             =   720
            Width           =   320
         End
         Begin VB.OptionButton OptionAT 
            Caption         =   "0,14 (Trait�)"
            Height          =   195
            Left            =   2940
            TabIndex        =   64
            Top             =   360
            Width           =   1180
         End
         Begin VB.OptionButton OptionANT 
            Caption         =   "0,12 (Non Trait�)"
            Height          =   195
            Left            =   1200
            TabIndex        =   63
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "cm"
            Height          =   195
            Left            =   1600
            TabIndex        =   100
            Top             =   780
            Width           =   210
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Coefficient A :"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Epaisseur : "
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   780
            Width           =   825
         End
      End
      Begin VB.Frame FrameSolSup 
         Caption         =   "Sol Support"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   57
         Top             =   1740
         Width           =   1815
         Begin VB.TextBox TextPente 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   840
            MaxLength       =   7
            TabIndex        =   61
            Top             =   1050
            Width           =   825
         End
         Begin VB.OptionButton OptionNGel 
            Caption         =   "Non g�lif"
            Height          =   195
            Left            =   60
            TabIndex        =   60
            Top             =   780
            Width           =   1095
         End
         Begin VB.OptionButton OptionPGel 
            Caption         =   "Peu g�lif"
            Height          =   195
            Left            =   60
            TabIndex        =   59
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton OptionTGel 
            Caption         =   "Tr�s g�lif"
            Height          =   195
            Left            =   60
            TabIndex        =   58
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Pente : "
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   555
         End
      End
      Begin VB.Frame FrameHiver 
         Caption         =   "Hiver de R�f�rence"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   47
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton OptionHC 
            Caption         =   "HC (Courant)"
            Height          =   195
            Left            =   60
            TabIndex        =   50
            Top             =   900
            Width           =   1575
         End
         Begin VB.OptionButton OptionHRNE 
            Caption         =   "HRNE (Rigoureux)"
            Height          =   195
            Left            =   60
            TabIndex        =   49
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton OptionHE 
            Caption         =   "HE (Exceptionnel)"
            Height          =   195
            Left            =   60
            TabIndex        =   48
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FrameStationGel 
         Height          =   1575
         Left            =   -72960
         TabIndex        =   51
         Top             =   380
         Width           =   4215
         Begin VB.TextBox TextIndGelPerso 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   55
            Top             =   690
            Width           =   495
         End
         Begin VB.CheckBox CheckIndGelPerso 
            Alignment       =   1  'Right Justify
            Caption         =   "Indice de gel personnel"
            Height          =   495
            Left            =   2730
            TabIndex        =   101
            Top             =   600
            Value           =   2  'Grayed
            Width           =   1335
         End
         Begin VB.TextBox TextHAgglo 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   54
            Top             =   840
            Width           =   495
         End
         Begin VB.ComboBox ComboStation 
            Height          =   315
            ItemData        =   "frmDocument.frx":8B4E
            Left            =   1440
            List            =   "frmDocument.frx":8B50
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   180
            Width           =   2655
         End
         Begin VB.ComboBox ComboTailleAgglo 
            Height          =   315
            ItemData        =   "frmDocument.frx":8B52
            Left            =   1440
            List            =   "frmDocument.frx":8B5F
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1170
            Width           =   2655
         End
         Begin VB.Label LabelIndGelPerso 
            AutoSize        =   -1  'True
            Caption         =   "Indice de gel : "
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label LabelCJ 
            AutoSize        =   -1  'True
            Caption         =   "�C.J"
            Height          =   195
            Left            =   1800
            TabIndex        =   102
            Top             =   720
            Width           =   285
         End
         Begin VB.Label LabAltUnit 
            AutoSize        =   -1  'True
            Caption         =   "m�tres"
            Height          =   195
            Left            =   2040
            TabIndex        =   96
            Top             =   840
            Width           =   465
         End
         Begin VB.Label LabAgglo 
            AutoSize        =   -1  'True
            Caption         =   "Taille Agglo : "
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   1200
            Width           =   960
         End
         Begin VB.Label LabelHStation 
            AutoSize        =   -1  'True
            Caption         =   "H Station : "
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   540
            Width           =   795
         End
         Begin VB.Label LabAlt 
            AutoSize        =   -1  'True
            Caption         =   "H Agglom�ration : "
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label LabGel 
            AutoSize        =   -1  'True
            Caption         =   "Station R�f : "
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   930
         End
         Begin VB.Label LabelHStatUnit 
            AutoSize        =   -1  'True
            Caption         =   "m�tres"
            Height          =   195
            Left            =   2040
            TabIndex        =   91
            Top             =   540
            Width           =   465
         End
         Begin VB.Label HStation 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1440
            TabIndex        =   53
            Top             =   540
            Width           =   495
         End
      End
      Begin VB.Frame FrameChoixPF 
         Caption         =   "Choix de la Classe de Plateforme : "
         Height          =   2535
         Left            =   -74760
         TabIndex        =   40
         Top             =   600
         Width           =   5895
         Begin VB.OptionButton OptionPF2Plus 
            Caption         =   "PF2+        (EV2 > 80 MPa)"
            Height          =   255
            Left            =   360
            TabIndex        =   104
            Top             =   1440
            Width           =   2415
         End
         Begin VB.OptionButton OptionPF3 
            Caption         =   "PF3          (EV2 > 120 MPa)"
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   1920
            Width           =   2415
         End
         Begin VB.OptionButton OptionPF2 
            Caption         =   "PF2          (EV2 > 50 MPa)"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   960
            Width           =   2415
         End
         Begin VB.OptionButton OptionPF1 
            Caption         =   "PF1          (EV2 > 20 MPa)"
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Frame FrameInfoValCAM 
         Caption         =   "Pour le type de voie : "
         Height          =   1695
         Left            =   -74880
         TabIndex        =   87
         Top             =   1440
         Width           =   6135
         Begin VB.Label LabelValMaxCAM 
            AutoSize        =   -1  'True
            Caption         =   "Valeur maximale du CAM     = "
            Height          =   195
            Left            =   720
            TabIndex        =   90
            Top             =   1260
            Width           =   2115
         End
         Begin VB.Label LabelValMinCAM 
            AutoSize        =   -1  'True
            Caption         =   "Valeur minimale du CAM      = "
            Height          =   195
            Left            =   720
            TabIndex        =   89
            Top             =   840
            Width           =   2115
         End
         Begin VB.Label LabelValPrecCAM 
            AutoSize        =   -1  'True
            Caption         =   "Valeur pr�conis�e du CAM  = "
            Height          =   195
            Left            =   720
            TabIndex        =   88
            Top             =   420
            Width           =   2115
         End
      End
      Begin MSMask.MaskEdBox MaskCAM 
         Height          =   300
         Left            =   -72360
         TabIndex        =   39
         Top             =   795
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin VB.OptionButton OptionCA 
         Caption         =   "4%"
         Height          =   255
         Index           =   4
         Left            =   -69940
         TabIndex        =   19
         Top             =   1480
         Width           =   615
      End
      Begin VB.OptionButton OptionCA 
         Caption         =   "3%"
         Height          =   255
         Index           =   3
         Left            =   -70590
         TabIndex        =   18
         Top             =   1480
         Width           =   615
      End
      Begin VB.OptionButton OptionCA 
         Caption         =   "2%"
         Height          =   255
         Index           =   2
         Left            =   -71220
         TabIndex        =   17
         Top             =   1480
         Width           =   615
      End
      Begin VB.OptionButton OptionCA 
         Caption         =   "1%"
         Height          =   255
         Index           =   1
         Left            =   -71850
         TabIndex        =   16
         Top             =   1480
         Width           =   615
      End
      Begin VB.OptionButton OptionCA 
         Caption         =   "0%"
         Height          =   255
         Index           =   0
         Left            =   -72480
         TabIndex        =   15
         Top             =   1480
         Width           =   615
      End
      Begin VB.CommandButton cmdInfoMS 
         Caption         =   "Informations Mat�riau couche de surface : "
         Height          =   375
         Left            =   -72970
         TabIndex        =   34
         Top             =   1800
         Width           =   4260
      End
      Begin VB.ComboBox ComboCompQ2 
         Height          =   315
         ItemData        =   "frmDocument.frx":8BBC
         Left            =   -71880
         List            =   "frmDocument.frx":8BBE
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1200
         Width           =   3135
      End
      Begin ComctlLib.ListView ListViewMat 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   46
         Top             =   2160
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   1931
         View            =   1
         Arrange         =   2
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageListMat"
         SmallIcons      =   "ImageListMat"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox ComboCompQ1 
         Height          =   315
         ItemData        =   "frmDocument.frx":8BC0
         Left            =   -71880
         List            =   "frmDocument.frx":8BC2
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton CmdInfoMF 
         Caption         =   "Informations Mat�riau de la couche de fondation : "
         Height          =   375
         Left            =   -72970
         TabIndex        =   36
         Top             =   2760
         Width           =   4260
      End
      Begin VB.CommandButton CmdInfoMB 
         Caption         =   "Informations Mat�riau de la couche de base : "
         Height          =   375
         Left            =   -72970
         TabIndex        =   35
         Top             =   2280
         Width           =   4260
      End
      Begin VB.CheckBox CheckFichPerso 
         Caption         =   "Utiliser le fichier de structures personnelles"
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   2580
         Value           =   2  'Grayed
         Width           =   3735
      End
      Begin VB.TextBox TextTrafCUM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72480
         MaxLength       =   11
         TabIndex        =   21
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TextTrafIni 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TextDur�eS 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox TextVar 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   2490
         Width           =   2175
      End
      Begin VB.TextBox TextTitre 
         Height          =   960
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Frame FrameVoie 
         Height          =   2415
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   2895
         Begin VB.OptionButton OptionVoieParking 
            Caption         =   "Parking VL, piste cyclable, ..."
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2020
            Width           =   2655
         End
         Begin VB.OptionButton OptionVoieDes 
            Caption         =   "Voie de desserte"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   320
            Width           =   1935
         End
         Begin VB.OptionButton OptionVoieBus 
            Caption         =   "Voie r�serv�e aux bus"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1595
            Width           =   2655
         End
         Begin VB.OptionButton OptionVoiePL 
            Caption         =   "Voie principale avec PL"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1170
            Width           =   2655
         End
         Begin VB.OptionButton OptionVoieDis 
            Caption         =   "Voie de distribution"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   745
            Width           =   1935
         End
      End
      Begin VB.Frame FrameTypeEtude 
         Caption         =   "Type d'am�nagement :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   105
         Top             =   480
         Width           =   3135
         Begin VB.OptionButton OptionEtudeStandard 
            Caption         =   "Section courante"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   260
            Width           =   1575
         End
         Begin VB.OptionButton OptionEtudeGiratoire 
            Caption         =   "Giratoire"
            Height          =   255
            Left            =   1920
            TabIndex        =   106
            Top             =   260
            Width           =   975
         End
      End
      Begin VB.Label LabelCoefCAM 
         AutoSize        =   -1  'True
         Caption         =   "Coefficient d'Agressivit� Moyen : "
         Height          =   195
         Left            =   -74760
         TabIndex        =   86
         Top             =   840
         Width           =   2340
      End
      Begin VB.Label LabelCompChoisieQ2 
         Caption         =   "Composition choisie Q2 :"
         Height          =   435
         Left            =   -72840
         TabIndex        =   85
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label LabelEpPrecQ2 
         Caption         =   "Epaisseur pr�conis�e Q2 :"
         Height          =   435
         Left            =   -74880
         TabIndex        =   84
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label LabelValEpPrecQ2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   -73680
         TabIndex        =   83
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label LabelQ2cm 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   -73245
         TabIndex        =   82
         Top             =   1275
         Width           =   210
      End
      Begin VB.Label LabelInfo 
         Caption         =   "Cliquez sur un mat�riau pour avoir ses caract�ristiques"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   -71280
         TabIndex        =   81
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label LabelListMatSurfComp 
         Caption         =   "Mat�riaux composants possibles pour cette �paisseur pr�conis�e :"
         Height          =   495
         Left            =   -74880
         TabIndex        =   80
         Top             =   1680
         Width           =   2535
      End
      Begin ComctlLib.ImageList ImageListMat 
         Left            =   -72120
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDocument.frx":8BC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDocument.frx":8EDE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label LabelQ1cm 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   -73250
         TabIndex        =   79
         Top             =   680
         Width           =   210
      End
      Begin VB.Label LabelValEpPrecQ1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   -73680
         TabIndex        =   78
         Top             =   645
         Width           =   360
      End
      Begin VB.Label LabelEpPrecQ1 
         Caption         =   "Epaisseur pr�conis�e Q1 :"
         Height          =   435
         Left            =   -74880
         TabIndex        =   77
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label LabelCompChoisieQ1 
         Caption         =   "Composition choisie Q1 :"
         Height          =   435
         Left            =   -72840
         TabIndex        =   76
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label LabelFichPerso 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   -74760
         TabIndex        =   38
         Top             =   2880
         Width           =   5895
      End
      Begin VB.Label LabelTrafCUM 
         Caption         =   "Trafic cumul� de PL sur la dur�e de service :"
         Height          =   495
         Left            =   -74880
         TabIndex        =   75
         Top             =   2160
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "par voie, par sens et par jour"
         Height          =   195
         Left            =   -71760
         TabIndex        =   74
         Top             =   600
         Width           =   2010
      End
      Begin VB.Label LabelDate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25/01/2000"
         Height          =   255
         Left            =   2205
         TabIndex        =   73
         Top             =   2880
         Width           =   1050
      End
      Begin VB.Label LabelTrafIni 
         AutoSize        =   -1  'True
         Caption         =   "Trafic initial � la mise en service :"
         Height          =   195
         Left            =   -74880
         TabIndex        =   72
         Top             =   600
         Width           =   2325
      End
      Begin VB.Label LabelCroissA 
         AutoSize        =   -1  'True
         Caption         =   "Croissance annuelle :"
         Height          =   195
         Left            =   -74070
         TabIndex        =   71
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ans"
         Height          =   195
         Left            =   -71760
         TabIndex        =   70
         Top             =   1020
         Width           =   255
      End
      Begin VB.Label LabelDur�eS 
         AutoSize        =   -1  'True
         Caption         =   "Dur�e de service :"
         Height          =   195
         Left            =   -73860
         TabIndex        =   69
         Top             =   1020
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Variante : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titre de l'�tude :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Derni�re modification : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant l'�paisseur totale maximale entre les carottes Q1 et Q2
Private monEpTotMaxReel As Integer
    
'Variable indiquant si toutes les saisies sont finies
'ou valides dans cette fen�tre
Public maFinSaisie As Byte 'valeur initiale = 0

'Variable indiquant si une erreur s'est produit
'lors d'un lost focus dans l'une des TextBox de l'onglet Trafic
'Initialis� � FAUX
Private maTrafIniErreur As Boolean
Private maDurServErreur As Boolean
Private maTrafCumErreur As Boolean

'Variable indiquant si on veut juste d'ouvrir l'�tude
Public monJustOpen As Boolean 'Initialis� � False par VB

'Variable indiquant si un change event de textbox
'(traf ini, traf cum ou dur�ee service) est survenue
'Initialisation par d�faut FALSE
Public monChangeEvent As Boolean

'Variable indiquant s'il faut afficher le message d'erreur
'lors du calcul du NE lorqu'une erreur se produit
Public monAffichageErreurNE As Boolean

'Variables stockant les champs d'une �tude
'Initialement elles valent toutes 0
Public monTypeEtude As Byte
Public monTypeChantier As Byte
Public monTypeVoie As Integer
Public maDate As String
Public monTitreEtude As String
Public maVariante As String
Public monTraficIni As Integer
Public monTraficCumul� As Long
Public maDur�eService As Byte
Public maCroisAnnuel As Byte
Public monCAM As String
Public monNEEquiv As Long
Public monNEth As Long
Public monIndicePF As Byte
Public monIndiceGelRefQ1 As Integer
Public monIndiceGelAdmQ1 As Long
Public monIndiceGelRefQ2 As Integer
Public monIndiceGelAdmQ2 As Long
Public monIndStructChoisie As Integer
Public monFichPersoSTR As String
Public monUtilFichPerso As Byte '=0 pas d'utilisation fichier perso
                                '=1 utiisation d'un fichier perso
Public monIndGelPerso As Integer
Public monUtilIndGelPerso As Byte '=0 pas d'utilisation d'un indice de gel perso
                                  '=1 utiisation d'un indice de gel perso

Public monFichId As Integer 'Id du fichier de stockage .URB

'Variable indiquant si on a trouv� les �paisseurs des qualit�s Q1 et Q2
Public monEpQ1Trouv As Boolean
Public monEpQ2Trouv As Boolean

'Variable indiquant les Qm des qualit�s Q1 et Q2
Public monQmQ1 As Single
Public monQmQ2 As Single

'Variable indiquant si les min techno et max pratiques sont atteints
'pour les qualit�s Q1 et Q2 (valeurs possibles "Oui" ou "Non")
Public monMinTecQ1 As String
Public monMaxPraQ1 As String
Public monMinTecQ2 As String
Public monMaxPraQ2 As String
    
'Tableau contenant les �paisseurs des couches de la structure choisie
'De 1 � 6 ==> Epaisseurs pour la qualit� Q1
'1 et 2 �paisseur couche surface, 3 et 4 �paisseur couche base, 5 et 6 �paisseur couche fondation
'De 7 � 12 ==> Epaisseurs pour la qualit� Q2
'7 et 8 �paisseur couche surface, 9 et 10 �paisseur couche base, 11 et 12 �paisseur couche fondation
Public monTabEp As Variant

'Variable contenant les mat�riaux composants
'du mat�riau de surface compos� (deux maximuns) pour qualit� Q1 et Q2
Public monMSComp1Q1 As String
Public monMSComp2Q1 As String
Public monMSComp1Q2 As String
Public monMSComp2Q2 As String
'Variables de l'�paisseur pr�conis� trouv� pour Q1 et Q2
Public monEpPrecQ1 As Integer
Public monEpPrecQ2 As Integer
'Variables contenant l'indice de la composition choisie
'pour le mat�riau de surface compos� pour Q1 et Q2
Public monIndCompQ1 As Integer
Public monIndCompQ2 As Integer

'Variables concernant la v�rif au gel
Public monIndHiver As Integer
Public monIndStation As Integer
Public monHAgglo As Integer
Public monIndTailleAgglo As Integer
Public monCoefA As Single
Public monEpNonGel As Integer
Public maPente As String
Public monIndGelSol As Integer

'Variables indiquant une modif dans l'�tude
Public maModif As Boolean

'Variables pour indiquer le premier activate de la fen�tre Etude = frmDocument
Private unNbActivate As Byte 'valeur par d�faut = 0
    

Private Sub BtnGelQ1_Click()
    MsgBox "Cette chauss�e n'est pas prot�g�e au gel en condition de chantier standard (qualit� Q1)," + Chr(13) + Chr(13) + "car l'Indice de Gel de R�f�rence > Indice de Gel Admissible pour Q1.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
End Sub

Private Sub BtnGelQ2_Click()
    MsgBox "Cette chauss�e n'est pas prot�g�e au gel en condition de chantier difficile (qualit� Q2)," + Chr(13) + Chr(13) + "car l'Indice de Gel de R�f�rence > Indice de Gel Admissible pour Q2.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
End Sub

Private Sub BtnInfoStruct_Click()
    'Affichage de la fiche commentaire de la structure choisie
    Dim uneStruct As Structure
    
    'Recup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Chargement sans affichage
    Load frmInfoStruct
    
    'Remplissage de frmInfoStruct
    frmInfoStruct.Caption = "Informations sur " + uneStruct.monAbr�g�
    frmInfoStruct.RichTextInfo.TextRTF = uneStruct.monComment
    
    'Centrage de la fiche mat�riau
    CentrerFenetreEcran frmInfoStruct
    
    'Affichage modal
    frmInfoStruct.Show vbModal
End Sub

Private Sub BtnOKGel_Click()
    If OptionChoixQ1.Value Then
        MsgBox "Cette chauss�e est prot�g�e au gel en condition de chantier standard (qualit� Q1)," + Chr(13) + Chr(13) + "car l'Indice de Gel de R�f�rence <= Indice de Gel Admissible pour Q1.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
    End If
    If OptionChoixQ2.Value Then
        MsgBox "Cette chauss�e est prot�g�e au gel en condition de chantier difficile (qualit� Q2)," + Chr(13) + Chr(13) + "car l'Indice de Gel de R�f�rence <= Indice de Gel Admissible pour Q2.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
    End If
End Sub

Private Sub BtnPlusInfo_Click()
    MsgBox MsgMoreInfoForGiratoire, vbInformation
    RichTextAide.SetFocus
End Sub

Private Sub CheckFichPerso_Click()
    AfficherFichierPerso Me
End Sub

Private Sub CheckIndGelPerso_Click()
    Dim uneVisibilit� As Boolean
    
    maModif = True 'Indication de changement pour la sauvegarde
    
    uneVisibilit� = (CheckIndGelPerso.Value = 1)
    monUtilIndGelPerso = CheckIndGelPerso.Value
    
    'Cas o� l'on coche la case c'est vrai
    'Affichage si vrai ou masquage si faux
    'de la zone de saisie de l'indice de gel personnel
    LabelIndGelPerso.Visible = uneVisibilit�
    TextIndGelPerso.Visible = uneVisibilit�
    LabelCJ.Visible = uneVisibilit�
        
    'Mise � jour des valeurs d'indices de gel de r�f�rence
    If uneVisibilit� Then
        TextIndGelPerso.Text = Format(monIndGelPerso)
        TextIndGelPerso.ForeColor = QBColor(0)
        Me.monIndiceGelRefQ1 = monIndGelPerso
        Me.monIndiceGelRefQ2 = monIndGelPerso
    Else
        'Recalcul et affichage des indices de gel de r�f�rences
        'Q1 et Q2 � partir de la station de r�f�rence
        CalculerIndiceGelRef Me
    End If
    
    'Mise en gris�e si vrai ou d�gris�e si faux
    'de la frame de choix des hivers de r�f�rene
    FrameHiver.Enabled = Not uneVisibilit�
    OptionHE.Enabled = Not uneVisibilit�
    OptionHC.Enabled = Not uneVisibilit�
    OptionHRNE.Enabled = Not uneVisibilit�
    
    'Affichage si vrai, masquage si faux
    'de la zone de saisie de l'indice de gel personnel
    LabGel.Visible = Not uneVisibilit�
    ComboStation.Visible = Not uneVisibilit�
    LabelHStation.Visible = Not uneVisibilit�
    HStation.Visible = Not uneVisibilit�
    LabelHStatUnit.Visible = Not uneVisibilit�
    LabAlt.Visible = False 'Not uneVisibilit� (car Hagglo plus utilis�e)
    TextHAgglo.Visible = False 'Not uneVisibilit� (car Hagglo plus utilis�e)
    LabAltUnit.Visible = False 'Not uneVisibilit� (car Hagglo plus utilis�e)
    LabAgglo.Visible = Not uneVisibilit�
    ComboTailleAgglo.Visible = Not uneVisibilit�

    'Affichage dans la frame r�sultat de l'�tude active
    ActualiserFrameVerifGel Me
End Sub

Private Sub CmdInfoMB_Click()
    Dim uneStruct As Structure
    
    'R�cup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Affichage de la fiche caract�ristique du mat�riau
    'Param�tre string = TypeMat + "/" + Abr�g� du Mat�riau
    AfficherFicheMat "FondBase" + "/" + uneStruct.maCoucheBase
End Sub

Private Sub CmdInfoMF_Click()
    Dim uneStruct As Structure
    
    'R�cup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Affichage de la fiche caract�ristique du mat�riau
    'Param�tre string = TypeMat + "/" + Abr�g� du Mat�riau
    AfficherFicheMat "FondBase" + "/" + uneStruct.maCoucheFondation
End Sub

Private Sub cmdInfoMS_Click()
    Dim unTypeMat As String
    Dim uneStruct As Structure
    
    'R�cup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'R�cup du type de mat�riau en couche de surface
    If TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatSimple Then
        unTypeMat = "Simple"
    ElseIf TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatCompos� Then
        unTypeMat = "Compos�"
    Else
        MsgBox MsgErreurProg + MsgErreurMat�riauInconnu + MsgIn + "frmDocument:cmdInfoMS_Click", vbCritical
    End If
        
    'Affichage de la fiche caract�ristique du mat�riau
    'Param�tre string = TypeMat + "/" + Abr�g� du Mat�riau
    AfficherFicheMat unTypeMat + "/" + uneStruct.maCoucheSurface
End Sub

Private Sub ComboCompQ1_Click()
    Dim lesComp As Collection
    'R�cup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    If Not (uneStruct Is Nothing) And ComboCompQ1.ListIndex > -1 Then
        unIndComp = ComboCompQ1.ItemData(ComboCompQ1.ListIndex)
        Set lesComp = maColMatSurf(uneStruct.maCoucheSurface).mesCompositions
        'R�cup des noms de mat�riaux composants et de leur �paisseur
        monTabEp(1) = CInt(lesComp(unIndComp + 1))
        monMSComp1Q1 = Format(lesComp(unIndComp + 2))
        monTabEp(2) = CInt(lesComp(unIndComp + 3))
        monMSComp2Q1 = Format(lesComp(unIndComp + 4))
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
        'Si les �paisseurs pr�conis�es Q1 et Q2 sont les m�mes
        'la valeur mise pour Q1 est affect� aussi � Q2 par d�faut
        If monEpPrecQ1 = monEpPrecQ2 Then ComboCompQ2.ListIndex = ComboCompQ1.ListIndex
        'Calcul de l'indice de gel admissible
        If mesOptionsGen.maVerifGel Then AfficherEtCalculerIndGelAdm Me
    End If
End Sub


Private Sub ComboCompQ2_Click()
    Dim lesComp As Collection
    'R�cup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    If Not (uneStruct Is Nothing) And ComboCompQ2.ListIndex > -1 Then
        unIndComp = ComboCompQ2.ItemData(ComboCompQ2.ListIndex)
        Set lesComp = maColMatSurf(uneStruct.maCoucheSurface).mesCompositions
        'R�cup des noms de mat�riaux composants et de leur �paisseur
        monTabEp(7) = CInt(lesComp(unIndComp + 1))
        monMSComp1Q2 = Format(lesComp(unIndComp + 2))
        monTabEp(8) = CInt(lesComp(unIndComp + 3))
        monMSComp2Q2 = Format(lesComp(unIndComp + 4))
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
        'Calcul de l'indice de gel admissible
        If mesOptionsGen.maVerifGel Then AfficherEtCalculerIndGelAdm Me
    End If
End Sub


Private Sub ComboStation_Click()
    'Affichage de l'altitude de la station de r�f�rence choisie
    HStation.Caption = Format(monTabStation(ComboStation.ListIndex + 1).monAltitude)
    
    AfficherEtCalculerIndGelRef Me
End Sub

Private Sub ComboStruct_Click()
    Dim uneStructChoisie As Boolean
    Dim uneColStruct As Collection
    Dim unMatSurf As Object
    Dim uneStruct As Structure
    Dim unMatBase As String, unMatFond As String
    Dim unTabEp(1 To 12) As Integer
    
    If ComboStruct.Tag = "NoClickEvent" Then
        ComboStruct.Tag = ""
        Exit Sub
    End If
        
    'Modif de structure
    '==> Remise des �paisseurs Q1/Q2 � non trouv�es
    'et des composants 1 et 2 en couche de surface � vide
    If Val(ComboStruct.Tag) <> ComboStruct.ListIndex And ComboStruct.Tag <> "-1" Then
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        monMinTecQ1 = "Non"
        monMaxPraQ1 = "Non"
        monMinTecQ2 = "Non"
        monMaxPraQ2 = "Non"
        monMSComp1Q1 = ""
        monMSComp2Q1 = ""
        monMSComp1Q2 = ""
        monMSComp2Q2 = ""
        monNEEquiv = 0
        'Remise � vide du CAM
        MaskCAM.Mask = ""
        MaskCAM.Text = ""
        LabelCAM.Caption = LabelCAMCaption + MsgInconnu
        LabelNEequiv.Caption = LabelNEequivCaption + MsgInconnu
    End If
    
    If Val(ComboStruct.Tag) = ComboStruct.ListIndex Then
        'Choix par click de la m�me structure que celle
        'pr�c�demment s�lectionn�e
        Exit Sub
    Else
        ComboStruct.Tag = ComboStruct.ListIndex
    End If
            
    uneStructChoisie = (ComboStruct.ListIndex > -1)
    If CheckFichPerso.Value = 0 Or LabelFichPerso.Caption = "" Then
        Set uneColStruct = maColStructCERTU
    Else
        Set uneColStruct = maColStructPerso
    End If
    
    'Mise � jour des boutons de visu des mat�riau surface simple, base
    'et fondation
    cmdInfoMS.Enabled = uneStructChoisie
    cmdInfoMS.Caption = LabelCmdInfoMS
    CmdInfoMB.Enabled = uneStructChoisie
    CmdInfoMB.Caption = LabelCmdInfoMB
    CmdInfoMF.Enabled = uneStructChoisie
    CmdInfoMF.Caption = LabelCmdInfoMF
    
    LabelRisk.Visible = uneStructChoisie
    
    'L'utilisation de l'onglet Couche de surface est possible
    'si une structure est choisie et si la couche de surface
    '�ventuelle est faite d'un mat�riau de surface compos�
    If uneStructChoisie Then
        Set uneStruct = uneColStruct(ComboStruct.ItemData(ComboStruct.ListIndex))
        unMatBase = uneStruct.maCoucheBase
        unMatFond = uneStruct.maCoucheFondation
        CmdInfoMB.Caption = LabelCmdInfoMB + unMatBase
        'Test s'il y a une couche de fondation,
        'par contre il y a toujours une couche de base
        If unMatFond = "Aucune" Then
            CmdInfoMF.Enabled = False
        Else
            CmdInfoMF.Caption = LabelCmdInfoMF + unMatFond
        End If
        'Affichage du risque de calcul si non nul
        If uneStruct.monTauxRisque = 0 Then
            LabelRisk.Visible = False
        Else
            LabelRisk.Caption = "Risque de calcul = " + Format(uneStruct.monTauxRisque) + " %"
        End If
        'Activation de l'onglet CAM
        TabData.TabEnabled(OngletCAM) = uneStructChoisie
        'Affectation de la valeur pr�conis�e du CAM
        ActualiserOngletLabelCAM
        
        'En version 2 si on n'est pas en voie de parking, on grise l'onglet CAM
        'ainsi le CAM a la valeur pr�conis�e poour les voies de parking
        If OptionVoieParking.Value = True Then
            TrouverCaract�reD�cimalUtilis�
            MaskCAM.Mask = "#" + monCarDeci + "##"
            If uneStruct.monTypeStructure = Souple Then
                MaskCAM.Text = Format(0.2, "fixed")
            ElseIf uneStruct.monTypeStructure = PavesDalles And UCase(uneStruct.maCoucheBase) = "GNT" Then
                MaskCAM.Text = Format(0.2, "fixed")
            Else
                MaskCAM.Text = Format(0.1, "fixed")
            End If
            TabData.TabEnabled(OngletCAM) = False
        End If
        
        'Calcul et affichage �ventuel du NE
        CalculerEtAfficherNE
        
        'R�cup du mat�riau de surface �ventuel
        'et test s'il est compos� ==> Activation Onglet Couche surface
        If uneStruct.maCoucheSurface = "Aucune" Then
            TabData.TabEnabled(OngletSurf) = False
            cmdInfoMS.Enabled = False
            'On remet � vide les compositions possibles de la couche de surface
            ComboCompQ1.ListIndex = -1
            ComboCompQ2.ListIndex = -1
        Else
            Set unMatSurf = maColMatSurf(uneStruct.maCoucheSurface)
            TabData.TabEnabled(OngletSurf) = (TypeOf unMatSurf Is MatCompos�) And (monEpQ1Trouv Or monEpQ2Trouv)
            If TabData.TabEnabled(OngletSurf) Then
                MettreAJourOngletCoucheSurface Me, monEpPrecQ1, monEpPrecQ2
            End If
            cmdInfoMS.Caption = LabelCmdInfoMS + uneStruct.maCoucheSurface
        End If

        'Affichage commentaire structure choisie
        'RichTextAide.TextRTF = uneStruct.monComment
        'Activation du bouton permettant d'afficher les commentaires
        'de la structure choisie
        BtnInfoStruct.Enabled = True
    Else
        'Pas de structure choisie
        'd�sactivation des onglets CAM et couche de surface
        TabData.TabEnabled(OngletCAM) = uneStructChoisie
        TabData.TabEnabled(OngletSurf) = uneStructChoisie
        'Inhiber bouton info mat�riaux couches
        InhiberBoutonMat Me
        'Affichage aide g�n�rale
        unFileName = CorrigerNomFichier(App.Path + "\OngletStructure.rtf")
        RichTextAide.LoadFile unFileName, rtfRTF
        DoEvents
    End If

    'Mise � jour de l'affichage des carottes Q1 et Q2
    FrameCarotte.Visible = uneStructChoisie
    If uneStructChoisie Then AfficherCarottes Me
End Sub

Private Sub ComboStruct_DropDown()
    'Modifie par click  dans la liste d�roulante de la combobox
    'Stockage de l'indice de la structure s�lectionn�e
    'On en le fait que si on change de structure dans la liste de structures
    'du m�me type, si on change de type la liste n'est plus la m�me, donc on ne
    'stocke pas dans le tag le num�ro d'index de la structure choisie dans
    'la liste pr�c�dente
    If ComboStruct.Tag <> Format(ChangeTypeStruct) Then
        ComboStruct.Tag = Format(ComboStruct.ListIndex)
    End If
End Sub

Private Sub ComboStruct_GotFocus()
    If ComboStruct.ListCount = 0 Then
        MsgBox "Choisissez d'abord un type de structure", vbInformation
        Me.RichTextAide.SetFocus
    End If
End Sub

Private Sub ComboTailleAgglo_Click()
    AfficherEtCalculerIndGelRef Me
End Sub

Private Sub TestSaisirEtAfficherCarottes()
    SaisirEpaisseursCarottePourTestDessin
    FrameCarotte.Visible = True
End Sub



Private Sub CalculerNEth�orique()
    Dim uneStruct As New Structure
    Dim unNE As Long, unInd As Integer
    
    uneString = InputBox("Entrez unNE et unNEMin : ")
    unePos = InStr(1, uneString, " ")
    unNE = CLng(Mid(uneString, 1, unePos - 1))
    uneStruct.monNbEssieuxMin = CLng(Mid(uneString, unePos + 1))
    MsgBox "unNE = " + Format(unNE) + " et unNEmin = " + Format(uneStruct.monNbEssieuxMin)
    TrouverNEthEtInd Me, uneStruct, unNE, unInd
    MsgBox "unNE = " + Format(unNE) + " et unInd = " + Format(unInd)
End Sub


Private Sub Form_Activate()
    'V�rification que la saisie de la fen�tre pr�c�demment active
    'soit finie et valide
    If maFinSaisie = 0 Then
        If VerifierFinSaisie Then
            'Si OK ==> Changement d'�tude active
            'Stockage de l'�tude en cours
            Set monEtude = Me
            monEtude.maFinSaisie = 0
        Else
            'On revient dans la fen�tre pr�c�dent
            'on la remet au premier plan
            monEtude.maFinSaisie = 1 'Pour ne pas red�clencher la v�rif de fin de saisie
            monEtude.WindowState = vbNormal
            monEtude.ZOrder 0
        End If
    Else
        monEtude.maFinSaisie = 0 'Pour pouvoir red�clencher la v�rif de fin de saisie
    End If
    'Mise � jour du contexte id sur l'aide de l'onglet actif
    ChangerHelpID TabData.Tab
    
    If unNbActivate = 0 Then
        'R�-affichage du type d'�tude et les autres radio boutons, cause bug au load de la form
        'Pour la version 2 b�tatest
        'UNIQUEMENT LORS DU PREMIER ACTVIVATE (cela �quivaut au load)
        If monTypeEtude = TypeEtudeStandard Then
            OptionEtudeStandard.Value = True
        ElseIf monTypeEtude = TypeEtudeGiratoire Then
            OptionEtudeGiratoire.Value = True
        Else
            'Cas erreur de programmation ==> on met �tude standard
            MsgBox MsgErreurProg + MsgErreurTypeEtudeInconnue + MsgIn + "ModuleMain:MettreAJourOngletVoie", vbCritical
            OptionEtudeStandard.Value = True
        End If
        'Idem type de chantier
        'Par d�faut on est en condition Q1
        If monTypeChantier = TypeChantierQ1 Then
            OptionChoixQ1.Value = True
            OptionChoixQ2.Value = False
            LabelQualit�.Caption = "" 'en V2 on ne met rien "Chantier standard (Q1)"
        Else
            OptionChoixQ1.Value = False
            OptionChoixQ2.Value = True
            LabelQualit�.Caption = "" 'en V2 on ne met rien "Chantier difficile (Qualit� Q2)"
        End If
        'Affectation du nbActivate pour ne plus faire ce code l�
        unNbActivate = 1
        'Indication pour ne pas enregistrer lors d'une ouverture
        'Au premier affichage rien n'a pu �tre modifi�
        If maNewEtude = False Then maModif = False
    End If
End Sub


Private Sub Form_Initialize()
    'Variable indiquant s'il faut afficher le message d'erreur
    'lors du calcul du NE lorqu'une erreur se produit
    'Initialis� � vrai au d�part
    Dim unTypeVoieEtudeChantier As Integer
    
    monAffichageErreurNE = True
            
    If maNewEtude Then
        'Mise � jour des indices dans les combobox
        'dans le cas d'une nouvelle �tude car elles vont de 0 � n-1
        monIndStructChoisie = 0
        monIndStation = mesOptionsGen.monIndStationRef
        monIndTailleAgglo = mesOptionsGen.monTailleAgglo
        monHAgglo = mesOptionsGen.monAltiAgglo
        monEpNonGel = 0
        monCoefA = 0.12
        maPente = MsgInfinie
        monIndGelSol = 1
        monMinTecQ1 = "Non"
        monMaxPraQ1 = "Non"
        monMinTecQ2 = "Non"
        monMaxPraQ2 = "Non"
        maCroisAnnuel = 100 'Valeur pour non renseign�e
        monIndiceGelRefQ1 = -1
        monIndiceGelRefQ2 = -1
        
        'Pour la version 2 b�tatest
        monTypeEtude = TypeEtudeStandard
        monTypeChantier = TypeChantierQ1
        
        'Signal que l'on vient de cr�er une �tude
        monJustOpen = True
    Else
        'Cas d'une ouverture d'une �tude existante
        'R�cup�ration dans la collection de lecture de donn�es
        '(cf fonction ModuleFichier:OuvrirEtude + format de fichier *.urb)
        'Donn�es pour l'onglet Voie
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurOuverture
        
        monFichId = Val(maColLectFich(1))
        monTitreEtude = maColLectFich(3)
        
        'D�codage � partir de la version 2, de l'entier lu en quatri�me
        'position dans le fichier *.urb
        unTypeVoieEtudeChantier = maColLectFich(4)
        If unTypeVoieEtudeChantier < ChantierDifficile Then
            'Cas d'un chantier standard = qualit� Q1
            monTypeChantier = TypeChantierQ1
        Else
            'Cas d'un chantier difficile = qualit� Q2
            monTypeChantier = TypeChantierQ2
            'R�cup�ration du type de voie et d'�tude en enlevant
            'le rajout forfaitaire de la valeur ChantierDifficile
            unTypeVoieEtudeChantier = unTypeVoieEtudeChantier - ChantierDifficile
        End If
        'R�cup du type d'�tude
        If unTypeVoieEtudeChantier < TypeGiratoireDistribution Then
            'Type d'�tude standard
            monTypeEtude = TypeEtudeStandard
        Else
            'Type d'�tude giratoire
            monTypeEtude = TypeEtudeGiratoire
        End If
        'R�cup du type de voie
        monTypeVoie = unTypeVoieEtudeChantier
        
        maVariante = maColLectFich(5)
        maDate = maColLectFich(6)
        'Donn�es pour l'onglet Trafic
        If maColLectFich(7) = "" Then
            monTraficIni = 0
        Else
            monTraficIni = CInt(maColLectFich(7))
        End If
        maDur�eService = Val(maColLectFich(8))
        maCroisAnnuel = Val(maColLectFich(9))
        If maColLectFich(10) = "" Then
            monTraficCumul� = 0
        Else
            monTraficCumul� = CLng(maColLectFich(10))
        End If
        'Donn�es pour l'onglet Structure
        monIndStructChoisie = Val(maColLectFich(12))
        monUtilFichPerso = Val(maColLectFich(13))
        monFichPersoSTR = maColLectFich(14)
        'donn�es pour les ongletsCAM et plateforme
        monCAM = maColLectFich(15)
        monIndicePF = Val(maColLectFich(16))
        'Donn�es pour l'onglet Couche de surface
        monIndCompQ1 = Val(maColLectFich(17))
        monIndCompQ2 = Val(maColLectFich(18))
        'Donn�es pour l'onglet Gel
        monIndHiver = Val(maColLectFich(19))
        monIndStation = Val(maColLectFich(20))
        monHAgglo = Val(maColLectFich(21))
        monIndTailleAgglo = Val(maColLectFich(22))
        monIndGelSol = Val(maColLectFich(23))
        maPente = maColLectFich(24)
        monCoefA = CSng(maColLectFich(25))
        monEpNonGel = CInt(maColLectFich(26))
        'R�cup �ventuelle de l'indice de gel perso si fichier au format de
        'la version finale, apr�s les corrections suite aux sites pilotes
        If maColLectFich(28) = FormatFichierVersionFinale Then
            monIndGelPerso = CInt(maColLectFich(29))
            monUtilIndGelPerso = CByte(maColLectFich(30))
        End If
        
        'Signal que l'on vient d'ouvrir l'�tude
        monJustOpen = True
        
        ' D�sactive la r�cup�ration d'erreur.
        On Error GoTo 0
   End If
   
    'Sortie pour �viter la routine de gestion d'erreur
    monOuverture = True 'Pas d'erreur � signaler
    Exit Sub
   
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurOuverture:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + MsgEtude + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    MsgBox unMsg, vbCritical
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    monOuverture = False
End Sub


Private Sub Form_Load()
    Dim uneString As String, unTextCAM As String
    
    'Masquage du titre de la frame carotte donnant les deux qualit�s
    'en V2 ce titre disparait
    LabelQualit�.Visible = False
    
    'Remplissage et positionnement des labelInfo de la frame chantier
    LabelInfo1.Caption = LabelInfo1Caption
    LabelInfo2.Caption = LabelInfo2Caption
    LabelInfo1.Left = OptionChoixQ1.Left
    LabelInfo1.Top = OptionChoixQ1.Top
    LabelInfo2.Left = OptionChoixQ2.Left
    LabelInfo2.Top = OptionChoixQ2.Top
    BtnPlusInfo.Top = LabelInfo2.Top + (LabelInfo2.Height - BtnPlusInfo.Height) / 2
    BtnPlusInfo.Left = LabelInfo2.Left + LabelInfo2.Width + 60
       
    'Masquage des info sur l'altitude de l'agglo (pas utilis�e)
    LabAlt.Visible = False
    TextHAgglo.Visible = False
    LabAltUnit.Visible = False
    'Centrage des info sur l'altitude de la station de r�f�rence
    'm�me position que l'indice de gel perso
    LabelHStatUnit.Top = LabelIndGelPerso.Top
    LabelHStation.Top = LabelIndGelPerso.Top
    HStation.Top = LabelIndGelPerso.Top
    
    If monOuverture = False Then
        'Cas d'une erreur dans le form_initialize
        'on sort en fermant la fenetre
        monOuverture = True
        ViderCollection maColLectFich
        Unload Me
        Exit Sub
    End If
    
    'D�calage des fen�tres de 330 twpis en X et Y � chaque fen�tre
    Top = (Forms.Count - 2) * 330
    Left = (Forms.Count - 2) * 330
    fMainForm.Arrange vbCascade
    
    If maNewEtude = False Then
        'Mise � jour du titre de la fen�tre de l'�tude
        Caption = maColLectFich(27)
        ViderCollection maColLectFich
    End If
    
    'Tableau contenant les �paisseurs des couches de la structure choisie
    'Pour indice = 0 on prend 2 cm pour afficher l'ent�te de la carotte en impression
    'De 1 � 6 ==> Epaisseurs pour la qualit� Q1
    '1 et 2 �paisseur couche surface, 3 et 4 �paisseur couche base, 5 et 6 �paisseur couche fondation
    'De 7 � 12 ==> Epaisseurs pour la qualit� Q2
    '7 et 8 �paisseur couche surface, 9 et 10 �paisseur couche base, 11 et 12 �paisseur couche fondation
    monTabEp = Array(2, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0)
    
    'Mise � jour des boutons dans la toolbar permettant l'impression
    'et la sauvegarde car il n'y a pas de fen�tre fille ouverte
    '==> Impression et sauvegarde impossible
    fMainForm.tbToolBar.Buttons("Print").Visible = True
    fMainForm.tbToolBar.Buttons("Save").Visible = True
    
    'Calcul de la largeur de la zone aide rtf
    RichTextAide.RightMargin = RichTextAide.Width - 500
    'Initialisation de l'aide contextuelle d'onglet
    unFileName = CorrigerNomFichier(App.Path + "\OngletVoie.rtf")
    RichTextAide.LoadFile unFileName, rtfRTF
            
    'Mise � jour des couleurs des couches des carottes Q1 et Q2
    'en prenant celles des options g�n�rales
    ChangerCouleurCouches Me
    
    'Mise � jour de l'affichage de la fen�tre
    If AfficherTypeVoie(monTypeVoie) Then
        'Cas sans erreur � l'affichage du type de voie
        
        'Mise � jour de l'affichage des diff�rents onglets
        MettreAJourOngletVoie Me
        MettreAJourOngletTrafic Me
                
        'R�cup du s�parateur d�cimale . ou ,
        'fix� dans les param�tres r�gionaux de Windows
        TrouverCaract�reD�cimalUtilis�
        
        'Mise � jour de l'onglet CAM (deux chiffres apr�s la virgule)
        If monCAM <> "" Then
            'Affectation du bon mask pour la zone de saisie CAM
            MaskCAM.Mask = "#" + monCarDeci + "##"
            'MaskCAM.Text = Format(monCAM, "fixed")
            'bug si stockage de la valeur en . et
            'si s�parateur d�cimale actuel = virgule
            MaskCAM.Text = Mid(monCAM, 1, 1) + monCarDeci + Mid(monCAM, 3)
            'On prend de part et d'autre du s�parateur d�cimale
            'car le cam a un format #.##
        Else
            MaskCAM.Text = ""
        End If
        
        MettreAJourOngletStructure Me, monUtilFichPerso, monIndStructChoisie

        'Mise � jour de l'onglet plateforme
        MettreAjourOngletPlateForme Me
        
        'Mise � jour de l'onglet gel
        MettreAJourOngletGel Me
        
        'Mise � jour de l'affichage de la frame R�sultat
        MettreAJourFrameR�sultat Me
        
        'Mise � jour de la frame de v�rif au gel
        ActualiserFrameVerifGel Me
    Else
        'Cas avec erreur � l'affichage du type de voie
        '==> on indique que la fen�tre est � fermer
        'pour les fonctions LoadNew et ouverture d'une �tude
        Tag = "A_Fermer"
    End If
    
    'Indication qu'aucun change event de textbox n'a �t� d�clench�e
    monChangeEvent = False
    If maNewEtude = False Then maModif = False 'Indication pour l'enregistrer
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim unNomFich As String
    
    Cancel = Not VerifierFinSaisie
    'si Cancel = true on ne sort pas de l'application
    If Cancel = False Then
        'Demande de sauvegarde si modif
        If ModifierEtude(Me) Or EstNouvelleEtude(Me) Then
            If EstNouvelleEtude(Me) Then
                unNomFich = Caption
            Else
                unNomFich = Mid(Caption, 7)
            End If
            uneRep = MsgBox(MsgSaveFile + unNomFich + " ?", vbExclamation + vbYesNoCancel)
            If uneRep = vbCancel Then
                'Pas de sortie, on ne fait rien
                Cancel = True
            ElseIf uneRep = vbYes Then
                'Sauvegarde puis sortie
                If SauverEtude(Me, unNomFich, False) = "" Then
                    'Si pas de fichier choisie ==> on ne sort pas
                    Cancel = True
                Else
                    Cancel = False
                End If
            Else
                'Cas du click sur Non ==> On sort
                Cancel = False
            End If
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Forms.Count = 2 Then
        'quand on ferme la derni�re fen�tre fille ouverte, on est
        'dans le unload donc elle existe encore et il reste la
        'fenetre MDI m�re
        '==> Lors de la Fermeture derni�re fille, Nombre de Forms = 2
        
        'Mise � jour des boutons dans la toolbar permettant l'impression
        'et la sauvegarde s'il n'y a plus de fen�tre fille ouverte
        '==> Impression et sauvegarde impossible
        fMainForm.tbToolBar.Buttons("Print").Visible = False
        fMainForm.tbToolBar.Buttons("Save").Visible = False
        
        'On remet le contexte id � 0 car plus de fen�tres filles ouvertes
        fMainForm.HelpContextID = 0
    End If
    
    Close #monFichId
    Set monEtude = Nothing
End Sub

Public Function AfficherChoixTypesVoies() As Boolean
    'Affiche les types de voie dans l'onglet Voie suivant le type d'�tude
    'Standard ou Giratoire
    If OptionEtudeStandard.Value Then
        monTypeEtude = TypeEtudeStandard
        'Affichage des donn�es en �tude standard (= section courante)
        FrameVoie.Visible = True
        OptionChoixQ1.Visible = True
        OptionChoixQ2.Visible = True
        'Masquage des donn�es en �tude giratoire
        FrameGiratoire.Visible = False
        LabelInfo1.Visible = False
        LabelInfo2.Visible = False
        BtnPlusInfo.Visible = False
        'La plateforme PF1 est possible
        OptionPF1.Enabled = True
    ElseIf OptionEtudeGiratoire.Value Then
        monTypeEtude = TypeEtudeGiratoire
        'Masquage des donn�es en �tude standard (= section courante)
        FrameVoie.Visible = False
        OptionChoixQ1.Visible = False
        OptionChoixQ2.Visible = False
        'On remet condition de chantier � Q1, car tout est remis � z�ro
        OptionChoixQ1.Value = True
        'Masquage des donn�es en �tude giratoire
        FrameGiratoire.Visible = True
        LabelInfo1.Visible = True
        LabelInfo2.Visible = True
        BtnPlusInfo.Visible = True
        'La plateforme PF1 n'est pas possible
        OptionPF1.Enabled = False
        If OptionPF1.Value Then
            'Si PF1 �tait la plateforme choisie, on se met en PF2
            OptionPF2.Value = True
        End If
    Else
        'Cas erreur de programmation ==> on met �tude standard
        MsgBox MsgErreurProg + MsgErreurTypeEtudeInconnue + MsgIn + "frmDocument:AfficherChoixTypesVoies", vbCritical
        FrameVoie.Visible = True
        FrameGiratoire.Visible = False
    End If

    LabelQualit�.Caption = "" 'en V2 on enl�ve le titre de la frame carottes
    'On met les voies giratoires au m�me endroit que les voies
    FrameGiratoire.Left = FrameVoie.Left
    FrameGiratoire.Top = FrameVoie.Top
    
    If monFichId = 0 Or unNbActivate > 0 Then
        'On fait ce traitement si c'est une nouvelle �tude (monFichId = 0,
        'sinon valeur >= 1 issue du fichier urb) et si ce n'est pas le premier
        'appel � la fonction AfficherChoixTypesVoies lors d'une ouverture d'une
        '�tude issue d'un ficher urb, cet appel se d�clenche au premier activate
        'de la form, donc pour unNbActivate = 0
        
        'Remise � vide des boutons s�lectionn�s
        Me.OptionVoieDes.Value = False
        Me.OptionVoieDis.Value = False
        Me.OptionVoiePL.Value = False
        Me.OptionVoieBus.Value = False
        Me.OptionVoieParking.Value = False
        
        Me.OptionGirDis.Value = False
        Me.OptionGirPL.Value = False
        
        AfficherTypeVoie TypeVoieInconnu
        'Inhibition de l'onglet gel
        Me.TabData.TabEnabled(OngletGel) = False
        'Mettre � jour affichage de l'indice de gel admissible et r�f�rence
        'monIndiceGelAdmQ1 = 0
        'monIndiceGelAdmQ2 = 0
        'monIndiceGelRefQ1 = -1
        'monIndiceGelRefQ2 = -1
        'ActualiserFrameVerifGel Me
    End If
End Function

Public Function AfficherTypeVoie(unTypeVoie As Integer) As Boolean
    'Affiche le type de voie dans la fen�tre r�sultat de gauche
    'et dans l'ongle Voie � droite et la valeur pr�conis�e pour
    'ce type de voie pour le CAM dans l'onglet CAM
    'Retourne Vrai si aucune erreur, faux sinon
    Dim uneString As String
    
    'Initialisation � vide du trafic initial et du trafic cumul�
    'Utile lors d'un changement de type de voies
    TextTrafIni.Text = ""
    TextTrafCUM.Text = ""
    
    'Initialisation � vide du CAM
    'la bonne valeur est mis apr�s dans le Form_Load
    MaskCAM.Mask = ""
    MaskCAM.Text = ""
    LabelCAM.Caption = LabelCAMCaption + MsgInconnu
    
    'Initialisation � vide du NE
    LabelNEequiv.ForeColor = QBColor(0)
    LabelNEequiv.Caption = LabelNEequivCaption + MsgInconnu
    
    'Initialisation de la valeur de retour
    AfficherTypeVoie = True
    
    'S�lection du bon bouton option pour le type de voies
    'et Mise � jour des valeurs possible et par d�faut des
    'hivers dans l'onglet Gel si c'est une nouvelle �tude
    If unTypeVoie = TypeVoieInconnu Then
        uneString = MsgInconnu
        'Pas d'hiver par d�faut
        OptionHE.Value = False
        OptionHRNE.Value = False
        OptionHC.Value = False
        'Indice de gel de r�f�rence inconnu
        monIndiceGelRefQ1 = -1
        monIndiceGelRefQ2 = -1
        'On consid�re que c'est comme une nouvelle �tude
        maNewEtude = True
    ElseIf unTypeVoie = TypeVoieDesserte Then
        uneString = OptionVoieDes.Caption
        OptionVoieDes.Value = True
        OptionHE.Enabled = False
        If maNewEtude Then OptionHC.Value = True
    ElseIf unTypeVoie = TypeVoieDistribution Then
        uneString = OptionVoieDis.Caption
        OptionVoieDis.Value = True
        OptionHE.Enabled = False
        If maNewEtude Then OptionHC.Value = True
    ElseIf unTypeVoie = TypeVoieTraficLourd Then
        uneString = OptionVoiePL.Caption
        OptionVoiePL.Value = True
        OptionHE.Enabled = True
        If maNewEtude Then OptionHRNE.Value = True
    ElseIf unTypeVoie = TypeVoieBus Then
        uneString = OptionVoieBus.Caption
        OptionVoieBus.Value = True
        OptionHE.Enabled = True
        If maNewEtude Then OptionHRNE.Value = True
    ElseIf unTypeVoie = TypeVoieParking Then
        uneString = OptionVoieParking.Caption
        OptionVoieParking.Value = True
        OptionHE.Enabled = False
        If maNewEtude Then OptionHC.Value = True
    ElseIf unTypeVoie = TypeGiratoireDistribution Then
        uneString = OptionGirDis.Caption
        OptionGirDis.Value = True
        OptionHE.Enabled = False
        If maNewEtude Then OptionHC.Value = True
    ElseIf unTypeVoie = TypeGiratoireTraficLourd Then
        uneString = OptionGirPL.Caption
        OptionGirPL.Value = True
        OptionHE.Enabled = True
        If maNewEtude Then OptionHRNE.Value = True
    Else
        'Cas erreur de programmation ==> on sort de la fen�tre fille
        MsgBox MsgErreurProg + MsgErreurTypeVoieInconnu + MsgIn + "frmDocument:AfficherTypeVoie", vbCritical
        AfficherTypeVoie = False
    End If
    
    LabelTypeVoie.Caption = LabelTypeVoieCaption + uneString
    
    'Mise � jour du gris� des onglets ci-dessous
    'enable si untypevoie > � 0
    TabData.TabEnabled(OngletTrafic) = (unTypeVoie > 0)
    TabData.TabEnabled(OngletStruct) = (unTypeVoie > 0)
    TabData.TabEnabled(OngletPF) = (unTypeVoie > 0)
    TabData.TabEnabled(OngletGel) = (unTypeVoie > 0) And mesOptionsGen.maVerifGel
       
    'Mise � jour du type de voie de l'etude active
    'et de la liste de structures possibles et aucune structure choisie
    If unTypeVoie > 0 Then
        monTypeVoie = unTypeVoie
        'En version, la liste des structures est rempli une fois
        'le type de structure choisi ==> On ne fait plus RemplirComboStructures
        'RemplirComboStructures Me
        'On active juste les radio boutons de type de structures possibles
        'par rapport au type de voie choisi
        ActiverTypeStructure Me
        maModif = True
    End If
    InhiberBoutonMat Me
    
    'Autres onglets
    'correction d'un bug V1 en V2, apr�s ouvir etude, puis changement de voie,
    'l'onglet CAM est encore actif et le click sur cet onglet CAM fait planter
    'Struct-urb version 1 d'o� la correction de la ligne ci-dessous
    'TabData.TabEnabled(OngletCAM) = (monIndStructChoisie > 0)
    TabData.TabEnabled(OngletCAM) = (ComboStruct.ListIndex >= 0)
    TabData.TabEnabled(OngletSurf) = False
    
    'On enl�ve les carottes visibles
    FrameCarotte.Visible = False
    
    'Inhibition du bouton info structure
    BtnInfoStruct.Enabled = False
    
    'Mettre � jour affichage de l'indice de gel admissible
    monIndiceGelAdmQ1 = 0
    monIndiceGelAdmQ2 = 0
    ActualiserFrameVerifGel Me
    
    'Remise � z�ro du NE equivalent
    monNEEquiv = 0
    
    'En version 2, si le type de voie est parking
    'on inhibe l'onglet Trafic
    'et on prend TraficCumul� = 100 000, on mettra CAM = 0.1 lors du choix
    'de la structure
    'pour avoir un NE = 10 000, seul disponible pour les
    'structures de voie de parking
    If unTypeVoie = TypeVoieParking Then
        TabData.TabEnabled(OngletTrafic) = False
        TabData.TabEnabled(OngletCAM) = False
        Me.TextTrafCUM.Text = "100 000"
        'Mise � jour du trafic cumul� dans la frame R�sultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
    End If
End Function




Private Sub LabelInfoMaxPQ1_Click()
    MsgBox MsgMaxPra1 + "Q1" + MsgMaxPra2 + Chr(13) + Chr(13) + MsgMaxPra3 + Chr(13) + Chr(13) + MsgMaxPra5 + Chr(13) + Chr(13) + MsgMaxPra6, vbInformation
    RichTextAide.SetFocus
End Sub

Private Sub LabelInfoMaxPQ2_Click()
    Dim unMsg As String
    unMsg = MsgMaxPra1 + "Q2" + MsgMaxPra2 + Chr(13) + Chr(13) + MsgMaxPra3
    If monEpQ1Trouv Then unMsg = unMsg + Chr(13) + Chr(13) + MsgMaxPra4
    unMsg = unMsg + Chr(13) + Chr(13) + MsgMaxPra5 + Chr(13) + Chr(13) + MsgMaxPra6
    MsgBox unMsg, vbInformation
    RichTextAide.SetFocus
End Sub

Private Sub LabelInfoMinTechnoQ1_Click()
    MsgBox MsgMinTechno1 + Chr(13) + Chr(13) + MsgMinTechno2, vbInformation
    RichTextAide.SetFocus
End Sub

Private Sub LabelInfoMinTechnoQ2_Click()
    MsgBox MsgMinTechno1 + Chr(13) + Chr(13) + MsgMinTechno2, vbInformation
    RichTextAide.SetFocus
End Sub

Private Sub ListViewMat_Click()
    'Affichage de la fiche caract�ristique du mat�riau
    'Param�tre string = TypeMat + "/" + Abr�g� du Mat�riau
    AfficherFicheMat "Composant" + "/" + ListViewMat.SelectedItem.Text
 End Sub


Private Sub MaskCAM_Change()
    If MaskCAM.Tag <> MaskCAM.Text Then monAffichageErreurNE = True
    LabelCAM.Caption = LabelCAMCaption + MaskCAM.Text
End Sub

Private Sub MaskCAM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        'D�but de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    End If
End Sub

Private Sub MaskCAM_KeyPress(KeyAscii As Integer)
    TrouverCaract�reD�cimalUtilis�
    
    If KeyAscii = 27 Then  'Touche Escape
        'Restauration derni�re valeur valide
        MaskCAM.Text = MaskCAM.Tag
        'Changer le control actif pour d�clencher le LostFocus de MaskCAM
        RichTextAide.SetFocus
    ElseIf KeyAscii = 13 Then 'touche Return
        'Changer le control actif pour d�clencher le LostFocus de MaskCAM
        RichTextAide.SetFocus
    ElseIf KeyAscii = 46 And monCarDeci = "," Then
        'Touche . tap�e or caract�re d�cimale = virgule, on remplace
        KeyAscii = 44 '44 = virgule
        'D�but de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    ElseIf KeyAscii = 44 And monCarDeci = "." Then
        'Touche , tap�e or caract�re d�cimale = point, on remplace
        KeyAscii = 46 '46 = point
        'D�but de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        'Cas o� l'on tape un chiffre entre 0 et 9 ou la touche baskspace
        'D�but de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    End If
End Sub

Private Sub MaskCAM_LostFocus()
    If DonnerStructChoisie(Me) Is Nothing Or MaskCAM.Text = "" Then Exit Sub
    
    If VerifierMinMaxCAM(Me, MaskCAM.Text) Then
        TrouverCaract�reD�cimalUtilis�
        MaskCAM.Mask = "#" + monCarDeci + "##"
        MaskCAM.Text = Format(MaskCAM.Text, "fixed")
        'Fin de modif ==> Remise de la zone de saisie en noir
        MaskCAM.ForeColor = QBColor(0)
        'Stockage dans la derni�re valeur valide
        MaskCAM.Tag = MaskCAM.Text
        'Calcul du Nombre d'essieux �quivalents si trafic cumul� connu
        CalculerEtAfficherNE
    Else
        'Cas d'erreur ==> on retourne dans MaskCAM
        TabData.Tab = OngletCAM 'Retour � l'onglet CAM
        MaskCAM.SetFocus
    End If
End Sub


Private Sub OptionANT_Click()
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
End Sub

Private Sub OptionAT_Click()
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
End Sub

Private Sub OptionCA_Click(Index As Integer)
    monAffichageErreurNE = True
    CalculerTraficCum Me
    'Mise � jour du trafic cumul� dans la frame R�sultats
    LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
    TextTrafCUM.ForeColor = QBColor(0)
End Sub

Private Sub OptionChoixQ1_Click()
    monTypeChantier = TypeChantierQ1
    LabelQualit�.Caption = "" 'en V2 on ne met plus rien "Chantier standard (Q1)"
    AfficherCarottes Me
    ActualiserFrameVerifGel Me
    AfficherCSurfUneQualite
    'Indication d'un changement pour d�clencher le save as
    maModif = True
End Sub

Private Sub OptionChoixQ2_Click()
    monTypeChantier = TypeChantierQ2
    LabelQualit�.Caption = "" 'en V2 on ne met plus rien "Chantier difficile (Q2)"
    AfficherCarottes Me
    ActualiserFrameVerifGel Me
    AfficherCSurfUneQualite
    'Indication d'un changement pour d�clencher le save as
    maModif = True
    'Message d'avertissement pour travailler en qualit� Q2 ou chantier difficile
    MsgBox MsgPassageEnQ2_1 + Chr(13) + Chr(13) + MsgPassageEnQ2_2 + Chr(13) + MsgPassageEnQ2_3 + Chr(13) + MsgPassageEnQ2_4, vbInformation
End Sub

Private Sub OptionEtudeGiratoire_Click()
    AfficherChoixTypesVoies
End Sub

Private Sub OptionEtudeStandard_Click()
    AfficherChoixTypesVoies
End Sub


Private Sub OptionGirDis_Click()
    AfficherTypeVoie TypeGiratoireDistribution
End Sub

Private Sub OptionGirPL_Click()
    AfficherTypeVoie TypeGiratoireTraficLourd
End Sub

Private Sub OptionHC_Click()
    AfficherEtCalculerIndGelRef Me
End Sub

Private Sub OptionHE_Click()
    AfficherEtCalculerIndGelRef Me
End Sub

Private Sub OptionHRNE_Click()
    AfficherEtCalculerIndGelRef Me
End Sub

Private Sub OptionNGel_Click()
    If OptionNGel.Tag = "" Then TextPente.Text = Format(0)
    TextPente.ForeColor = QBColor(0)
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
End Sub

Private Sub OptionPF1_Click()
    LabelPFQ1.Caption = "PF1"
    LabelPFQ2.Caption = "PF1"
    'Recherche des �paisseurs si tout est d�fini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre � jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPF2_Click()
    LabelPFQ1.Caption = "PF2"
    LabelPFQ2.Caption = "PF2"
    'Recherche des �paisseurs si tout est d�fini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre � jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPF2Plus_Click()
    LabelPFQ1.Caption = "PF2+"
    LabelPFQ2.Caption = "PF2+"
    'Recherche des �paisseurs si tout est d�fini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre � jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPF3_Click()
    LabelPFQ1.Caption = "PF3"
    LabelPFQ2.Caption = "PF3"
    'Recherche des �paisseurs si tout est d�fini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre � jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPGel_Click()
    If OptionPGel.Tag = "" Then TextPente.Text = Format(0.4)
    TextPente.ForeColor = QBColor(0)
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
End Sub

Private Sub OptionTGel_Click()
    If OptionTGel.Tag = "" Then TextPente.Text = MsgInfinie
    TextPente.ForeColor = QBColor(0)
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
End Sub


Private Sub OptionTypeStruct_Click(Index As Integer)
    'En version 2, on remplit la liste des structures avec
    'le bon type de structures
    DoEvents
    RemplirComboStructures Me, Index
    DoEvents
    'L'index a la m�me valeur que les constantes de type
    'de structures
    'Const Souple As Byte = 1
    'Const Bitumineuse As Byte = 2
    'Const GTLH As Byte = 3
    'Const Beton As Byte = 4
    'Const Mixte As Byte = 5
    'Const PavesDalles As Byte = 6
    
    'Inhibition du bouton info structure
    BtnInfoStruct.Enabled = False
    'Inhibition des boutons d'info sur les diff�rents mat�riaux de structure
    InhiberBoutonMat Me
    'Inhibition de l'onglet CAM et couche de surface
    TabData.TabEnabled(OngletCAM) = False
    TabData.TabEnabled(OngletSurf) = False
    
    'Mettre � jour affichage de l'indice de gel admissible
    monIndiceGelAdmQ1 = 0
    monIndiceGelAdmQ2 = 0
    ActualiserFrameVerifGel Me
    'Masquage du taux de risque car plus de structure s�lectionn�e
    LabelRisk.Visible = False
    'On enl�ve les carottes visibles
    FrameCarotte.Visible = False
End Sub

Private Sub OptionVoieBus_Click()
    AfficherTypeVoie TypeVoieBus
End Sub

Private Sub OptionVoieDes_Click()
    AfficherTypeVoie TypeVoieDesserte
End Sub

Private Sub OptionVoieDis_Click()
    AfficherTypeVoie TypeVoieDistribution
End Sub

Private Sub OptionVoieParking_Click()
    AfficherTypeVoie TypeVoieParking
End Sub

Private Sub OptionVoiePL_Click()
    AfficherTypeVoie TypeVoieTraficLourd
End Sub


Private Sub TabData_Click(PreviousTab As Integer)
    Dim uneStruct As Structure
    
    'Mise � jour du contexte pour l'aide avec F1
    ChangerHelpID TabData.Tab
    
    'Mise de l'aide contextuelle RTF
    Set uneStruct = DonnerStructChoisie(Me)
    unFileName = CorrigerNomFichier(App.Path + "\Onglet")
    RichTextAide.LoadFile unFileName + TabData.TabCaption(TabData.Tab) + ".rtf", rtfRTF
    'R�cup du s�parateur d�cimale . ou ,
    'fix� dans les param�tres r�gionaux de Windows
    TrouverCaract�reD�cimalUtilis�
    
    If TabData.Tab = OngletTrafic Then
        'Si le CAM est valide (pas en rouge), le trafic ini prend le focus
        If MaskCAM.ForeColor <> QBColor(12) Then
            TextTrafIni.SetFocus
            TextTrafIni.SelStart = Len(TextTrafIni.Text)
        End If
    ElseIf TabData.Tab = OngletCAM Then
        'Mise � jour de l'onglet CAM qui met aussi � jour le label CAM
        'de la frame r�sultat
        ActualiserOngletLabelCAM
        MaskCAM.SetFocus
        MaskCAM.SelStart = Len(MaskCAM.Text)
    ElseIf TabData.Tab = OngletGel Then
        'Mise � jour de l'onglet Gel
        'Affichage avec le bon caract�re d�cimal des
        'valeurs possibles du coefficient A
        OptionANT.Caption = "0" + monCarDeci + "12 (Non Trait�)"
        OptionAT.Caption = "0" + monCarDeci + "14 (Trait�)"
        If monUtilIndGelPerso = 0 Then
            ComboStation.SetFocus
            'Mise en commentaires des lignes ci-dessous car HAgglo pas utilis�e
            'TextHAgglo.SetFocus
            'TextHAgglo.SelStart = Len(TextHAgglo.Text)
        Else
            TextIndGelPerso.SetFocus
            TextIndGelPerso.SelStart = Len(TextIndGelPerso.Text)
        End If
    ElseIf TabData.Tab = OngletStruct Then
        'Masquage des info sur le fichir personnel de structure
        'en version 2, cel ane sert plus
        Me.LabelFichPerso.Visible = False
        Me.CheckFichPerso.Visible = False
    End If
End Sub


Private Sub TextDur�eS_Change()
    Dim unAjoutEnFinChaine As Boolean
    Dim uneSelStartDeb As Integer
                
    monAffichageErreurNE = True
    monChangeEvent = True
    'Modif du contenu ==> Mise en rouge
    TextDur�eS.ForeColor = QBColor(12)
    
    'Remise � inconnu du trafic cumul� dans la frame R�sultats
    LabelTraficCum.Caption = LabelTraficCumCaption + MsgInconnu
    
    If Val(TextDur�eS.Text) = 0 Then
        TextDur�eS.Text = "0"
        Exit Sub
    End If
    
    'R�cup de la point d'insertion initial
    uneSelStartDeb = TextDur�eS.SelStart
    
    'Test si on rajoute des caract�res en fin de chaine
    unAjoutEnFinChaine = (TextDur�eS.SelStart = Len(TextDur�eS))
    TextDur�eS.Text = Format(TextDur�eS.Text, "##")
    If unAjoutEnFinChaine Then
        TextDur�eS.SelStart = Len(TextDur�eS.Text)
    Else
        TextDur�eS.SelStart = uneSelStartDeb
    End If
End Sub


Private Sub TextDur�eS_KeyPress(KeyAscii As Integer)
    'V�rification qu'un chiffre est tap� et rien d'autre
    '8 = backspace
    
    'si on tape retour (=13), on fait comme un Tab,
    'on passe � la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If KeyAscii = 27 Then TextDur�eS.Text = TextDur�eS.Tag
        'Fin de modification ==> On remet le texte en noir
        TextDur�eS.ForeColor = QBColor(0)
        'Activation du controle suivant
        If TextDur�eS.Text <> "" Then RichTextAide.SetFocus
        Exit Sub
    End If
    
    If (KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57 Then
        Beep
        KeyAscii = 0 'Par annuler la frappe de ce non-chiffre
    Else
        'Cas o� la touche tap�e au clavier est un chiffre
        'Entr�e en modification ==> On met le texte en rouge
        TextDur�eS.ForeColor = QBColor(12)
        monChangeEvent = True
    End If
End Sub

Private Sub TextDur�eS_LostFocus()
    'Test pour �viter les lostfocus en cascade
    If maTrafIniErreur Or maTrafCumErreur Or Not (Screen.ActiveForm Is Me) Or monChangeEvent = False Then Exit Sub

    If VerifierMinMaxDur�eService(Me) Then
        'Contenu valide ==> Mise en noir
        TextDur�eS.ForeColor = QBColor(0)
        monChangeEvent = False
        'Sauvegarde de la saisie valide dans le tag
        'pour la touche echappement
        TextDur�eS.Tag = TextDur�eS.Text
        CalculerTraficCum Me
        maDurServErreur = False
        'Mise � jour du trafic cumul� dans la frame R�sultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
        TextTrafCUM.ForeColor = QBColor(0)
    Else
        'Contenu invalide ==> Mise en rouge
        TextDur�eS.ForeColor = QBColor(12)
        TabData.Tab = OngletTrafic 'Retour � l'onglet Trafic
        TextDur�eS.SetFocus
        maDurServErreur = True
    End If
End Sub

Private Sub TextEpaisseur_Change()
    VerifierSaisieEntier TextEpaisseur
    maModif = True 'Indication pour l'enregistrer
End Sub

Private Sub TextEpaisseur_KeyPress(KeyAscii As Integer)
    'si on tape retour (=13), on fait comme un Tab
    '==> Sortie donc Lost focus,
    'on passe � la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant et on passe � la text box suivante
   If KeyAscii = 27 Then
        TextEpaisseur.Text = TextEpaisseur.Tag
        RichTextAide.SetFocus
   ElseIf KeyAscii = 13 Then
        RichTextAide.SetFocus
   End If
End Sub

Private Sub TextEpaisseur_LostFocus()
    'Save pour la touche Echap'
    TextEpaisseur.Tag = TextEpaisseur.Text
    'Calcul de l'indice de gel admissible
    TextEpaisseur.ForeColor = QBColor(0)
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
End Sub

Private Sub TextHAgglo_Change()
    VerifierSaisieEntier TextHAgglo
    maModif = True 'Indication pour l'enregistrer
End Sub


Private Sub TextHAgglo_KeyPress(KeyAscii As Integer)
    'si on tape retour (=13), on fait comme un Tab
    '==> Sortie donc Lost focus,
    'on passe � la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant et on passe � la text box suivante
   If KeyAscii = 27 Then
        TextHAgglo.Text = TextHAgglo.Tag
        TextEpaisseur.SetFocus
   ElseIf KeyAscii = 13 Then
        TextEpaisseur.SetFocus
   End If
End Sub

Private Sub TextHAgglo_LostFocus()
    'Save pour la touche Echap'
    TextHAgglo.Tag = TextHAgglo.Text
    'Calcul de l'indice de gel corrig�
    AfficherEtCalculerIndGelRef Me
    TextHAgglo.ForeColor = QBColor(0)
End Sub


Private Sub TextIndGelPerso_Change()
    VerifierSaisieEntier TextIndGelPerso
    maModif = True 'Indication pour l'enregistrer
End Sub

Private Sub TextIndGelPerso_KeyPress(KeyAscii As Integer)
    'si on tape retour (=13), on fait comme un Tab
    '==> Sortie donc Lost focus,
    'on passe � la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant et on passe � la text box suivante
   If KeyAscii = 27 Then
        TextIndGelPerso.Text = TextIndGelPerso.Tag
        TextEpaisseur.SetFocus
   ElseIf KeyAscii = 13 Then
        TextEpaisseur.SetFocus
   End If
End Sub

Private Sub TextIndGelPerso_LostFocus()
    'Save pour la touche Echap'
    TextIndGelPerso.Tag = TextIndGelPerso.Text
    'Calcul de l'indice de gel corrig�
    monIndGelPerso = Format(TextIndGelPerso.Text)
    monIndiceGelRefQ1 = Format(TextIndGelPerso.Text)
    monIndiceGelRefQ2 = Format(TextIndGelPerso.Text)
    'Affichage dans la frame r�sultat de l'�tude active
    ActualiserFrameVerifGel Me
    TextIndGelPerso.ForeColor = QBColor(0)
End Sub

Private Sub TextPente_Change()
    If TextPente.Text = MsgInfinie Then
        maModif = True 'Indication pour l'enregistrer
        Exit Sub
    End If
    
    If Mid(TextPente.Text, 1, 1) = "0" And InStr(1, TextPente.Text, monCarDeci) = 0 And Len(TextPente.Text) > 1 Then
        TextPente.Text = Mid(TextPente.Text, 2)
        TextPente.SelStart = 1
    End If
    
    'D�but de modification ==> mise en rouge de la zone de saisie
    TextPente.ForeColor = QBColor(12)
    
    If IsNumeric(TextPente.Text) And InStr(1, TextPente.Text, "-") = 0 And InStr(1, TextPente.Text, "+") = 0 Then
        'Les instr recherchant + et - sont rajout�s car la fonction
        'IsNumeric consid�re qu'une chaine les contenant est num�rique
        'm�me si le + et le - sont n'importe o� dans la chaine
        maModif = True 'Indication pour l'enregistrer
    Else
        If TextPente.Text <> "" And TextPente.Text <> monCarDeci Then
            MsgBox MsgSaisieR�elPositif, vbCritical
            unePos = TextPente.SelStart
            TextPente.Text = Mid(TextPente.Text, 1, unePos - 1) + Mid(TextPente.Text, unePos + 1)
            TextPente.SelStart = unePos - 1
        End If
    End If
End Sub

Private Sub TextPente_KeyPress(KeyAscii As Integer)
    'R�cup du s�parateur d�cimale . ou ,
    'fix� dans les param�tres r�gionaux de Windows
    TrouverCaract�reD�cimalUtilis�
    
    If KeyAscii = 27 Then  'Touche Escape
        'Restauration derni�re valeur valide
        TextPente.Text = TextPente.Tag
        'Changer le control actif pour d�clencher le LostFocus de c
        TextEpaisseur.SetFocus
    ElseIf KeyAscii = 13 Then 'touche Return
        'Changer le control actif pour d�clencher le LostFocus de TextPente
        TextEpaisseur.SetFocus
    ElseIf KeyAscii = 46 And monCarDeci = "," Then
        'Touche . tap�e or caract�re d�cimale = virgule, on remplace
        KeyAscii = 44 '44 = virgule
    ElseIf KeyAscii = 44 And monCarDeci = "." Then
        'Touche , tap�e or caract�re d�cimale = point, on remplace
        KeyAscii = 46 '46 = point
    Else
        If TextPente.Text = MsgInfinie Then 'And IsNumeric(Chr(KeyAscii)) = False Then
            TextPente.Text = ""
        End If
    End If
End Sub


Private Sub TextPente_LostFocus()
    If TextPente.Text = monCarDeci Then
        TextPente.Text = "0"
    Else
        TextPente.Text = Format(TextPente.Text)
    End If
    'Save pour la touche Echap'
    TextPente.Tag = TextPente.Text
    'Calcul de l'indice de gel admissible
    AfficherEtCalculerIndGelAdm Me
    'Fin de modif on remet en noir
    TextPente.ForeColor = QBColor(0)
    'S�lection du bon type de sol support d'apr�s la pente
    If TextPente.Text = "" Then
    ElseIf TextPente.Text = MsgInfinie Then
        OptionTGel.Tag = "PasDeModif" 'Evite la remise � jour dans Option.click
        OptionTGel.Value = True
        OptionTGel.Tag = ""
    ElseIf CSng(TextPente.Text) > 0.4 Then
        OptionTGel.Tag = "PasDeModif" 'Evite la remise � jour dans Option.click
        OptionTGel.Value = True
        OptionTGel.Tag = ""
    ElseIf CSng(TextPente.Text) <= 0.05 Then
        OptionNGel.Tag = "PasDeModif" 'Evite la remise � jour dans Option.click
        OptionNGel.Value = True
        OptionNGel.Tag = ""
    Else
        OptionPGel.Tag = "PasDeModif" 'Evite la remise � jour dans Option.click
        OptionPGel.Value = True
        OptionPGel.Tag = ""
    End If
End Sub



Private Sub TextTrafCUM_Change()
    Dim unAjoutEnFinChaine As Boolean
    Dim uneSelStartDeb As Integer
    
    'Suppression du caract�re s�parateur millier si on enl�ve le premier
    'chiffre car format traf cum = ###,###,###
    'donc "200 456 888" peut devenir " 456 888" ou  " 888"
    '==> trafcum.text devient vide d'o� la correction
    If Mid(TextTrafCUM.Text, 1, 1) = DonnerSepMillier Then
        'Suppression du premier caract�re qui est le s�parateur de millier
        TextTrafCUM.Text = Mid(TextTrafCUM.Text, 2)
    End If
    
    'Modif du contenu ==> Mise en rouge
    If Format(TextTrafCUM.Text, "###,###,###") <> TextTrafCUM.Tag Then
        TextTrafCUM.ForeColor = QBColor(12)
        monAffichageErreurNE = True
    End If
    
    'Remise � inconnu du trafic cumul� dans la frame R�sultats
    LabelTraficCum.Caption = LabelTraficCumCaption + MsgInconnu
    
    'R�cup de la point d'insertion initial
    uneSelStartDeb = TextTrafCUM.SelStart
    
    'Test si on rajoute des caract�res en fin de chaine
    unAjoutEnFinChaine = (TextTrafCUM.SelStart = Len(TextTrafCUM))
    TextTrafCUM.Text = Format(TextTrafCUM.Text, "###,###,###")
    If unAjoutEnFinChaine Then
        TextTrafCUM.SelStart = Len(TextTrafCUM.Text)
    Else
        TextTrafCUM.SelStart = uneSelStartDeb
    End If
End Sub


Private Sub TextTrafCUM_KeyPress(KeyAscii As Integer)
    'V�rification qu'un chiffre est tap� et rien d'autre
    '8 = backspace
    
    'si on tape retour (=13), on fait comme un Tab,
    'on passe � la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If KeyAscii = 27 Then TextTrafCUM.Text = TextTrafCUM.Tag
        'Fin de modification ==> On remet le texte en noir
        TextTrafCUM.ForeColor = QBColor(0)
        'Activation du controle suivant
        If Val(TextTrafCUM.Text) > 0 Then
            RichTextAide.SetFocus
        Else
            TextTrafCUM.Text = "0"
            TextTrafIni.Text = ""
        End If
        Exit Sub
    End If
    
    If (KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57 Then
        Beep
        KeyAscii = 0 'Par annuler la frappe de ce non-chiffre
    Else
        'Cas o� la touche tap�e au clavier est un chiffre
        'Entr�e en modification ==> On met le texte en rouge
        TextTrafCUM.ForeColor = QBColor(12)
    End If
End Sub

Private Sub TextTrafCUM_LostFocus()
    'Test pour �viter les lostfocus en cascade
    If maTrafIniErreur Or maDurServErreur Or Not (Screen.ActiveForm Is Me) Then Exit Sub
    
    'Calcul du trafic initial
    If CalculerTraficIni(Me) Then
        'Contenu valide ==> Mise en noir
        TextTrafCUM.ForeColor = QBColor(0)
        'Sauvegarde de la saisie valide dans le tag
        'pour la touche echappement
        TextTrafCUM.Tag = TextTrafCUM.Text
        TextTrafIni.Tag = TextTrafIni.Text
        maTrafCumErreur = False
        monChangeEvent = False
        TextTrafIni.ForeColor = QBColor(0) 'Pas de modif
        'Mise � jour du trafic cumul� dans la frame R�sultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
        'Calcul du Nombre d'essieux �quivalents si trafic cumul� connu
        CalculerEtAfficherNE
    Else
        'Contenu invalide ==> Mise en rouge
        TextTrafCUM.ForeColor = QBColor(12)
        TabData.Tab = OngletTrafic 'Retour � l'onglet Trafic
        TextTrafCUM.SetFocus
        maTrafCumErreur = True
    End If
End Sub

Private Sub TextTrafIni_Change()
    Dim unAjoutEnFinChaine As Boolean
    Dim uneSelStartDeb As Integer
    
    monAffichageErreurNE = True
    monChangeEvent = True
    'Modif du contenu ==> Mise en rouge
    TextTrafIni.ForeColor = QBColor(12)
    
    'Suppression du caract�re s�parateur millier si on enl�ve le premier
    'chiffre car format traf ini = #,### donc "2 888" peut devenir " 888"
    '==> Plantage
    If Mid(TextTrafIni.Text, 1, 1) = DonnerSepMillier Then
        'Suppression du premier caract�re qui est le s�parateur de millier
        TextTrafIni.Text = Mid(TextTrafIni.Text, 2)
    End If
    
    'Remise � inconnu du trafic cumul� dans la frame R�sultats
    LabelTraficCum.Caption = LabelTraficCumCaption + MsgInconnu
    
    'R�cup de la point d'insertion initial
    uneSelStartDeb = TextTrafIni.SelStart
    
    'Test si on rajoute des caract�res en fin de chaine
    unAjoutEnFinChaine = (TextTrafIni.SelStart = Len(TextTrafIni))
    TextTrafIni.Text = Format(TextTrafIni.Text, "#,###")
    If unAjoutEnFinChaine Then
        TextTrafIni.SelStart = Len(TextTrafIni.Text)
    Else
        TextTrafIni.SelStart = uneSelStartDeb
    End If
End Sub


Private Sub TextTrafIni_KeyPress(KeyAscii As Integer)
    'V�rification qu'un chiffre est tap� et rien d'autre
    '8 = backspace
    
    'si on tape retour (=13), on fait comme un Tab,
    'on passe � la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If KeyAscii = 27 Then TextTrafIni.Text = TextTrafIni.Tag
        'Fin de modification ==> On remet le texte en noir
        TextTrafIni.ForeColor = QBColor(0)
        'Activation du controle suivant
        If TextTrafIni.Text <> "" Then
            RichTextAide.SetFocus
        Else
            TextTrafIni.Text = "0"
            TextTrafCUM.Text = ""
        End If
        Exit Sub
    End If
    
    If (KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57 Then
        Beep
        KeyAscii = 0 'Par annuler la frappe de ce non-chiffre
    Else
        'Cas o� la touche tap�e au clavier est un chiffre
        'Entr�e en modification ==> On met le texte en rouge
        monChangeEvent = True
        TextTrafIni.ForeColor = QBColor(12)
    End If
End Sub

Private Sub TextTrafIni_LostFocus()
    'Test pour �viter les lostfocus en cascade
    If maDurServErreur Or maTrafCumErreur Or Not (Screen.ActiveForm Is Me) Or monChangeEvent = False Then Exit Sub
    
    If VerifierMinMaxTraficIni(Me, TextTrafIni.Text) Then
        'Contenu valide ==> Mise en noir
        TextTrafIni.ForeColor = QBColor(0)
        monChangeEvent = False
        'Sauvegarde de la saisie valide dans le tag
        'pour la touche echappement
        TextTrafIni.Tag = TextTrafIni.Text
        CalculerTraficCum Me
        TextTrafCUM.Tag = TextTrafCUM.Text
        TextTrafCUM.ForeColor = QBColor(0) 'Pas de modif
        'Mise � jour du trafic cumul� dans la frame R�sultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
        'Passage au champ de saisie suivant
        TextDur�eS.SetFocus
        maTrafIniErreur = False
    Else
        'Contenu invalide ==> Mise en rouge
        TextTrafIni.ForeColor = QBColor(12)
        TabData.Tab = OngletTrafic 'Retour � l'onglet Trafic
        TextTrafIni.SetFocus
        maTrafIniErreur = True
    End If
End Sub


Public Sub CalculerEtAfficherNE()
    'Calcul du Nombre d'essieux �quivalents si trafic cumul� connu
    'et recherche des �paisseurs si tout est d�fini pour cela
    Dim uneStruct As Structure
    Dim unMsg As String, unNEEquiv As Long
    
    Set uneStruct = Nothing
    'Calcul et Affichage du Nombre d'essieux �quivalents
    'si trafic cumul� connu
    If Val(TextTrafCUM.Text) > 0 And MaskCAM.Text <> "" Then
        'NE = NPL * CAM
        unNEEquiv = CLng(CLng(TextTrafCUM.Text) * CSng(MaskCAM.Text))
        If unNEEquiv = monNEEquiv Then ' And (TabData.Tab = OngletTrafic Or TabData.Tab = OngletCAM) Then
            Exit Sub
        Else
            monNEEquiv = unNEEquiv
        End If
        
        If monNEEquiv > 10000000 Then
            LabelNEequiv.ForeColor = QBColor(12)
            LabelNEequiv.Caption = LabelNEequivCaption + MsgNEHorsLimite
            If monAffichageErreurNE Then
                monAffichageErreurNE = False
                MsgBox MsgNED�passement1 + Chr(13) + Chr(13) + MsgNED�passement2, vbInformation
            End If
        Else
            LabelNEequiv.ForeColor = QBColor(0)
            LabelNEequiv.Caption = LabelNEequivCaption + Format(monNEEquiv, "### ### ###")
            'R�cup de la structure choisie
            Set uneStruct = DonnerStructChoisie(Me)
            'Test si le NE calcul� est entre les bornes min et max
            'de la structure choisie �ventuelle
            If (uneStruct Is Nothing) = False Then
                If monNEEquiv < uneStruct.monNbEssieuxMin Then
                    'Cas o� le NE est inf�rieur au min de la structure choisie
                    If monAffichageErreurNE Then
                        monAffichageErreurNE = False
                        unMsg = MsgNEInfMin1 + Chr(13) + Chr(13)
                        unMsg = unMsg + MsgNECalcul� + " = " + Format(monNEEquiv, "## ### ###") + "    NE min = " + Format(uneStruct.monNbEssieuxMin, "## ### ###")
                        unMsg = unMsg + "    NE max = " + Format(uneStruct.monNbEssieuxMax, "## ### ###")
                        unMsg = unMsg + Chr(13) + Chr(13) + MsgNEInfMin2
                        MsgBox unMsg, vbInformation
                    End If
                ElseIf monNEEquiv > uneStruct.monNbEssieuxMax Then
                    'Cas o� le NE est sup�rieur au max de la structure choisie
                    If monAffichageErreurNE Then
                        monAffichageErreurNE = False
                        unMsg = MsgNESupMax1 + Chr(13) + Chr(13)
                        unMsg = unMsg + MsgNECalcul� + " = " + Format(monNEEquiv, "## ### ###") + "    NE min = " + Format(uneStruct.monNbEssieuxMin, "## ### ###")
                        unMsg = unMsg + "    NE max = " + Format(uneStruct.monNbEssieuxMax, "## ### ###")
                        unMsg = unMsg + Chr(13) + Chr(13) + MsgNESupMax2
                        MsgBox unMsg, vbInformation
                    End If
                End If
            End If
        End If
    End If
    
    'Recherche des �paisseurs si tout est d�fini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        monMinTecQ1 = "Non"
        monMaxPraQ1 = "Non"
        monMinTecQ2 = "Non"
        monMaxPraQ2 = "Non"
        monMSComp1Q1 = ""
        monMSComp2Q1 = ""
        monMSComp1Q2 = ""
        monMSComp2Q2 = ""
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre � jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Public Sub ActualiserOngletLabelCAM()
    'Mise � jour de l'onglet CAM qui met aussi � jour le label CAM
    'de la frame r�sultat
    Dim uneValPrec As Single, uneValMin As Single, uneValMax As Single
        
    'R�cup du s�parateur d�cimale . ou ,
    'fix� dans les param�tres r�gionaux de Windows
    TrouverCaract�reD�cimalUtilis�
    
    'Affectation du bon mask pour la zone de saisie CAM
    MaskCAM.Mask = "#" + monCarDeci + "##"
    
    uneString = DonnerPrecMinMaxCAM(Me, uneValPrec, uneValMin, uneValMax)
    FrameInfoValCAM.Caption = "Pour une " + uneString + " : "
    LabelValPrecCAM = LabelValPrecCAMCaption + Format(uneValPrec)
    LabelValMinCAM = LabelValMinCAMCaption + Format(uneValMin)
    LabelValMaxCAM = LabelValMaxCAMCaption + Format(uneValMax)
    'Affichage du CAM par d�faut ou celui de l'�tude
    If MaskCAM.Text = (" " + monCarDeci + "  ") Then
        'Cas de d�part, nouvelle �tude (CAM non renseign�
        'donc vide = 0 + s�parateur d�cimale + deux blancs
        'avec le mask fix� au d�but de cette fonction
        '==> mettre la valeur pr�conis�e
        MaskCAM.Text = Format(uneValPrec, "fixed")
        'Stockage dans le tag pour un annuler saisie
        'par la touche Echap
        MaskCAM.Tag = MaskCAM.Text
    Else
        MaskCAM.Text = Format(CSng(MaskCAM.Text), "fixed")
        'Ainsi en cas de changement de caract�re d�cimal
        '==> affichage CAM OK
        'Stockage dans le tag pour un annuler saisie
        'par la touche Echap
        MaskCAM.Tag = MaskCAM.Text
    End If
End Sub

Public Function DonnerEpaisseurTotale(uneQualit� As Byte)
    'Fonction retournant l'�paisseur total suivant
    'la qualit� pass�e en param�tres (1 ou 2)
    Dim uneStruct As Structure, uneSurfSansEpaisseur As Boolean
    
    If uneQualit� <> 1 And uneQualit� <> 2 Then
        DonnerEpaisseurTotale = 55
        MsgBox "ERREUR de Programmation dans frmDocument:DonnerEpaisseurTotale, Qualit� inconnue", vbCritical
        Exit Function
    End If
    
    'Affectation des libell�s des diff�rentes couches
    Set uneStruct = DonnerStructChoisie(Me)
        
    'Recherche si on a une structure avec une couche de surface sans �paisseur
    uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
    
    'Calcul de l'�paisseur totale de la qualit� uneQualit� :
    '(1 � 6 pour Q1 et 7 � 12 pour Q2 dans le tableau des �paisseurs)
    'D'abord l'�paisseur de couche de surface partie 1
    DonnerEpaisseurTotale = monTabEp(1 + 6 * (uneQualit� - 1)) * Abs(Not uneSurfSansEpaisseur)
    'Les autres �paisseurs
    For i = 2 To 6
        DonnerEpaisseurTotale = DonnerEpaisseurTotale + monTabEp(i + 6 * (uneQualit� - 1))
    Next i
End Function

Private Sub AfficherCSurfUneQualite()
    'Masquage des �paisseurs pr�conis�es et composition Q1 ou Q2
    'suivant le choix de l'utilisateur
    LabelEpPrecQ1.Visible = (monTypeChantier = TypeChantierQ1)
    LabelValEpPrecQ1.Visible = (monTypeChantier = TypeChantierQ1)
    LabelCompChoisieQ1.Visible = (monTypeChantier = TypeChantierQ1)
    ComboCompQ1.Visible = (monTypeChantier = TypeChantierQ1)
    LabelQ1cm.Visible = (monTypeChantier = TypeChantierQ1)
    
    If monTypeEtude = TypeEtudeGiratoire Then
        'En �tude giratoire, on enl�ve la mention � Q1
        LabelEpPrecQ1.Caption = Mid(LabelEpPrecQ1Caption, 1, Len(LabelEpPrecQ1Caption) - 4) + ":"
        LabelCompChoisieQ1.Caption = Mid(LabelCompChoisieQ1Caption, 1, Len(LabelCompChoisieQ1Caption) - 4) + ":"
    Else
        LabelEpPrecQ1.Caption = LabelEpPrecQ1Caption
        LabelCompChoisieQ1.Caption = LabelCompChoisieQ1Caption
    End If
    
    LabelEpPrecQ2.Visible = (monTypeChantier = TypeChantierQ2)
    LabelValEpPrecQ2.Visible = (monTypeChantier = TypeChantierQ2)
    LabelCompChoisieQ2.Visible = (monTypeChantier = TypeChantierQ2)
    ComboCompQ2.Visible = (monTypeChantier = TypeChantierQ2)
    LabelQ2cm.Visible = (monTypeChantier = TypeChantierQ2)
    
    'Centrage des �paisseurs et des compositions
    ComboCompQ1.Top = TabData.TabHeight * 3 + ComboCompQ1.Height / 2 - LabelValEpPrecQ1.Height
    LabelValEpPrecQ1.Top = TabData.TabHeight * 3 - ComboCompQ1.Height + LabelValEpPrecQ1.Height
    LabelEpPrecQ1.Top = TabData.TabHeight * 3 - LabelEpPrecQ1.Height / 2
    LabelCompChoisieQ1.Top = TabData.TabHeight * 3 - LabelCompChoisieQ1.Height / 2
    LabelQ1cm.Top = LabelEpPrecQ1.Top + LabelEpPrecQ1.Height / 2
    
    ComboCompQ2.Top = TabData.TabHeight * 3 + ComboCompQ2.Height / 2 - LabelValEpPrecQ2.Height
    LabelValEpPrecQ2.Top = TabData.TabHeight * 3 - ComboCompQ2.Height + LabelValEpPrecQ2.Height
    LabelEpPrecQ2.Top = TabData.TabHeight * 3 - LabelEpPrecQ2.Height / 2
    LabelCompChoisieQ2.Top = TabData.TabHeight * 3 - LabelCompChoisieQ2.Height / 2
    LabelQ2cm.Top = LabelEpPrecQ2.Top + LabelEpPrecQ2.Height / 2
End Sub
