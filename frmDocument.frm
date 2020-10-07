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
   Begin VB.Frame FrameRésultat 
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
            Caption         =   "Q1 <---- Indice de Gel en °C.J ----> Q2"
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
            Caption         =   " <---- Référence Corrigé ----> "
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
         Begin VB.Label LabelQualité 
            AutoSize        =   -1  'True
            Caption         =   "Q1 <----- Qualité -----> Q2"
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
         Caption         =   "Trafic Cumulé : "
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
      Tab(1).Control(8)=   "TextDuréeS"
      Tab(1).Control(9)=   "LabelTrafCUM"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "LabelTrafIni"
      Tab(1).Control(12)=   "LabelCroissA"
      Tab(1).Control(13)=   "Label5"
      Tab(1).Control(14)=   "LabelDuréeS"
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
            Caption         =   "Pavée ou Dallée"
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
            Caption         =   "Béton"
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
         Caption         =   "Couche de forme non gélive : "
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
            Caption         =   "0,14 (Traité)"
            Height          =   195
            Left            =   2940
            TabIndex        =   64
            Top             =   360
            Width           =   1180
         End
         Begin VB.OptionButton OptionANT 
            Caption         =   "0,12 (Non Traité)"
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
            Caption         =   "Non gélif"
            Height          =   195
            Left            =   60
            TabIndex        =   60
            Top             =   780
            Width           =   1095
         End
         Begin VB.OptionButton OptionPGel 
            Caption         =   "Peu gélif"
            Height          =   195
            Left            =   60
            TabIndex        =   59
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton OptionTGel 
            Caption         =   "Trés gélif"
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
         Caption         =   "Hiver de Référence"
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
            Caption         =   "°C.J"
            Height          =   195
            Left            =   1800
            TabIndex        =   102
            Top             =   720
            Width           =   285
         End
         Begin VB.Label LabAltUnit 
            AutoSize        =   -1  'True
            Caption         =   "mètres"
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
            Caption         =   "H Agglomération : "
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label LabGel 
            AutoSize        =   -1  'True
            Caption         =   "Station Réf : "
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   930
         End
         Begin VB.Label LabelHStatUnit 
            AutoSize        =   -1  'True
            Caption         =   "mètres"
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
            Caption         =   "Valeur préconisée du CAM  = "
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
         Caption         =   "Informations Matériau couche de surface : "
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
         Caption         =   "Informations Matériau de la couche de fondation : "
         Height          =   375
         Left            =   -72970
         TabIndex        =   36
         Top             =   2760
         Width           =   4260
      End
      Begin VB.CommandButton CmdInfoMB 
         Caption         =   "Informations Matériau de la couche de base : "
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
      Begin VB.TextBox TextDuréeS 
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
            Caption         =   "Voie réservée aux bus"
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
         Caption         =   "Type d'aménagement :"
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
         Caption         =   "Coefficient d'Agressivité Moyen : "
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
         Caption         =   "Epaisseur préconisée Q2 :"
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
         Caption         =   "Cliquez sur un matériau pour avoir ses caractéristiques"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   -71280
         TabIndex        =   81
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label LabelListMatSurfComp 
         Caption         =   "Matériaux composants possibles pour cette épaisseur préconisée :"
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
         Caption         =   "Epaisseur préconisée Q1 :"
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
         Caption         =   "Trafic cumulé de PL sur la durée de service :"
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
         Caption         =   "Trafic initial à la mise en service :"
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
      Begin VB.Label LabelDuréeS 
         AutoSize        =   -1  'True
         Caption         =   "Durée de service :"
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
         Caption         =   "Titre de l'étude :"
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
         Caption         =   "Dernière modification : "
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
'Variable stockant l'épaisseur totale maximale entre les carottes Q1 et Q2
Private monEpTotMaxReel As Integer
    
'Variable indiquant si toutes les saisies sont finies
'ou valides dans cette fenêtre
Public maFinSaisie As Byte 'valeur initiale = 0

'Variable indiquant si une erreur s'est produit
'lors d'un lost focus dans l'une des TextBox de l'onglet Trafic
'Initialisé à FAUX
Private maTrafIniErreur As Boolean
Private maDurServErreur As Boolean
Private maTrafCumErreur As Boolean

'Variable indiquant si on veut juste d'ouvrir l'étude
Public monJustOpen As Boolean 'Initialisé à False par VB

'Variable indiquant si un change event de textbox
'(traf ini, traf cum ou duréee service) est survenue
'Initialisation par défaut FALSE
Public monChangeEvent As Boolean

'Variable indiquant s'il faut afficher le message d'erreur
'lors du calcul du NE lorqu'une erreur se produit
Public monAffichageErreurNE As Boolean

'Variables stockant les champs d'une étude
'Initialement elles valent toutes 0
Public monTypeEtude As Byte
Public monTypeChantier As Byte
Public monTypeVoie As Integer
Public maDate As String
Public monTitreEtude As String
Public maVariante As String
Public monTraficIni As Integer
Public monTraficCumulé As Long
Public maDuréeService As Byte
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

'Variable indiquant si on a trouvé les épaisseurs des qualités Q1 et Q2
Public monEpQ1Trouv As Boolean
Public monEpQ2Trouv As Boolean

'Variable indiquant les Qm des qualités Q1 et Q2
Public monQmQ1 As Single
Public monQmQ2 As Single

'Variable indiquant si les min techno et max pratiques sont atteints
'pour les qualités Q1 et Q2 (valeurs possibles "Oui" ou "Non")
Public monMinTecQ1 As String
Public monMaxPraQ1 As String
Public monMinTecQ2 As String
Public monMaxPraQ2 As String
    
'Tableau contenant les épaisseurs des couches de la structure choisie
'De 1 à 6 ==> Epaisseurs pour la qualité Q1
'1 et 2 épaisseur couche surface, 3 et 4 épaisseur couche base, 5 et 6 épaisseur couche fondation
'De 7 à 12 ==> Epaisseurs pour la qualité Q2
'7 et 8 épaisseur couche surface, 9 et 10 épaisseur couche base, 11 et 12 épaisseur couche fondation
Public monTabEp As Variant

'Variable contenant les matériaux composants
'du matériau de surface composé (deux maximuns) pour qualité Q1 et Q2
Public monMSComp1Q1 As String
Public monMSComp2Q1 As String
Public monMSComp1Q2 As String
Public monMSComp2Q2 As String
'Variables de l'épaisseur préconisé trouvé pour Q1 et Q2
Public monEpPrecQ1 As Integer
Public monEpPrecQ2 As Integer
'Variables contenant l'indice de la composition choisie
'pour le matériau de surface composé pour Q1 et Q2
Public monIndCompQ1 As Integer
Public monIndCompQ2 As Integer

'Variables concernant la vérif au gel
Public monIndHiver As Integer
Public monIndStation As Integer
Public monHAgglo As Integer
Public monIndTailleAgglo As Integer
Public monCoefA As Single
Public monEpNonGel As Integer
Public maPente As String
Public monIndGelSol As Integer

'Variables indiquant une modif dans l'étude
Public maModif As Boolean

'Variables pour indiquer le premier activate de la fenêtre Etude = frmDocument
Private unNbActivate As Byte 'valeur par défaut = 0
    

Private Sub BtnGelQ1_Click()
    MsgBox "Cette chaussée n'est pas protégée au gel en condition de chantier standard (qualité Q1)," + Chr(13) + Chr(13) + "car l'Indice de Gel de Référence > Indice de Gel Admissible pour Q1.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
End Sub

Private Sub BtnGelQ2_Click()
    MsgBox "Cette chaussée n'est pas protégée au gel en condition de chantier difficile (qualité Q2)," + Chr(13) + Chr(13) + "car l'Indice de Gel de Référence > Indice de Gel Admissible pour Q2.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
End Sub

Private Sub BtnInfoStruct_Click()
    'Affichage de la fiche commentaire de la structure choisie
    Dim uneStruct As Structure
    
    'Recup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Chargement sans affichage
    Load frmInfoStruct
    
    'Remplissage de frmInfoStruct
    frmInfoStruct.Caption = "Informations sur " + uneStruct.monAbrégé
    frmInfoStruct.RichTextInfo.TextRTF = uneStruct.monComment
    
    'Centrage de la fiche matériau
    CentrerFenetreEcran frmInfoStruct
    
    'Affichage modal
    frmInfoStruct.Show vbModal
End Sub

Private Sub BtnOKGel_Click()
    If OptionChoixQ1.Value Then
        MsgBox "Cette chaussée est protégée au gel en condition de chantier standard (qualité Q1)," + Chr(13) + Chr(13) + "car l'Indice de Gel de Référence <= Indice de Gel Admissible pour Q1.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
    End If
    If OptionChoixQ2.Value Then
        MsgBox "Cette chaussée est protégée au gel en condition de chantier difficile (qualité Q2)," + Chr(13) + Chr(13) + "car l'Indice de Gel de Référence <= Indice de Gel Admissible pour Q2.", vbMsgBoxHelpButton + vbInformation, , App.HelpFile, IDhlp_VerifGel
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
    Dim uneVisibilité As Boolean
    
    maModif = True 'Indication de changement pour la sauvegarde
    
    uneVisibilité = (CheckIndGelPerso.Value = 1)
    monUtilIndGelPerso = CheckIndGelPerso.Value
    
    'Cas où l'on coche la case c'est vrai
    'Affichage si vrai ou masquage si faux
    'de la zone de saisie de l'indice de gel personnel
    LabelIndGelPerso.Visible = uneVisibilité
    TextIndGelPerso.Visible = uneVisibilité
    LabelCJ.Visible = uneVisibilité
        
    'Mise à jour des valeurs d'indices de gel de référence
    If uneVisibilité Then
        TextIndGelPerso.Text = Format(monIndGelPerso)
        TextIndGelPerso.ForeColor = QBColor(0)
        Me.monIndiceGelRefQ1 = monIndGelPerso
        Me.monIndiceGelRefQ2 = monIndGelPerso
    Else
        'Recalcul et affichage des indices de gel de références
        'Q1 et Q2 à partir de la station de référence
        CalculerIndiceGelRef Me
    End If
    
    'Mise en grisée si vrai ou dégrisée si faux
    'de la frame de choix des hivers de référene
    FrameHiver.Enabled = Not uneVisibilité
    OptionHE.Enabled = Not uneVisibilité
    OptionHC.Enabled = Not uneVisibilité
    OptionHRNE.Enabled = Not uneVisibilité
    
    'Affichage si vrai, masquage si faux
    'de la zone de saisie de l'indice de gel personnel
    LabGel.Visible = Not uneVisibilité
    ComboStation.Visible = Not uneVisibilité
    LabelHStation.Visible = Not uneVisibilité
    HStation.Visible = Not uneVisibilité
    LabelHStatUnit.Visible = Not uneVisibilité
    LabAlt.Visible = False 'Not uneVisibilité (car Hagglo plus utilisée)
    TextHAgglo.Visible = False 'Not uneVisibilité (car Hagglo plus utilisée)
    LabAltUnit.Visible = False 'Not uneVisibilité (car Hagglo plus utilisée)
    LabAgglo.Visible = Not uneVisibilité
    ComboTailleAgglo.Visible = Not uneVisibilité

    'Affichage dans la frame résultat de l'étude active
    ActualiserFrameVerifGel Me
End Sub

Private Sub CmdInfoMB_Click()
    Dim uneStruct As Structure
    
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Affichage de la fiche caractéristique du matériau
    'Paramètre string = TypeMat + "/" + Abrégé du Matériau
    AfficherFicheMat "FondBase" + "/" + uneStruct.maCoucheBase
End Sub

Private Sub CmdInfoMF_Click()
    Dim uneStruct As Structure
    
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Affichage de la fiche caractéristique du matériau
    'Paramètre string = TypeMat + "/" + Abrégé du Matériau
    AfficherFicheMat "FondBase" + "/" + uneStruct.maCoucheFondation
End Sub

Private Sub cmdInfoMS_Click()
    Dim unTypeMat As String
    Dim uneStruct As Structure
    
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    
    'Récup du type de matériau en couche de surface
    If TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatSimple Then
        unTypeMat = "Simple"
    ElseIf TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé Then
        unTypeMat = "Composé"
    Else
        MsgBox MsgErreurProg + MsgErreurMatériauInconnu + MsgIn + "frmDocument:cmdInfoMS_Click", vbCritical
    End If
        
    'Affichage de la fiche caractéristique du matériau
    'Paramètre string = TypeMat + "/" + Abrégé du Matériau
    AfficherFicheMat unTypeMat + "/" + uneStruct.maCoucheSurface
End Sub

Private Sub ComboCompQ1_Click()
    Dim lesComp As Collection
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    If Not (uneStruct Is Nothing) And ComboCompQ1.ListIndex > -1 Then
        unIndComp = ComboCompQ1.ItemData(ComboCompQ1.ListIndex)
        Set lesComp = maColMatSurf(uneStruct.maCoucheSurface).mesCompositions
        'Récup des noms de matériaux composants et de leur épaisseur
        monTabEp(1) = CInt(lesComp(unIndComp + 1))
        monMSComp1Q1 = Format(lesComp(unIndComp + 2))
        monTabEp(2) = CInt(lesComp(unIndComp + 3))
        monMSComp2Q1 = Format(lesComp(unIndComp + 4))
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
        'Si les épaisseurs préconisées Q1 et Q2 sont les mêmes
        'la valeur mise pour Q1 est affecté aussi à Q2 par défaut
        If monEpPrecQ1 = monEpPrecQ2 Then ComboCompQ2.ListIndex = ComboCompQ1.ListIndex
        'Calcul de l'indice de gel admissible
        If mesOptionsGen.maVerifGel Then AfficherEtCalculerIndGelAdm Me
    End If
End Sub


Private Sub ComboCompQ2_Click()
    Dim lesComp As Collection
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(Me)
    If Not (uneStruct Is Nothing) And ComboCompQ2.ListIndex > -1 Then
        unIndComp = ComboCompQ2.ItemData(ComboCompQ2.ListIndex)
        Set lesComp = maColMatSurf(uneStruct.maCoucheSurface).mesCompositions
        'Récup des noms de matériaux composants et de leur épaisseur
        monTabEp(7) = CInt(lesComp(unIndComp + 1))
        monMSComp1Q2 = Format(lesComp(unIndComp + 2))
        monTabEp(8) = CInt(lesComp(unIndComp + 3))
        monMSComp2Q2 = Format(lesComp(unIndComp + 4))
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
        'Calcul de l'indice de gel admissible
        If mesOptionsGen.maVerifGel Then AfficherEtCalculerIndGelAdm Me
    End If
End Sub


Private Sub ComboStation_Click()
    'Affichage de l'altitude de la station de référence choisie
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
    '==> Remise des épaisseurs Q1/Q2 à non trouvées
    'et des composants 1 et 2 en couche de surface à vide
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
        'Remise à vide du CAM
        MaskCAM.Mask = ""
        MaskCAM.Text = ""
        LabelCAM.Caption = LabelCAMCaption + MsgInconnu
        LabelNEequiv.Caption = LabelNEequivCaption + MsgInconnu
    End If
    
    If Val(ComboStruct.Tag) = ComboStruct.ListIndex Then
        'Choix par click de la même structure que celle
        'précédemment sélectionnée
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
    
    'Mise à jour des boutons de visu des matériau surface simple, base
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
    'éventuelle est faite d'un matériau de surface composé
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
        'Affectation de la valeur préconisée du CAM
        ActualiserOngletLabelCAM
        
        'En version 2 si on n'est pas en voie de parking, on grise l'onglet CAM
        'ainsi le CAM a la valeur préconisée poour les voies de parking
        If OptionVoieParking.Value = True Then
            TrouverCaractèreDécimalUtilisé
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
        
        'Calcul et affichage éventuel du NE
        CalculerEtAfficherNE
        
        'Récup du matériau de surface éventuel
        'et test s'il est composé ==> Activation Onglet Couche surface
        If uneStruct.maCoucheSurface = "Aucune" Then
            TabData.TabEnabled(OngletSurf) = False
            cmdInfoMS.Enabled = False
            'On remet à vide les compositions possibles de la couche de surface
            ComboCompQ1.ListIndex = -1
            ComboCompQ2.ListIndex = -1
        Else
            Set unMatSurf = maColMatSurf(uneStruct.maCoucheSurface)
            TabData.TabEnabled(OngletSurf) = (TypeOf unMatSurf Is MatComposé) And (monEpQ1Trouv Or monEpQ2Trouv)
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
        'désactivation des onglets CAM et couche de surface
        TabData.TabEnabled(OngletCAM) = uneStructChoisie
        TabData.TabEnabled(OngletSurf) = uneStructChoisie
        'Inhiber bouton info matériaux couches
        InhiberBoutonMat Me
        'Affichage aide générale
        unFileName = CorrigerNomFichier(App.Path + "\OngletStructure.rtf")
        RichTextAide.LoadFile unFileName, rtfRTF
        DoEvents
    End If

    'Mise à jour de l'affichage des carottes Q1 et Q2
    FrameCarotte.Visible = uneStructChoisie
    If uneStructChoisie Then AfficherCarottes Me
End Sub

Private Sub ComboStruct_DropDown()
    'Modifie par click  dans la liste déroulante de la combobox
    'Stockage de l'indice de la structure sélectionnée
    'On en le fait que si on change de structure dans la liste de structures
    'du même type, si on change de type la liste n'est plus la même, donc on ne
    'stocke pas dans le tag le numéro d'index de la structure choisie dans
    'la liste précédente
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



Private Sub CalculerNEthéorique()
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
    'Vérification que la saisie de la fenêtre précédemment active
    'soit finie et valide
    If maFinSaisie = 0 Then
        If VerifierFinSaisie Then
            'Si OK ==> Changement d'étude active
            'Stockage de l'étude en cours
            Set monEtude = Me
            monEtude.maFinSaisie = 0
        Else
            'On revient dans la fenêtre précédent
            'on la remet au premier plan
            monEtude.maFinSaisie = 1 'Pour ne pas redéclencher la vérif de fin de saisie
            monEtude.WindowState = vbNormal
            monEtude.ZOrder 0
        End If
    Else
        monEtude.maFinSaisie = 0 'Pour pouvoir redéclencher la vérif de fin de saisie
    End If
    'Mise à jour du contexte id sur l'aide de l'onglet actif
    ChangerHelpID TabData.Tab
    
    If unNbActivate = 0 Then
        'Ré-affichage du type d'étude et les autres radio boutons, cause bug au load de la form
        'Pour la version 2 bétatest
        'UNIQUEMENT LORS DU PREMIER ACTVIVATE (cela équivaut au load)
        If monTypeEtude = TypeEtudeStandard Then
            OptionEtudeStandard.Value = True
        ElseIf monTypeEtude = TypeEtudeGiratoire Then
            OptionEtudeGiratoire.Value = True
        Else
            'Cas erreur de programmation ==> on met étude standard
            MsgBox MsgErreurProg + MsgErreurTypeEtudeInconnue + MsgIn + "ModuleMain:MettreAJourOngletVoie", vbCritical
            OptionEtudeStandard.Value = True
        End If
        'Idem type de chantier
        'Par défaut on est en condition Q1
        If monTypeChantier = TypeChantierQ1 Then
            OptionChoixQ1.Value = True
            OptionChoixQ2.Value = False
            LabelQualité.Caption = "" 'en V2 on ne met rien "Chantier standard (Q1)"
        Else
            OptionChoixQ1.Value = False
            OptionChoixQ2.Value = True
            LabelQualité.Caption = "" 'en V2 on ne met rien "Chantier difficile (Qualité Q2)"
        End If
        'Affectation du nbActivate pour ne plus faire ce code là
        unNbActivate = 1
        'Indication pour ne pas enregistrer lors d'une ouverture
        'Au premier affichage rien n'a pu être modifié
        If maNewEtude = False Then maModif = False
    End If
End Sub


Private Sub Form_Initialize()
    'Variable indiquant s'il faut afficher le message d'erreur
    'lors du calcul du NE lorqu'une erreur se produit
    'Initialisé à vrai au départ
    Dim unTypeVoieEtudeChantier As Integer
    
    monAffichageErreurNE = True
            
    If maNewEtude Then
        'Mise à jour des indices dans les combobox
        'dans le cas d'une nouvelle étude car elles vont de 0 à n-1
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
        maCroisAnnuel = 100 'Valeur pour non renseignée
        monIndiceGelRefQ1 = -1
        monIndiceGelRefQ2 = -1
        
        'Pour la version 2 bétatest
        monTypeEtude = TypeEtudeStandard
        monTypeChantier = TypeChantierQ1
        
        'Signal que l'on vient de créer une étude
        monJustOpen = True
    Else
        'Cas d'une ouverture d'une étude existante
        'Récupération dans la collection de lecture de données
        '(cf fonction ModuleFichier:OuvrirEtude + format de fichier *.urb)
        'Données pour l'onglet Voie
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurOuverture
        
        monFichId = Val(maColLectFich(1))
        monTitreEtude = maColLectFich(3)
        
        'Décodage à partir de la version 2, de l'entier lu en quatrième
        'position dans le fichier *.urb
        unTypeVoieEtudeChantier = maColLectFich(4)
        If unTypeVoieEtudeChantier < ChantierDifficile Then
            'Cas d'un chantier standard = qualité Q1
            monTypeChantier = TypeChantierQ1
        Else
            'Cas d'un chantier difficile = qualité Q2
            monTypeChantier = TypeChantierQ2
            'Récupération du type de voie et d'étude en enlevant
            'le rajout forfaitaire de la valeur ChantierDifficile
            unTypeVoieEtudeChantier = unTypeVoieEtudeChantier - ChantierDifficile
        End If
        'Récup du type d'étude
        If unTypeVoieEtudeChantier < TypeGiratoireDistribution Then
            'Type d'étude standard
            monTypeEtude = TypeEtudeStandard
        Else
            'Type d'étude giratoire
            monTypeEtude = TypeEtudeGiratoire
        End If
        'Récup du type de voie
        monTypeVoie = unTypeVoieEtudeChantier
        
        maVariante = maColLectFich(5)
        maDate = maColLectFich(6)
        'Données pour l'onglet Trafic
        If maColLectFich(7) = "" Then
            monTraficIni = 0
        Else
            monTraficIni = CInt(maColLectFich(7))
        End If
        maDuréeService = Val(maColLectFich(8))
        maCroisAnnuel = Val(maColLectFich(9))
        If maColLectFich(10) = "" Then
            monTraficCumulé = 0
        Else
            monTraficCumulé = CLng(maColLectFich(10))
        End If
        'Données pour l'onglet Structure
        monIndStructChoisie = Val(maColLectFich(12))
        monUtilFichPerso = Val(maColLectFich(13))
        monFichPersoSTR = maColLectFich(14)
        'données pour les ongletsCAM et plateforme
        monCAM = maColLectFich(15)
        monIndicePF = Val(maColLectFich(16))
        'Données pour l'onglet Couche de surface
        monIndCompQ1 = Val(maColLectFich(17))
        monIndCompQ2 = Val(maColLectFich(18))
        'Données pour l'onglet Gel
        monIndHiver = Val(maColLectFich(19))
        monIndStation = Val(maColLectFich(20))
        monHAgglo = Val(maColLectFich(21))
        monIndTailleAgglo = Val(maColLectFich(22))
        monIndGelSol = Val(maColLectFich(23))
        maPente = maColLectFich(24)
        monCoefA = CSng(maColLectFich(25))
        monEpNonGel = CInt(maColLectFich(26))
        'Récup éventuelle de l'indice de gel perso si fichier au format de
        'la version finale, aprés les corrections suite aux sites pilotes
        If maColLectFich(28) = FormatFichierVersionFinale Then
            monIndGelPerso = CInt(maColLectFich(29))
            monUtilIndGelPerso = CByte(maColLectFich(30))
        End If
        
        'Signal que l'on vient d'ouvrir l'étude
        monJustOpen = True
        
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
   End If
   
    'Sortie pour éviter la routine de gestion d'erreur
    monOuverture = True 'Pas d'erreur à signaler
    Exit Sub
   
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurOuverture:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + MsgEtude + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    MsgBox unMsg, vbCritical
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    monOuverture = False
End Sub


Private Sub Form_Load()
    Dim uneString As String, unTextCAM As String
    
    'Masquage du titre de la frame carotte donnant les deux qualités
    'en V2 ce titre disparait
    LabelQualité.Visible = False
    
    'Remplissage et positionnement des labelInfo de la frame chantier
    LabelInfo1.Caption = LabelInfo1Caption
    LabelInfo2.Caption = LabelInfo2Caption
    LabelInfo1.Left = OptionChoixQ1.Left
    LabelInfo1.Top = OptionChoixQ1.Top
    LabelInfo2.Left = OptionChoixQ2.Left
    LabelInfo2.Top = OptionChoixQ2.Top
    BtnPlusInfo.Top = LabelInfo2.Top + (LabelInfo2.Height - BtnPlusInfo.Height) / 2
    BtnPlusInfo.Left = LabelInfo2.Left + LabelInfo2.Width + 60
       
    'Masquage des info sur l'altitude de l'agglo (pas utilisée)
    LabAlt.Visible = False
    TextHAgglo.Visible = False
    LabAltUnit.Visible = False
    'Centrage des info sur l'altitude de la station de référence
    'même position que l'indice de gel perso
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
    
    'Décalage des fenêtres de 330 twpis en X et Y à chaque fenêtre
    Top = (Forms.Count - 2) * 330
    Left = (Forms.Count - 2) * 330
    fMainForm.Arrange vbCascade
    
    If maNewEtude = False Then
        'Mise à jour du titre de la fenêtre de l'étude
        Caption = maColLectFich(27)
        ViderCollection maColLectFich
    End If
    
    'Tableau contenant les épaisseurs des couches de la structure choisie
    'Pour indice = 0 on prend 2 cm pour afficher l'entête de la carotte en impression
    'De 1 à 6 ==> Epaisseurs pour la qualité Q1
    '1 et 2 épaisseur couche surface, 3 et 4 épaisseur couche base, 5 et 6 épaisseur couche fondation
    'De 7 à 12 ==> Epaisseurs pour la qualité Q2
    '7 et 8 épaisseur couche surface, 9 et 10 épaisseur couche base, 11 et 12 épaisseur couche fondation
    monTabEp = Array(2, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0, EpParDefaut, 0)
    
    'Mise à jour des boutons dans la toolbar permettant l'impression
    'et la sauvegarde car il n'y a pas de fenêtre fille ouverte
    '==> Impression et sauvegarde impossible
    fMainForm.tbToolBar.Buttons("Print").Visible = True
    fMainForm.tbToolBar.Buttons("Save").Visible = True
    
    'Calcul de la largeur de la zone aide rtf
    RichTextAide.RightMargin = RichTextAide.Width - 500
    'Initialisation de l'aide contextuelle d'onglet
    unFileName = CorrigerNomFichier(App.Path + "\OngletVoie.rtf")
    RichTextAide.LoadFile unFileName, rtfRTF
            
    'Mise à jour des couleurs des couches des carottes Q1 et Q2
    'en prenant celles des options générales
    ChangerCouleurCouches Me
    
    'Mise à jour de l'affichage de la fenêtre
    If AfficherTypeVoie(monTypeVoie) Then
        'Cas sans erreur à l'affichage du type de voie
        
        'Mise à jour de l'affichage des différents onglets
        MettreAJourOngletVoie Me
        MettreAJourOngletTrafic Me
                
        'Récup du séparateur décimale . ou ,
        'fixé dans les paramètres régionaux de Windows
        TrouverCaractèreDécimalUtilisé
        
        'Mise à jour de l'onglet CAM (deux chiffres après la virgule)
        If monCAM <> "" Then
            'Affectation du bon mask pour la zone de saisie CAM
            MaskCAM.Mask = "#" + monCarDeci + "##"
            'MaskCAM.Text = Format(monCAM, "fixed")
            'bug si stockage de la valeur en . et
            'si séparateur décimale actuel = virgule
            MaskCAM.Text = Mid(monCAM, 1, 1) + monCarDeci + Mid(monCAM, 3)
            'On prend de part et d'autre du séparateur décimale
            'car le cam a un format #.##
        Else
            MaskCAM.Text = ""
        End If
        
        MettreAJourOngletStructure Me, monUtilFichPerso, monIndStructChoisie

        'Mise à jour de l'onglet plateforme
        MettreAjourOngletPlateForme Me
        
        'Mise à jour de l'onglet gel
        MettreAJourOngletGel Me
        
        'Mise à jour de l'affichage de la frame Résultat
        MettreAJourFrameRésultat Me
        
        'Mise à jour de la frame de vérif au gel
        ActualiserFrameVerifGel Me
    Else
        'Cas avec erreur à l'affichage du type de voie
        '==> on indique que la fenêtre est à fermer
        'pour les fonctions LoadNew et ouverture d'une étude
        Tag = "A_Fermer"
    End If
    
    'Indication qu'aucun change event de textbox n'a été déclenchée
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
        'quand on ferme la dernière fenêtre fille ouverte, on est
        'dans le unload donc elle existe encore et il reste la
        'fenetre MDI mère
        '==> Lors de la Fermeture dernière fille, Nombre de Forms = 2
        
        'Mise à jour des boutons dans la toolbar permettant l'impression
        'et la sauvegarde s'il n'y a plus de fenêtre fille ouverte
        '==> Impression et sauvegarde impossible
        fMainForm.tbToolBar.Buttons("Print").Visible = False
        fMainForm.tbToolBar.Buttons("Save").Visible = False
        
        'On remet le contexte id à 0 car plus de fenêtres filles ouvertes
        fMainForm.HelpContextID = 0
    End If
    
    Close #monFichId
    Set monEtude = Nothing
End Sub

Public Function AfficherChoixTypesVoies() As Boolean
    'Affiche les types de voie dans l'onglet Voie suivant le type d'étude
    'Standard ou Giratoire
    If OptionEtudeStandard.Value Then
        monTypeEtude = TypeEtudeStandard
        'Affichage des données en étude standard (= section courante)
        FrameVoie.Visible = True
        OptionChoixQ1.Visible = True
        OptionChoixQ2.Visible = True
        'Masquage des données en étude giratoire
        FrameGiratoire.Visible = False
        LabelInfo1.Visible = False
        LabelInfo2.Visible = False
        BtnPlusInfo.Visible = False
        'La plateforme PF1 est possible
        OptionPF1.Enabled = True
    ElseIf OptionEtudeGiratoire.Value Then
        monTypeEtude = TypeEtudeGiratoire
        'Masquage des données en étude standard (= section courante)
        FrameVoie.Visible = False
        OptionChoixQ1.Visible = False
        OptionChoixQ2.Visible = False
        'On remet condition de chantier à Q1, car tout est remis à zéro
        OptionChoixQ1.Value = True
        'Masquage des données en étude giratoire
        FrameGiratoire.Visible = True
        LabelInfo1.Visible = True
        LabelInfo2.Visible = True
        BtnPlusInfo.Visible = True
        'La plateforme PF1 n'est pas possible
        OptionPF1.Enabled = False
        If OptionPF1.Value Then
            'Si PF1 était la plateforme choisie, on se met en PF2
            OptionPF2.Value = True
        End If
    Else
        'Cas erreur de programmation ==> on met étude standard
        MsgBox MsgErreurProg + MsgErreurTypeEtudeInconnue + MsgIn + "frmDocument:AfficherChoixTypesVoies", vbCritical
        FrameVoie.Visible = True
        FrameGiratoire.Visible = False
    End If

    LabelQualité.Caption = "" 'en V2 on enlève le titre de la frame carottes
    'On met les voies giratoires au même endroit que les voies
    FrameGiratoire.Left = FrameVoie.Left
    FrameGiratoire.Top = FrameVoie.Top
    
    If monFichId = 0 Or unNbActivate > 0 Then
        'On fait ce traitement si c'est une nouvelle étude (monFichId = 0,
        'sinon valeur >= 1 issue du fichier urb) et si ce n'est pas le premier
        'appel à la fonction AfficherChoixTypesVoies lors d'une ouverture d'une
        'étude issue d'un ficher urb, cet appel se déclenche au premier activate
        'de la form, donc pour unNbActivate = 0
        
        'Remise à vide des boutons sélectionnés
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
        'Mettre à jour affichage de l'indice de gel admissible et référence
        'monIndiceGelAdmQ1 = 0
        'monIndiceGelAdmQ2 = 0
        'monIndiceGelRefQ1 = -1
        'monIndiceGelRefQ2 = -1
        'ActualiserFrameVerifGel Me
    End If
End Function

Public Function AfficherTypeVoie(unTypeVoie As Integer) As Boolean
    'Affiche le type de voie dans la fenêtre résultat de gauche
    'et dans l'ongle Voie à droite et la valeur préconisée pour
    'ce type de voie pour le CAM dans l'onglet CAM
    'Retourne Vrai si aucune erreur, faux sinon
    Dim uneString As String
    
    'Initialisation à vide du trafic initial et du trafic cumulé
    'Utile lors d'un changement de type de voies
    TextTrafIni.Text = ""
    TextTrafCUM.Text = ""
    
    'Initialisation à vide du CAM
    'la bonne valeur est mis après dans le Form_Load
    MaskCAM.Mask = ""
    MaskCAM.Text = ""
    LabelCAM.Caption = LabelCAMCaption + MsgInconnu
    
    'Initialisation à vide du NE
    LabelNEequiv.ForeColor = QBColor(0)
    LabelNEequiv.Caption = LabelNEequivCaption + MsgInconnu
    
    'Initialisation de la valeur de retour
    AfficherTypeVoie = True
    
    'Sélection du bon bouton option pour le type de voies
    'et Mise à jour des valeurs possible et par défaut des
    'hivers dans l'onglet Gel si c'est une nouvelle étude
    If unTypeVoie = TypeVoieInconnu Then
        uneString = MsgInconnu
        'Pas d'hiver par défaut
        OptionHE.Value = False
        OptionHRNE.Value = False
        OptionHC.Value = False
        'Indice de gel de référence inconnu
        monIndiceGelRefQ1 = -1
        monIndiceGelRefQ2 = -1
        'On considère que c'est comme une nouvelle étude
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
        'Cas erreur de programmation ==> on sort de la fenêtre fille
        MsgBox MsgErreurProg + MsgErreurTypeVoieInconnu + MsgIn + "frmDocument:AfficherTypeVoie", vbCritical
        AfficherTypeVoie = False
    End If
    
    LabelTypeVoie.Caption = LabelTypeVoieCaption + uneString
    
    'Mise à jour du grisé des onglets ci-dessous
    'enable si untypevoie > à 0
    TabData.TabEnabled(OngletTrafic) = (unTypeVoie > 0)
    TabData.TabEnabled(OngletStruct) = (unTypeVoie > 0)
    TabData.TabEnabled(OngletPF) = (unTypeVoie > 0)
    TabData.TabEnabled(OngletGel) = (unTypeVoie > 0) And mesOptionsGen.maVerifGel
       
    'Mise à jour du type de voie de l'etude active
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
    'correction d'un bug V1 en V2, après ouvir etude, puis changement de voie,
    'l'onglet CAM est encore actif et le click sur cet onglet CAM fait planter
    'Struct-urb version 1 d'où la correction de la ligne ci-dessous
    'TabData.TabEnabled(OngletCAM) = (monIndStructChoisie > 0)
    TabData.TabEnabled(OngletCAM) = (ComboStruct.ListIndex >= 0)
    TabData.TabEnabled(OngletSurf) = False
    
    'On enlève les carottes visibles
    FrameCarotte.Visible = False
    
    'Inhibition du bouton info structure
    BtnInfoStruct.Enabled = False
    
    'Mettre à jour affichage de l'indice de gel admissible
    monIndiceGelAdmQ1 = 0
    monIndiceGelAdmQ2 = 0
    ActualiserFrameVerifGel Me
    
    'Remise à zéro du NE equivalent
    monNEEquiv = 0
    
    'En version 2, si le type de voie est parking
    'on inhibe l'onglet Trafic
    'et on prend TraficCumulé = 100 000, on mettra CAM = 0.1 lors du choix
    'de la structure
    'pour avoir un NE = 10 000, seul disponible pour les
    'structures de voie de parking
    If unTypeVoie = TypeVoieParking Then
        TabData.TabEnabled(OngletTrafic) = False
        TabData.TabEnabled(OngletCAM) = False
        Me.TextTrafCUM.Text = "100 000"
        'Mise à jour du trafic cumulé dans la frame Résultats
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
    'Affichage de la fiche caractéristique du matériau
    'Paramètre string = TypeMat + "/" + Abrégé du Matériau
    AfficherFicheMat "Composant" + "/" + ListViewMat.SelectedItem.Text
 End Sub


Private Sub MaskCAM_Change()
    If MaskCAM.Tag <> MaskCAM.Text Then monAffichageErreurNE = True
    LabelCAM.Caption = LabelCAMCaption + MaskCAM.Text
End Sub

Private Sub MaskCAM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        'Début de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    End If
End Sub

Private Sub MaskCAM_KeyPress(KeyAscii As Integer)
    TrouverCaractèreDécimalUtilisé
    
    If KeyAscii = 27 Then  'Touche Escape
        'Restauration dernière valeur valide
        MaskCAM.Text = MaskCAM.Tag
        'Changer le control actif pour déclencher le LostFocus de MaskCAM
        RichTextAide.SetFocus
    ElseIf KeyAscii = 13 Then 'touche Return
        'Changer le control actif pour déclencher le LostFocus de MaskCAM
        RichTextAide.SetFocus
    ElseIf KeyAscii = 46 And monCarDeci = "," Then
        'Touche . tapée or caractère décimale = virgule, on remplace
        KeyAscii = 44 '44 = virgule
        'Début de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    ElseIf KeyAscii = 44 And monCarDeci = "." Then
        'Touche , tapée or caractère décimale = point, on remplace
        KeyAscii = 46 '46 = point
        'Début de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        'Cas où l'on tape un chiffre entre 0 et 9 ou la touche baskspace
        'Début de modification ==> mise en rouge de la zone de saisie
        MaskCAM.ForeColor = QBColor(12)
    End If
End Sub

Private Sub MaskCAM_LostFocus()
    If DonnerStructChoisie(Me) Is Nothing Or MaskCAM.Text = "" Then Exit Sub
    
    If VerifierMinMaxCAM(Me, MaskCAM.Text) Then
        TrouverCaractèreDécimalUtilisé
        MaskCAM.Mask = "#" + monCarDeci + "##"
        MaskCAM.Text = Format(MaskCAM.Text, "fixed")
        'Fin de modif ==> Remise de la zone de saisie en noir
        MaskCAM.ForeColor = QBColor(0)
        'Stockage dans la dernière valeur valide
        MaskCAM.Tag = MaskCAM.Text
        'Calcul du Nombre d'essieux équivalents si trafic cumulé connu
        CalculerEtAfficherNE
    Else
        'Cas d'erreur ==> on retourne dans MaskCAM
        TabData.Tab = OngletCAM 'Retour à l'onglet CAM
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
    'Mise à jour du trafic cumulé dans la frame Résultats
    LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
    TextTrafCUM.ForeColor = QBColor(0)
End Sub

Private Sub OptionChoixQ1_Click()
    monTypeChantier = TypeChantierQ1
    LabelQualité.Caption = "" 'en V2 on ne met plus rien "Chantier standard (Q1)"
    AfficherCarottes Me
    ActualiserFrameVerifGel Me
    AfficherCSurfUneQualite
    'Indication d'un changement pour déclencher le save as
    maModif = True
End Sub

Private Sub OptionChoixQ2_Click()
    monTypeChantier = TypeChantierQ2
    LabelQualité.Caption = "" 'en V2 on ne met plus rien "Chantier difficile (Q2)"
    AfficherCarottes Me
    ActualiserFrameVerifGel Me
    AfficherCSurfUneQualite
    'Indication d'un changement pour déclencher le save as
    maModif = True
    'Message d'avertissement pour travailler en qualité Q2 ou chantier difficile
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
    'Recherche des épaisseurs si tout est défini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre à jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPF2_Click()
    LabelPFQ1.Caption = "PF2"
    LabelPFQ2.Caption = "PF2"
    'Recherche des épaisseurs si tout est défini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre à jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPF2Plus_Click()
    LabelPFQ1.Caption = "PF2+"
    LabelPFQ2.Caption = "PF2+"
    'Recherche des épaisseurs si tout est défini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre à jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Private Sub OptionPF3_Click()
    LabelPFQ1.Caption = "PF3"
    LabelPFQ2.Caption = "PF3"
    'Recherche des épaisseurs si tout est défini pour cela
    If TrouverEpaisseurPossible(Me) Then
        RechercherEpaisseur Me
    Else
        monEpQ1Trouv = False
        monEpQ2Trouv = False
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre à jour l'onglet couche de surface
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
    'L'index a la même valeur que les constantes de type
    'de structures
    'Const Souple As Byte = 1
    'Const Bitumineuse As Byte = 2
    'Const GTLH As Byte = 3
    'Const Beton As Byte = 4
    'Const Mixte As Byte = 5
    'Const PavesDalles As Byte = 6
    
    'Inhibition du bouton info structure
    BtnInfoStruct.Enabled = False
    'Inhibition des boutons d'info sur les différents matériaux de structure
    InhiberBoutonMat Me
    'Inhibition de l'onglet CAM et couche de surface
    TabData.TabEnabled(OngletCAM) = False
    TabData.TabEnabled(OngletSurf) = False
    
    'Mettre à jour affichage de l'indice de gel admissible
    monIndiceGelAdmQ1 = 0
    monIndiceGelAdmQ2 = 0
    ActualiserFrameVerifGel Me
    'Masquage du taux de risque car plus de structure sélectionnée
    LabelRisk.Visible = False
    'On enlève les carottes visibles
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
    
    'Mise à jour du contexte pour l'aide avec F1
    ChangerHelpID TabData.Tab
    
    'Mise de l'aide contextuelle RTF
    Set uneStruct = DonnerStructChoisie(Me)
    unFileName = CorrigerNomFichier(App.Path + "\Onglet")
    RichTextAide.LoadFile unFileName + TabData.TabCaption(TabData.Tab) + ".rtf", rtfRTF
    'Récup du séparateur décimale . ou ,
    'fixé dans les paramètres régionaux de Windows
    TrouverCaractèreDécimalUtilisé
    
    If TabData.Tab = OngletTrafic Then
        'Si le CAM est valide (pas en rouge), le trafic ini prend le focus
        If MaskCAM.ForeColor <> QBColor(12) Then
            TextTrafIni.SetFocus
            TextTrafIni.SelStart = Len(TextTrafIni.Text)
        End If
    ElseIf TabData.Tab = OngletCAM Then
        'Mise à jour de l'onglet CAM qui met aussi à jour le label CAM
        'de la frame résultat
        ActualiserOngletLabelCAM
        MaskCAM.SetFocus
        MaskCAM.SelStart = Len(MaskCAM.Text)
    ElseIf TabData.Tab = OngletGel Then
        'Mise à jour de l'onglet Gel
        'Affichage avec le bon caractère décimal des
        'valeurs possibles du coefficient A
        OptionANT.Caption = "0" + monCarDeci + "12 (Non Traité)"
        OptionAT.Caption = "0" + monCarDeci + "14 (Traité)"
        If monUtilIndGelPerso = 0 Then
            ComboStation.SetFocus
            'Mise en commentaires des lignes ci-dessous car HAgglo pas utilisée
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


Private Sub TextDuréeS_Change()
    Dim unAjoutEnFinChaine As Boolean
    Dim uneSelStartDeb As Integer
                
    monAffichageErreurNE = True
    monChangeEvent = True
    'Modif du contenu ==> Mise en rouge
    TextDuréeS.ForeColor = QBColor(12)
    
    'Remise à inconnu du trafic cumulé dans la frame Résultats
    LabelTraficCum.Caption = LabelTraficCumCaption + MsgInconnu
    
    If Val(TextDuréeS.Text) = 0 Then
        TextDuréeS.Text = "0"
        Exit Sub
    End If
    
    'Récup de la point d'insertion initial
    uneSelStartDeb = TextDuréeS.SelStart
    
    'Test si on rajoute des caractères en fin de chaine
    unAjoutEnFinChaine = (TextDuréeS.SelStart = Len(TextDuréeS))
    TextDuréeS.Text = Format(TextDuréeS.Text, "##")
    If unAjoutEnFinChaine Then
        TextDuréeS.SelStart = Len(TextDuréeS.Text)
    Else
        TextDuréeS.SelStart = uneSelStartDeb
    End If
End Sub


Private Sub TextDuréeS_KeyPress(KeyAscii As Integer)
    'Vérification qu'un chiffre est tapé et rien d'autre
    '8 = backspace
    
    'si on tape retour (=13), on fait comme un Tab,
    'on passe à la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant
    If KeyAscii = 13 Or KeyAscii = 27 Then
        If KeyAscii = 27 Then TextDuréeS.Text = TextDuréeS.Tag
        'Fin de modification ==> On remet le texte en noir
        TextDuréeS.ForeColor = QBColor(0)
        'Activation du controle suivant
        If TextDuréeS.Text <> "" Then RichTextAide.SetFocus
        Exit Sub
    End If
    
    If (KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57 Then
        Beep
        KeyAscii = 0 'Par annuler la frappe de ce non-chiffre
    Else
        'Cas où la touche tapée au clavier est un chiffre
        'Entrée en modification ==> On met le texte en rouge
        TextDuréeS.ForeColor = QBColor(12)
        monChangeEvent = True
    End If
End Sub

Private Sub TextDuréeS_LostFocus()
    'Test pour éviter les lostfocus en cascade
    If maTrafIniErreur Or maTrafCumErreur Or Not (Screen.ActiveForm Is Me) Or monChangeEvent = False Then Exit Sub

    If VerifierMinMaxDuréeService(Me) Then
        'Contenu valide ==> Mise en noir
        TextDuréeS.ForeColor = QBColor(0)
        monChangeEvent = False
        'Sauvegarde de la saisie valide dans le tag
        'pour la touche echappement
        TextDuréeS.Tag = TextDuréeS.Text
        CalculerTraficCum Me
        maDurServErreur = False
        'Mise à jour du trafic cumulé dans la frame Résultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
        TextTrafCUM.ForeColor = QBColor(0)
    Else
        'Contenu invalide ==> Mise en rouge
        TextDuréeS.ForeColor = QBColor(12)
        TabData.Tab = OngletTrafic 'Retour à l'onglet Trafic
        TextDuréeS.SetFocus
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
    'on passe à la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant et on passe à la text box suivante
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
    'on passe à la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant et on passe à la text box suivante
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
    'Calcul de l'indice de gel corrigé
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
    'on passe à la text box suivante et si on fait Echappement(=27)
    'on remet la valeur d'avant et on passe à la text box suivante
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
    'Calcul de l'indice de gel corrigé
    monIndGelPerso = Format(TextIndGelPerso.Text)
    monIndiceGelRefQ1 = Format(TextIndGelPerso.Text)
    monIndiceGelRefQ2 = Format(TextIndGelPerso.Text)
    'Affichage dans la frame résultat de l'étude active
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
    
    'Début de modification ==> mise en rouge de la zone de saisie
    TextPente.ForeColor = QBColor(12)
    
    If IsNumeric(TextPente.Text) And InStr(1, TextPente.Text, "-") = 0 And InStr(1, TextPente.Text, "+") = 0 Then
        'Les instr recherchant + et - sont rajoutés car la fonction
        'IsNumeric considère qu'une chaine les contenant est numérique
        'même si le + et le - sont n'importe où dans la chaine
        maModif = True 'Indication pour l'enregistrer
    Else
        If TextPente.Text <> "" And TextPente.Text <> monCarDeci Then
            MsgBox MsgSaisieRéelPositif, vbCritical
            unePos = TextPente.SelStart
            TextPente.Text = Mid(TextPente.Text, 1, unePos - 1) + Mid(TextPente.Text, unePos + 1)
            TextPente.SelStart = unePos - 1
        End If
    End If
End Sub

Private Sub TextPente_KeyPress(KeyAscii As Integer)
    'Récup du séparateur décimale . ou ,
    'fixé dans les paramètres régionaux de Windows
    TrouverCaractèreDécimalUtilisé
    
    If KeyAscii = 27 Then  'Touche Escape
        'Restauration dernière valeur valide
        TextPente.Text = TextPente.Tag
        'Changer le control actif pour déclencher le LostFocus de c
        TextEpaisseur.SetFocus
    ElseIf KeyAscii = 13 Then 'touche Return
        'Changer le control actif pour déclencher le LostFocus de TextPente
        TextEpaisseur.SetFocus
    ElseIf KeyAscii = 46 And monCarDeci = "," Then
        'Touche . tapée or caractère décimale = virgule, on remplace
        KeyAscii = 44 '44 = virgule
    ElseIf KeyAscii = 44 And monCarDeci = "." Then
        'Touche , tapée or caractère décimale = point, on remplace
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
    'Sélection du bon type de sol support d'après la pente
    If TextPente.Text = "" Then
    ElseIf TextPente.Text = MsgInfinie Then
        OptionTGel.Tag = "PasDeModif" 'Evite la remise à jour dans Option.click
        OptionTGel.Value = True
        OptionTGel.Tag = ""
    ElseIf CSng(TextPente.Text) > 0.4 Then
        OptionTGel.Tag = "PasDeModif" 'Evite la remise à jour dans Option.click
        OptionTGel.Value = True
        OptionTGel.Tag = ""
    ElseIf CSng(TextPente.Text) <= 0.05 Then
        OptionNGel.Tag = "PasDeModif" 'Evite la remise à jour dans Option.click
        OptionNGel.Value = True
        OptionNGel.Tag = ""
    Else
        OptionPGel.Tag = "PasDeModif" 'Evite la remise à jour dans Option.click
        OptionPGel.Value = True
        OptionPGel.Tag = ""
    End If
End Sub



Private Sub TextTrafCUM_Change()
    Dim unAjoutEnFinChaine As Boolean
    Dim uneSelStartDeb As Integer
    
    'Suppression du caractère séparateur millier si on enlève le premier
    'chiffre car format traf cum = ###,###,###
    'donc "200 456 888" peut devenir " 456 888" ou  " 888"
    '==> trafcum.text devient vide d'où la correction
    If Mid(TextTrafCUM.Text, 1, 1) = DonnerSepMillier Then
        'Suppression du premier caractère qui est le séparateur de millier
        TextTrafCUM.Text = Mid(TextTrafCUM.Text, 2)
    End If
    
    'Modif du contenu ==> Mise en rouge
    If Format(TextTrafCUM.Text, "###,###,###") <> TextTrafCUM.Tag Then
        TextTrafCUM.ForeColor = QBColor(12)
        monAffichageErreurNE = True
    End If
    
    'Remise à inconnu du trafic cumulé dans la frame Résultats
    LabelTraficCum.Caption = LabelTraficCumCaption + MsgInconnu
    
    'Récup de la point d'insertion initial
    uneSelStartDeb = TextTrafCUM.SelStart
    
    'Test si on rajoute des caractères en fin de chaine
    unAjoutEnFinChaine = (TextTrafCUM.SelStart = Len(TextTrafCUM))
    TextTrafCUM.Text = Format(TextTrafCUM.Text, "###,###,###")
    If unAjoutEnFinChaine Then
        TextTrafCUM.SelStart = Len(TextTrafCUM.Text)
    Else
        TextTrafCUM.SelStart = uneSelStartDeb
    End If
End Sub


Private Sub TextTrafCUM_KeyPress(KeyAscii As Integer)
    'Vérification qu'un chiffre est tapé et rien d'autre
    '8 = backspace
    
    'si on tape retour (=13), on fait comme un Tab,
    'on passe à la text box suivante et si on fait Echappement(=27)
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
        'Cas où la touche tapée au clavier est un chiffre
        'Entrée en modification ==> On met le texte en rouge
        TextTrafCUM.ForeColor = QBColor(12)
    End If
End Sub

Private Sub TextTrafCUM_LostFocus()
    'Test pour éviter les lostfocus en cascade
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
        'Mise à jour du trafic cumulé dans la frame Résultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
        'Calcul du Nombre d'essieux équivalents si trafic cumulé connu
        CalculerEtAfficherNE
    Else
        'Contenu invalide ==> Mise en rouge
        TextTrafCUM.ForeColor = QBColor(12)
        TabData.Tab = OngletTrafic 'Retour à l'onglet Trafic
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
    
    'Suppression du caractère séparateur millier si on enlève le premier
    'chiffre car format traf ini = #,### donc "2 888" peut devenir " 888"
    '==> Plantage
    If Mid(TextTrafIni.Text, 1, 1) = DonnerSepMillier Then
        'Suppression du premier caractère qui est le séparateur de millier
        TextTrafIni.Text = Mid(TextTrafIni.Text, 2)
    End If
    
    'Remise à inconnu du trafic cumulé dans la frame Résultats
    LabelTraficCum.Caption = LabelTraficCumCaption + MsgInconnu
    
    'Récup de la point d'insertion initial
    uneSelStartDeb = TextTrafIni.SelStart
    
    'Test si on rajoute des caractères en fin de chaine
    unAjoutEnFinChaine = (TextTrafIni.SelStart = Len(TextTrafIni))
    TextTrafIni.Text = Format(TextTrafIni.Text, "#,###")
    If unAjoutEnFinChaine Then
        TextTrafIni.SelStart = Len(TextTrafIni.Text)
    Else
        TextTrafIni.SelStart = uneSelStartDeb
    End If
End Sub


Private Sub TextTrafIni_KeyPress(KeyAscii As Integer)
    'Vérification qu'un chiffre est tapé et rien d'autre
    '8 = backspace
    
    'si on tape retour (=13), on fait comme un Tab,
    'on passe à la text box suivante et si on fait Echappement(=27)
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
        'Cas où la touche tapée au clavier est un chiffre
        'Entrée en modification ==> On met le texte en rouge
        monChangeEvent = True
        TextTrafIni.ForeColor = QBColor(12)
    End If
End Sub

Private Sub TextTrafIni_LostFocus()
    'Test pour éviter les lostfocus en cascade
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
        'Mise à jour du trafic cumulé dans la frame Résultats
        LabelTraficCum.Caption = LabelTraficCumCaption + TextTrafCUM.Text
        'Passage au champ de saisie suivant
        TextDuréeS.SetFocus
        maTrafIniErreur = False
    Else
        'Contenu invalide ==> Mise en rouge
        TextTrafIni.ForeColor = QBColor(12)
        TabData.Tab = OngletTrafic 'Retour à l'onglet Trafic
        TextTrafIni.SetFocus
        maTrafIniErreur = True
    End If
End Sub


Public Sub CalculerEtAfficherNE()
    'Calcul du Nombre d'essieux équivalents si trafic cumulé connu
    'et recherche des épaisseurs si tout est défini pour cela
    Dim uneStruct As Structure
    Dim unMsg As String, unNEEquiv As Long
    
    Set uneStruct = Nothing
    'Calcul et Affichage du Nombre d'essieux équivalents
    'si trafic cumulé connu
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
                MsgBox MsgNEDépassement1 + Chr(13) + Chr(13) + MsgNEDépassement2, vbInformation
            End If
        Else
            LabelNEequiv.ForeColor = QBColor(0)
            LabelNEequiv.Caption = LabelNEequivCaption + Format(monNEEquiv, "### ### ###")
            'Récup de la structure choisie
            Set uneStruct = DonnerStructChoisie(Me)
            'Test si le NE calculé est entre les bornes min et max
            'de la structure choisie éventuelle
            If (uneStruct Is Nothing) = False Then
                If monNEEquiv < uneStruct.monNbEssieuxMin Then
                    'Cas où le NE est inférieur au min de la structure choisie
                    If monAffichageErreurNE Then
                        monAffichageErreurNE = False
                        unMsg = MsgNEInfMin1 + Chr(13) + Chr(13)
                        unMsg = unMsg + MsgNECalculé + " = " + Format(monNEEquiv, "## ### ###") + "    NE min = " + Format(uneStruct.monNbEssieuxMin, "## ### ###")
                        unMsg = unMsg + "    NE max = " + Format(uneStruct.monNbEssieuxMax, "## ### ###")
                        unMsg = unMsg + Chr(13) + Chr(13) + MsgNEInfMin2
                        MsgBox unMsg, vbInformation
                    End If
                ElseIf monNEEquiv > uneStruct.monNbEssieuxMax Then
                    'Cas où le NE est supérieur au max de la structure choisie
                    If monAffichageErreurNE Then
                        monAffichageErreurNE = False
                        unMsg = MsgNESupMax1 + Chr(13) + Chr(13)
                        unMsg = unMsg + MsgNECalculé + " = " + Format(monNEEquiv, "## ### ###") + "    NE min = " + Format(uneStruct.monNbEssieuxMin, "## ### ###")
                        unMsg = unMsg + "    NE max = " + Format(uneStruct.monNbEssieuxMax, "## ### ###")
                        unMsg = unMsg + Chr(13) + Chr(13) + MsgNESupMax2
                        MsgBox unMsg, vbInformation
                    End If
                End If
            End If
        End If
    End If
    
    'Recherche des épaisseurs si tout est défini pour cela
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
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes Me
    End If
    'Mettre à jour l'onglet couche de surface
    ValiderOngletCoucheSurface Me
End Sub

Public Sub ActualiserOngletLabelCAM()
    'Mise à jour de l'onglet CAM qui met aussi à jour le label CAM
    'de la frame résultat
    Dim uneValPrec As Single, uneValMin As Single, uneValMax As Single
        
    'Récup du séparateur décimale . ou ,
    'fixé dans les paramètres régionaux de Windows
    TrouverCaractèreDécimalUtilisé
    
    'Affectation du bon mask pour la zone de saisie CAM
    MaskCAM.Mask = "#" + monCarDeci + "##"
    
    uneString = DonnerPrecMinMaxCAM(Me, uneValPrec, uneValMin, uneValMax)
    FrameInfoValCAM.Caption = "Pour une " + uneString + " : "
    LabelValPrecCAM = LabelValPrecCAMCaption + Format(uneValPrec)
    LabelValMinCAM = LabelValMinCAMCaption + Format(uneValMin)
    LabelValMaxCAM = LabelValMaxCAMCaption + Format(uneValMax)
    'Affichage du CAM par défaut ou celui de l'étude
    If MaskCAM.Text = (" " + monCarDeci + "  ") Then
        'Cas de départ, nouvelle étude (CAM non renseigné
        'donc vide = 0 + séparateur décimale + deux blancs
        'avec le mask fixé au début de cette fonction
        '==> mettre la valeur préconisée
        MaskCAM.Text = Format(uneValPrec, "fixed")
        'Stockage dans le tag pour un annuler saisie
        'par la touche Echap
        MaskCAM.Tag = MaskCAM.Text
    Else
        MaskCAM.Text = Format(CSng(MaskCAM.Text), "fixed")
        'Ainsi en cas de changement de caractère décimal
        '==> affichage CAM OK
        'Stockage dans le tag pour un annuler saisie
        'par la touche Echap
        MaskCAM.Tag = MaskCAM.Text
    End If
End Sub

Public Function DonnerEpaisseurTotale(uneQualité As Byte)
    'Fonction retournant l'épaisseur total suivant
    'la qualité passée en paramètres (1 ou 2)
    Dim uneStruct As Structure, uneSurfSansEpaisseur As Boolean
    
    If uneQualité <> 1 And uneQualité <> 2 Then
        DonnerEpaisseurTotale = 55
        MsgBox "ERREUR de Programmation dans frmDocument:DonnerEpaisseurTotale, Qualité inconnue", vbCritical
        Exit Function
    End If
    
    'Affectation des libellés des différentes couches
    Set uneStruct = DonnerStructChoisie(Me)
        
    'Recherche si on a une structure avec une couche de surface sans épaisseur
    uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
    
    'Calcul de l'épaisseur totale de la qualité uneQualité :
    '(1 à 6 pour Q1 et 7 à 12 pour Q2 dans le tableau des épaisseurs)
    'D'abord l'épaisseur de couche de surface partie 1
    DonnerEpaisseurTotale = monTabEp(1 + 6 * (uneQualité - 1)) * Abs(Not uneSurfSansEpaisseur)
    'Les autres épaisseurs
    For i = 2 To 6
        DonnerEpaisseurTotale = DonnerEpaisseurTotale + monTabEp(i + 6 * (uneQualité - 1))
    Next i
End Function

Private Sub AfficherCSurfUneQualite()
    'Masquage des épaisseurs préconisées et composition Q1 ou Q2
    'suivant le choix de l'utilisateur
    LabelEpPrecQ1.Visible = (monTypeChantier = TypeChantierQ1)
    LabelValEpPrecQ1.Visible = (monTypeChantier = TypeChantierQ1)
    LabelCompChoisieQ1.Visible = (monTypeChantier = TypeChantierQ1)
    ComboCompQ1.Visible = (monTypeChantier = TypeChantierQ1)
    LabelQ1cm.Visible = (monTypeChantier = TypeChantierQ1)
    
    If monTypeEtude = TypeEtudeGiratoire Then
        'En étude giratoire, on enlève la mention à Q1
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
    
    'Centrage des épaisseurs et des compositions
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
