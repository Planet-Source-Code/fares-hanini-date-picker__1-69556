VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   0  'None
   Caption         =   "Calendrier"
   ClientHeight    =   2535
   ClientLeft      =   3360
   ClientTop       =   2940
   ClientWidth     =   3465
   ClipControls    =   0   'False
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Height          =   255
      Left            =   3240
      TabIndex        =   56
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   0
      ScaleHeight     =   2370
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   2640
         X2              =   2640
         Y1              =   2205
         Y2              =   840
      End
      Begin VB.Label Sem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   62
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Sem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   61
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Sem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   60
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Sem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   59
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Sem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   58
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Sem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   57
         Top             =   840
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   2370
         Left            =   0
         Top             =   0
         Width           =   3075
      End
      Begin VB.Image imgClose 
         Height          =   240
         Left            =   2700
         Picture         =   "frmDate.frx":000C
         Top             =   120
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2640
         X2              =   240
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblMois 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novembre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   615
         TabIndex        =   55
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   53
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   840
         TabIndex        =   52
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   1200
         TabIndex        =   51
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   1560
         TabIndex        =   50
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   2280
         TabIndex        =   48
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   480
         TabIndex        =   46
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   840
         TabIndex        =   45
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   10
         Left            =   1200
         TabIndex        =   44
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   11
         Left            =   1560
         TabIndex        =   43
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   12
         Left            =   1920
         TabIndex        =   42
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   13
         Left            =   2280
         TabIndex        =   41
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   15
         Left            =   480
         TabIndex        =   39
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   16
         Left            =   840
         TabIndex        =   38
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   17
         Left            =   1200
         TabIndex        =   37
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   18
         Left            =   1560
         TabIndex        =   36
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   19
         Left            =   1920
         TabIndex        =   35
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   20
         Left            =   2280
         TabIndex        =   34
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   21
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   22
         Left            =   480
         TabIndex        =   32
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   24
         Left            =   1200
         TabIndex        =   30
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   25
         Left            =   1560
         TabIndex        =   29
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   26
         Left            =   1920
         TabIndex        =   28
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   28
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   29
         Left            =   480
         TabIndex        =   25
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   30
         Left            =   840
         TabIndex        =   24
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   31
         Left            =   1200
         TabIndex        =   23
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   32
         Left            =   1560
         TabIndex        =   22
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   33
         Left            =   1920
         TabIndex        =   21
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   34
         Left            =   2280
         TabIndex        =   20
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ma"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   17
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Je"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   16
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   15
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   14
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Lib 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Di"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   13
         Top             =   555
         Width           =   375
      End
      Begin VB.Label lblAnnée 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1997"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1650
         TabIndex        =   12
         Top             =   120
         Width           =   495
      End
      Begin VB.Label MoisG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   270
      End
      Begin VB.Label MoisD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   270
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   35
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   36
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   37
         Left            =   840
         TabIndex        =   7
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   38
         Left            =   1200
         TabIndex        =   6
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   39
         Left            =   1560
         TabIndex        =   5
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   40
         Left            =   1920
         TabIndex        =   4
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   41
         Left            =   2280
         TabIndex        =   3
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label AnnéeD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2400
         TabIndex        =   2
         Top             =   120
         Width           =   270
      End
      Begin VB.Label AnnéeG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   1
         Top             =   120
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   30
         Top             =   30
         Width           =   3030
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   23
         Left            =   840
         TabIndex        =   31
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   27
         Left            =   2280
         TabIndex        =   27
         Top             =   1560
         Width           =   360
      End
      Begin VB.Shape picNow 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   2520
         Shape           =   2  'Oval
         Top             =   1560
         Width           =   360
      End
      Begin VB.Shape picSel 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   210
         Left            =   2520
         Shape           =   2  'Oval
         Top             =   1800
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   0
      Picture         =   "frmDate.frx":0596
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMove 
      Height          =   240
      Left            =   0
      Picture         =   "frmDate.frx":0B20
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   0
      Picture         =   "frmDate.frx":10AA
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************private variables
Private PrevButton As Long
Private PrevColor As Long
Private jj As Integer
Private mm As Integer
Private aa As Integer
Private wDate As Date
Private Dayof(12) As Integer

Private Sub Mois(Signe As String)
    Select Case Signe
    Case "+":
        mm = mm + 1
        If mm = 13 Then
            mm = 1
            aa = aa + 1
        End If
    Case "-":
        mm = mm - 1
        If mm = 0 Then
            mm = 12
            aa = aa - 1
        End If
    End Select
    Redraw
End Sub

Private Sub Année(Signe As String)
    Select Case Signe
    Case "+": aa = aa + 1
    Case "-": aa = aa - 1
    End Select
    Redraw
End Sub

Private Sub RefreshCmds()
    MoisG.ForeColor = QBColor(0)
    MoisD.ForeColor = QBColor(0)
    AnnéeG.ForeColor = QBColor(0)
    AnnéeD.ForeColor = QBColor(0)
    lblMois.ForeColor = QBColor(0)
    lblAnnée.ForeColor = QBColor(0)
    Label2(PrevButton).BorderStyle = 0
    picSel.Visible = False
    imgClose.Picture = imgUp.Picture
    picSel.FillStyle = 1
End Sub

Private Sub MoveMouse(Index)
    If Index <> PrevButton Then
        On Error Resume Next
        Label2(PrevButton).BorderStyle = 0
    End If
    PrevButton = Index
    'Label2(Index).BorderStyle = 1
    picSel.Left = Label2(Index).Left
    picSel.Top = Label2(Index).Top
    picSel.FillStyle = 1
    picSel.Visible = True
End Sub

Private Sub Redraw()
    Dayof(1) = IIf(aa Mod 4 = 0, 29, 28)
    If mm = 2 And jj > Dayof(1) Then jj = Dayof(1)
    wDate = CDate(Format(LTrim(Str(jj)) + "/" + LTrim(Str(mm)) + "/" + LTrim(Str(aa)), fmtDate))
    s = Format(Str(mm) + "/1997", "mmmm"): Mid(s, 1, 1) = UCase(Mid(s, 1, 1))
    lblMois = s
    lblAnnée = LTrim(Str(aa))
    Dim i As Integer
    For i = 0 To 41
        Label2(i).ForeColor = QBColor(0)
        Label2(i) = ""
    Next i
    Dim fst As Integer
    '****
    fst = Weekday("01/" + Str(Month(wDate)) + "/" + Str(Year(wDate))) - 2
    If fst < 0 Then
        fst = 6
    End If
    For i = 1 To Dayof(mm - 1)
        Label2(fst + i - 1) = Str(i)
        If i = Day(wDate) Then
            'Label2(fst + i - 1).ForeColor = QBColor(12)
            picNow.Left = Label2(fst + i - 1).Left
            picNow.Top = Label2(fst + i - 1).Top
        End If
    Next i
    Dim X As Integer
    If mm > 1 Then
        X = mm - 1
    Else
        X = 0
    End If
    For i = fst - 1 To 0 Step -1
        Label2(i) = Str(Dayof(X) - (fst - 1 - i))
        Label2(i).ForeColor = QBColor(8)
    Next i
    For i = fst + Dayof(mm - 1) To 41
        Label2(i) = Str(i - fst - Dayof(mm - 1) + 1)
        Label2(i).ForeColor = QBColor(8)
    Next i
    
'        Sem(0) = Format(Label2(0) & "/" & mm & "/" & aa, "ww")
    For i = 0 To 5
        Sem(i) = ""
    Next i
End Sub

Private Sub AnnéeD_Click()
    Année "+"
End Sub

Private Sub AnnéeD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
    AnnéeD.ForeColor = QBColor(12)
End Sub

Private Sub AnnéeG_Click()
    Année "-"
End Sub

Private Sub AnnéeG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
    AnnéeG.ForeColor = QBColor(12)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case Asc("+"):          Mois "+"
    Case Asc("-"):          Mois "-"
    Case Asc("*"):          Année "+"
    Case Asc("/"):          Année "-"
    End Select
End Sub

Private Sub Form_Load()
    For i = 0 To 6
        Lib(i).Caption = Left(Format((22 + i) & "/10/07", "ddd"), 2) & "."
    Next i
    Dim X As Long
    Dim Y As Long
    GetMousePos X, Y
    Me.Top = Y * 15
    Me.Left = IIf(Screen.Width - 3075 > X * 15, X * 15, X * 15 - 3075)
    Me.Width = 3075
    Me.Height = 2370
    jj = Day(Date): mm = Month(Date): aa = Year(Date)
    Dayof(0) = 31: Dayof(1) = 28: Dayof(2) = 31: Dayof(3) = 30
    Dayof(4) = 31: Dayof(5) = 30: Dayof(6) = 31: Dayof(7) = 31
    Dayof(8) = 30: Dayof(9) = 31: Dayof(10) = 30: Dayof(11) = 31
    Redraw
        
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoisG.ForeColor = QBColor(0)
    MoisD.ForeColor = QBColor(0)
    AnnéeG.ForeColor = QBColor(0)
    AnnéeD.ForeColor = QBColor(0)
    RefreshCmds
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgDown.Picture
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgMove.Picture
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgUp.Picture
    Unload Me
End Sub

Private Sub lblAnnée_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Année IIf(Button = vbRightButton, "+", "-")
End Sub

Private Sub lblMois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
    lblMois.ForeColor = QBColor(12)
End Sub

Private Sub Label2_Click(Index As Integer)
'    Dim dd As String
    Dim X As Integer
    X = mm
    Y = aa
    If Label2(Index).ForeColor = QBColor(8) Then
      If Label2(Index).Caption > 15 Then
        If mm > 2 Then
            X = mm - 1
        Else
            X = 12
            Y = aa - 1
        End If
      Else
        If mm < 12 Then
            X = mm + 1
        Else
            X = 1
            Y = aa + 1
        End If
      End If
    End If
    strDate = LTrim(Label2(Index).Caption) + "/" + LTrim(Str(X)) + "/" + LTrim(Str(Y))
    Unload Me
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  PrevColor = Label2(Index).BackColor
  'Label2(Index).BackColor = QBColor(12)
  picSel.FillStyle = 0
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveMouse Index
End Sub

Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label2(Index).BackColor = PrevColor
    picSel.FillStyle = 1
End Sub

Private Sub lblAnnée_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
    lblAnnée.ForeColor = QBColor(12)
End Sub

Private Sub lblMois_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mois IIf(Button = vbRightButton, "+", "-")
End Sub

Private Sub Lib_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RefreshCmds
End Sub

Private Sub MoisD_Click()
    Mois "+"
End Sub

Private Sub MoisD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
    MoisD.ForeColor = QBColor(12)
End Sub

Private Sub MoisG_Click()
    Mois "-"
End Sub

Private Sub MoisG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
    MoisG.ForeColor = QBColor(12)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RefreshCmds
End Sub
