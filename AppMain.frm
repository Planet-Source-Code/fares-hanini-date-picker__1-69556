VERSION 5.00
Begin VB.Form AppMain 
   Caption         =   "Date Picker"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks for all users of PSC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   6495
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "lblHelp"
      Height          =   1335
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.devocom.com"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   780
      Width           =   855
   End
End
Attribute VB_Name = "AppMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDate_Click()
        frmDate.Show 1
        If strDate <> "" Then
            txtDate = Format(strDate, "dd/mm/yyyy")
            strDate = ""
        End If
End Sub

Private Sub Form_Load()
    txtDate = Format(Date, "dd/mm/yyyy")
    lblHelp = "* Press F2 key or the button to popup a date picker"
    lblHelp = lblHelp & vbLf & "* Press ESC to unload the date picker"
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'if you press F2 key in date textbox
    If KeyCode = vbKeyF2 Then
        frmDate.Show 1
        If strDate <> "" Then
            txtDate = Format(strDate, "dd/mm/yyyy")
            strDate = ""
        End If
        KeyCode = 0
    End If
End Sub

