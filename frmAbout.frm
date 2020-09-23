VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3000
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   4545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070.653
   ScaleMode       =   0  'User
   ScaleWidth      =   4267.989
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3465
      TabIndex        =   3
      Top             =   2160
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Height          =   48
      Left            =   225
      TabIndex        =   2
      Top             =   1980
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "sanjay@kattimani.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   990
      MouseIcon       =   "frmAbout.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "E-mail me"
      Top             =   1395
      Width           =   2070
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   225
      Picture         =   "frmAbout.frx":0594
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2250
      TabIndex        =   9
      Top             =   1665
      Width           =   1995
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "This is a freeware,  You can distribute && use its for any purpose, withought any paymnent or royalty."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   225
      TabIndex        =   8
      Top             =   2160
      Width           =   3075
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit my web site for screen savers && other free down loads."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   975
      TabIndex        =   7
      Top             =   495
      Width           =   3090
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.kattimani.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   990
      MouseIcon       =   "frmAbout.frx":089E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Go to my website"
      Top             =   1035
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Sanjay kattimani.  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1395
      TabIndex        =   4
      Top             =   195
      Width           =   2565
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Website : "
      Height          =   240
      Left            =   270
      TabIndex        =   1
      Top             =   1065
      Width           =   690
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   0
      Top             =   1395
      Width           =   690
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
lblVer.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Label2_Click()
    Call ShellExecute(hwnd, "open", "http://sanjay-kattimani.webjump.com", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Label4_Click()
    Call ShellExecute(hwnd, "open", "mailto:sanjaykattimani@hotmail.com", vbNullString, vbNullString, SW_NORMAL)
End Sub
