VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MS agent Text reader"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmReader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlFileOpen 
      Left            =   3645
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelectAgent 
      CausesValidation=   0   'False
      DownPicture     =   "frmReader.frx":030A
      Height          =   510
      Left            =   2925
      Picture         =   "frmReader.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Select MS Agent"
      Top             =   2160
      Width           =   555
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About.."
      Height          =   330
      Left            =   3600
      TabIndex        =   4
      Top             =   945
      Width           =   825
   End
   Begin VB.CheckBox chkAutoClipReader 
      Caption         =   "Auto read clipboard"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   2115
      Width           =   1905
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   540
      Width           =   825
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Default         =   -1  'True
      Height          =   330
      Left            =   3600
      TabIndex        =   1
      Top             =   135
      Width           =   825
   End
   Begin VB.TextBox txtToRead 
      Height          =   1905
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmReader.frx":091E
      Top             =   135
      Width           =   3345
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   225
      Top             =   1485
   End
   Begin VB.Label lblAgent 
      Caption         =   "Agent - Genie"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   2385
      Width           =   2670
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   3735
      Picture         =   "frmReader.frx":0933
      Top             =   1395
      Width           =   540
   End
   Begin AgentObjectsCtl.Agent MSAgent 
      Left            =   2745
      Top             =   1440
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Genie As IAgentCtlCharacterEx
Dim AgentName As String


Private Sub chkAutoClipReader_Click()
If chkAutoClipReader.Value = 1 Then
    Timer1.Enabled = True
    SaveSetting App.EXEName, "settings", "clip", "1"
Else
    Timer1.Enabled = False
    SaveSetting App.EXEName, "settings", "clip", "0"
End If
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdClose_Click()
Genie.Stop
Genie.Hide
DoEvents
Sleep (2)
MSAgent.Characters.Unload AgentName
Unload Me
End Sub

Private Sub cmdRead_Click()
Genie.Speak txtToRead.Text
End Sub

Private Sub cmdSelectAgent_Click()
cdlFileOpen.DialogTitle = "Select the agent."
'cdlFileOpen.FilterIndex =


Dim Filter As String
'On Error GoTo OpenError
Filter = Filter + "MSAgent files (*.acs)|*.acs|"
'Filter = "All files (*.*)|*.*|"
cdlFileOpen.Filter = Filter
cdlFileOpen.FilterIndex = 0
cdlFileOpen.InitDir = "c:\windows\msagent\chars\"
cdlFileOpen.Action = 1
If cdlFileOpen.FileTitle <> "" Then
    SaveSetting "MSAReader", "Settings", "Agent", Left(cdlFileOpen.FileTitle, Len(cdlFileOpen.FileTitle) - 4)
    Genie.Hide
    Sleep (1)
    MSAgent.Characters.Unload AgentName
    Call Form_Load
Else
    'Cancel selected
End If
Exit Sub
OpenError:
MsgBox "You have close Common Dialog ""Open"" with Cancel button!"
Exit Sub
End Sub



Function Sleep(Secs As Integer)
Dim StartTime
StartTime = Timer
Do While Timer <= StartTime + Secs
    DoEvents
Loop
End Function

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
'MSAgent.Characters.Unload "Genie"


'AgentName = "Genie"
AgentName = GetSetting("MSAReader", "Settings", "Agent", "Genie")
lblAgent.Caption = "Agent - " & AgentName
     On Error GoTo 0
      MSAgent.Characters.Load AgentName, AgentName & ".acs"
      Set Genie = MSAgent.Characters(AgentName)
      Genie.LanguageID = &H409
      IsAssistantVisible = True
   
   Genie.Show
'   Genie.MoveTo Me.Left, Me.Left
   If GetSetting(App.EXEName, "settings", "clip", "0") Then
       chkAutoClipReader.Value = 1
   End If
   If Hour(Now()) < 12 Then
    Genie.Speak "Good morning"
   ElseIf Hour(Now()) >= 12 And Hour(Now()) < 17 Then
    Genie.Speak "Good afternoon"
   Else
    Genie.Speak "Good evening"
   End If
   
   
   DoEvents
End Sub

Private Sub Timer1_Timer()
If Clipboard.GetFormat(1) = True Then
If Clipboard.GetText <> "" Then
    txtToRead.Text = Clipboard.GetText
    Genie.Speak txtToRead.Text
    Clipboard.SetText ("")
End If
End If
End Sub
