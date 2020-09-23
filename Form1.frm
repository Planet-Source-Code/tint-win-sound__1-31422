VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWinSound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Sound Editor"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   4380
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTestSetting 
      Caption         =   "Test Setting"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CmdsSelectSetting 
      Caption         =   "Select Wave"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ComSelect 
      Left            =   480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.wav|*.wav"
      Orientation     =   2
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset setting"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox label1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save Settings"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "System Startup"
      Top             =   60
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Setting"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmWinSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ApiRegistry As New ClsApiRegistry
Dim a As String
Dim Wich_System_File As String
Public Function Chech_Combo()

Select Case Combo1.Text
    Case "System Exit"
        Wich_System_File = "SystemExit"
    Case "System Startup"
        Wich_System_File = "SystemStart"
    Case "System Exclamation Point"
        Wich_System_File = "SystemExclamation"
    Case "System Asterisk"
        Wich_System_File = "systemAsterisk"
    Case "System Question Mark"
        Wich_System_File = "SystemQuestion"
    Case "Minimize"
        Wich_System_File = "Minimize"
    Case "Maximize"
        Wich_System_File = "Maximize"
    Case "Open"
        Wich_System_File = "Open"
    Case "Close"
        Wich_System_File = "Close"
End Select

End Function

Private Sub cmdReset_Click()
Chech_Combo
Let a = ApiRegistry.GetKeyValue(&H80000001, "AppEvents\Schemes\Apps\.default\" & Wich_System_File & "\.default\", "")
ApiRegistry.UpdateKey &H80000001, "AppEvents\Schemes\Apps\.default\" & Wich_System_File & "\.current\", "", a
End Sub

Private Sub CmdSave_Click()
Chech_Combo
a = label1.Text
ApiRegistry.UpdateKey &H80000001, "AppEvents\Schemes\Apps\.default\" & Wich_System_File & "\.current\", "", a
End Sub

Private Sub CmdsSelectSetting_Click()
Dim result As String
'ComSelect.Filter = "*.WAV"
ComSelect.ShowOpen
result = ComSelect.FileName
label1.Text = result
End Sub

Private Sub Command1_Click()
Chech_Combo
a = ApiRegistry.GetKeyValue(&H80000001, "AppEvents\Schemes\Apps\.default\" & Wich_System_File & "\.current\", "")
label1.Text = a

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub CmdTestSetting_Click()
Select Case Combo1.Text
    Case "System Exit"
        MsgBox "Cant test System Exit", vbInformation, "Testing..."
    Case "System Startup"
        MsgBox "Can't test System Startup", vbInformation, "Testing..."
    Case "System Asterisk"
        MsgBox "System Asterisk...", vbInformation, "Testing..."
    Case "System Exclamation Point"
        MsgBox "Exclamation Point...", vbExclamation, "Testing..."
    Case "System Question Mark"
        MsgBox "Question Mark...", vbQuestion, "Testing..."
    Case "Open"
        MsgBox "Can't test Open", vbInformation, "Testing..."
    Case "Close"
        MsgBox "Can't test Close", vbInformation, "Testing..."
    Case "Minimize"
        MsgBox "Can't test Minimize", vbInformation, "Testing..."
    Case "Maximize"
        MsgBox "Can't test Maximize", vbInformation, "Testing..."
End Select
End Sub

Private Sub Form_Load()

Combo1.AddItem "System Asterisk"
Combo1.AddItem "System Exit"
Combo1.AddItem "System Startup"
Combo1.AddItem "System Exclamation Point"
Combo1.AddItem "System Question Mark"
Combo1.AddItem "Minimize"
Combo1.AddItem "Maximize"
Combo1.AddItem "Open"
Combo1.AddItem "Close"
End Sub
