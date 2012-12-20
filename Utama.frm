VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enigma Tool -  GPL 2"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3315
   Icon            =   "Utama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Utama.frx":15162
   ScaleHeight     =   4590
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      Picture         =   "Utama.frx":168B7
      ScaleHeight     =   975
      ScaleWidth      =   3375
      TabIndex        =   9
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton pushMe 
      Caption         =   "APPLY !"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Computer Option     :  "
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3015
      Begin VB.Frame Frame2 
         Caption         =   "Task manager"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton tskE 
            Caption         =   "Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton tskD 
            Caption         =   "Disable"
            Height          =   255
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "USB Port"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
         Begin VB.OptionButton usbE 
            Caption         =   "Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton usbD 
            Caption         =   "Disable"
            Height          =   255
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Registry"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2775
         Begin VB.OptionButton regD 
            Caption         =   "Disable"
            Height          =   255
            Left            =   1560
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton regE 
            Caption         =   "Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About Me ?"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Form1.Enabled = False
    Form2.Show
End Sub





Private Sub Command2_Click()
Unload Me
End Sub

Private Sub pushMe_Click()
    ''# Task manager - DONE!
    If tskE.Value = True Then
        ''Enable task manager
        A = Shell("REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 0 /f", vbNormalFocus)
        MsgBox "Enabled TASK MANAGER !", vbInformation, "EIO!"
    ElseIf tskD.Value = True Then
        A = Shell("REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 1 /f", vbNormalFocus)
         MsgBox "Disabled TASK MANAGER !", vbInformation, "EIO!"
    End If
    
    ''# Registry - DONE !
    If regE.Value = True Then
        ''Enable registry
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "0", REG_DWORD
        SetKeyValue HKEY_USERS, ".Default\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "0", REG_DWORD
        MsgBox "Enabled REGEDIT !", vbInformation, "EIO!"
    ElseIf regD.Value = True Then
        ''Disable registry
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "1", REG_DWORD
        SetKeyValue HKEY_USERS, ".Default\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "1", REG_DWORD
        MsgBox "Disabled REGEDIT !", vbInformation, "EIO!"
    End If
     
    ''# USB Port - DONE !
    If usbE.Value = True Then
        ''Enable USB
        C = Shell("REG add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR /v Start /t REG_DWORD /d 3 /f", vbNormalFocus)
        MsgBox "Enabled USB PORT !", vbInformation, "EIO!"
    ElseIf usbD.Value = True Then
        ''Disable USB
        C = Shell("REG add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR /v Start /t REG_DWORD /d 4 /f", vbNormalFocus)
        MsgBox "Disabled USB PORT !", vbInformation, "EIO!"
    End If
    
End Sub
