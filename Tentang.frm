VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form2"
   ScaleHeight     =   1665
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H006A5F57&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   0
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   1785
      Left            =   0
      Top             =   0
      Width           =   3810
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function DrawText Lib "USER32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByValcrKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Const LWA_ALPHA = 2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'*****************************************************
Const ScrollText As String = " * * * * * * * * * * * * * * * *" & vbCrLf & _
                             "Enigmal Tool Ver 1.0" & vbCrLf & _
                             vbCrLf & _
                             " By :" & vbCrLf & _
                             " Robbi Nespu" & vbCrLf & _
                             "                      " & vbCrLf & _
                               " Greetinz to:" & _
                             vbCrLf & "All my teachers && lecturers !" & _
                             vbCrLf & "My family, friends.. " & _
                             vbCrLf & "Mambang team, SKDM, EIO group" & _
                             vbCrLf & "and ...you too :)" & vbCrLf & _
                             "* * *  ReLeAse dAtE 08/07/2012 * * * *" & vbCrLf & _
                             " " & vbCrLf & _
                             " * * * * * * * * * * * * * * * *" & vbCrLf & _
                              " Remember that silence is sometimes the best answer " & vbCrLf & _
                              " * * * * * * * * * * * * * * * *" & vbCrLf & _
                            ""

'*****************************************************

Dim EndingFlag As Boolean

Private Sub Form_Activate()
RunMain
End Sub

Private Sub Form_Load()
SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes hwnd, 0, 150, LWA_ALPHA
Picture1.ForeColor = &HFFFFFF
Picture1.FontSize = 11
Picture1.Font = "Lucida Sans Unicode"
End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 30
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long
Dim RectHeight As Long
Me.Refresh
rt = DrawText(Picture1.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)
If rt = 0 Then
    MsgBox "Error : bug! xleh scroll u_u' ", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = Picture1.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = Picture1.ScaleWidth

    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + Picture1.ScaleHeight
End If

Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                   
        Picture1.Cls
        
        DrawText Picture1.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
    
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        

        If DrawingRect.Top < -(RectHeight) Then
            DrawingRect.Top = Picture1.ScaleHeight
            DrawingRect.Bottom = RectHeight + Picture1.ScaleHeight
        End If
        
        Picture1.Refresh
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop

Unload Me
Set Form2 = Nothing
End Sub

Private Sub Form_click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Label1_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Picture1_Click()
Form1.Enabled = True
Unload Me
End Sub


