VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4740
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crashfree1XP.UserControl1 UserControl11 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1620
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   873
   End
   Begin VB.CheckBox Check1 
      Caption         =   "crash"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1620
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   435
      Left            =   6840
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2115
      Left            =   180
      Picture         =   "Form1.frx":000C
      Top             =   2280
      Width           =   7905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Crash"
      Height          =   1455
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   8115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 12/20/2007
' Author    : Dracull
' Purpose   : Demonstration of XP theme crash in compiled exe and how to avoid it
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private Sub Command1_Click()
  MsgBox "Example by DracullSoft"
End Sub

Private Sub Form_Initialize()
  '__ Must be done in Main() in a bas module or in the Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
 Check1.Value = 1
 Check1.Caption = "Usercontrol will NOT cause crash!"
 
 Label1 = "Crash happens at offset 0x773e65e1 when closing an app using xp theme." & _
  vbCrLf & "The Crash only happens in the Compiled exe when a usercontrol is on the form." & _
  vbCrLf & "The method for enabling xp theme consists of a single call to 'InitCommonControls' in the Form_Initialize() + a Res file." & _
  vbCrLf & "Method 1 to avoid the Crash is to add a common dialogs control to the form. See CrashFree1." & _
  vbCrLf & "Method 2 to avoid the Crash is to call 'InitCommonControls' from a Sub Main().  See CrashFree2."
 
End Sub


