VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
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
   Begin Crashfree2XP.UserControl1 UserControl11 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1500
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   873
   End
   Begin VB.CheckBox Check1 
      Caption         =   "crash"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Image Image1 
      Height          =   2115
      Left            =   180
      Picture         =   "Form1.frx":000C
      Top             =   2040
      Width           =   7905
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Crash"
      Height          =   1335
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



Private Sub Command1_Click()
  MsgBox "Example by DracullSoft"
End Sub



Private Sub Form_Load()
 Check1.Value = 1
 Check1.Caption = "Usercontrol will NOT cause crash!"
 
 Label1 = "Crash happens at offset 0x773e65e1 when closing an app using xp theme." & _
  vbCrLf & "The Crash only happens in the Compiled exe when a usercontrol is on the form." & _
  vbCrLf & "The method for enabling xp theme consists of a single call to 'InitCommonControls' but this time in Sub Main() and with a .manifest file" & _
  vbCrLf & "Method 1 to avoid the Crash is to add a common dialogs control to the form - As we did in CrashFree1." & _
  vbCrLf & "Method 2 to avoid the Crash is to add an ImageList from Common Controls 5 sp2. ( or another control)"
 
End Sub


