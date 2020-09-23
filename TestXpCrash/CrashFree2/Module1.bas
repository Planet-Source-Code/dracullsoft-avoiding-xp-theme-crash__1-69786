Attribute VB_Name = "Module1"
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Sub Main()
  '__ Must be done in Main() in a bas module or in the Form_Initialize()
  InitCommonControls
  Dim f As New Form1
  f.Show vbModal
  
End Sub
