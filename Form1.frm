VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   5535
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuOhelp 
         Caption         =   "&On-Line Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End
End Sub

Private Sub mnuOhelp_Click()
  ' Declare the Browser
  Dim frmB As New frmBrowser
  ' Starting address
  ' Change this to your help page.
  ' You can copy or link to my stylesheet
  ' to make your help page(s) look and behave
  ' like a help file
  ' http://fibdev.com/styles/help.css
  ' to link to it just place this link in
  ' the head of your html documents
  ' <link rel="text/stylesheet" href="http://fibdev.com/styles/help.css">
  frmB.StartingAddress = "http://fibdev.com/help/cgias"
  frmB.brwWebBrowser.Left = 5
  frmB.brwWebBrowser.Width = frmBrowser.Width - 10
  frmB.Show
End Sub
