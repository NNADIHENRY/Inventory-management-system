VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6210
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   5400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8400
      Top             =   5400
   End
   Begin VB.Image Image1 
      Height          =   6300
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   9540
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmMDI.Show
End Sub

Private Sub Image1_Click()
    Unload Me
    frmMDI.Show
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value >= 100 Then
        Unload Me
        frmMDI.Show
    End If
End Sub
