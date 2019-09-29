VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchSup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Supplier"
   ClientHeight    =   5550
   ClientLeft      =   50
   ClientTop       =   410
   ClientWidth     =   4810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4810
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtSupName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   4095
         _ExtentX        =   7214
         _ExtentY        =   5733
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Supplier ID"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Supplier Name"
            Object.Width           =   3776
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "* Double Click to select the Supplier Details"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   4920
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmSearchSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim strsql As String

Private Sub Display(sql As String)
    rs.Open sql, cn
    Dim lst As ListItem
    ListView1.ListItems.Clear
    Do While Not rs.EOF
        Set lst = ListView1.ListItems.Add(, , rs(0))
        lst.SubItems(1) = rs(1) & " " & rs(2)
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub cmdSearch_Click()
    strsql = "Select Sup_Id, Sup_FirstName,Sup_LastName from tblSupplier where Sup_FirstName like '" & txtSupName.Text & "%'"
    Set rs = New Recordset
    Display (strsql)
End Sub

Private Sub Form_Load()
    Call cmdSearch_Click
End Sub

Private Sub ListView1_DblClick()
    frmPurchase.txtSupId.Text = ListView1.SelectedItem.Text
    Dim fn As String
    fn = ListView1.SelectedItem.SubItems(1)
    frmPurchase.txtSupName.Text = fn
    Unload Me
    frmPurchase.Show
End Sub
