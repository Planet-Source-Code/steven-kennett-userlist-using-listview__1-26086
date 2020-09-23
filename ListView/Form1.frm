VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click to hide/show columns"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   4095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Status1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0354
            Key             =   "Status2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A8
            Key             =   "Status3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09FC
            Key             =   "Status4"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "This is the finnished userlist. It is what the user will see."
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ListView example simulating a multi level userlist
'
Dim ListViewColumn As Boolean

Private Sub Form_Load()

' Set the columns
With ListView1.ColumnHeaders
.Add , , "Status", 1000
.Add , , "Username", 2000
.Add , , "Sort", 2000
End With

' Set column view
ListViewColumn = False

' Set the sort method
ListView1.Sorted = True
ListView1.SortOrder = lvwAscending
ListView1.SortKey = 2

' Add the items
' Note: @ indicates top level user
ListViewAdd "Status2", "SKABB", "SKABB"
ListViewAdd "Status2", "ColonelKill", "ColonelKill"
ListViewAdd "Status1", "Tapewormz", "@Tapewormz"
ListViewAdd "Status2", "Grillo", "Grillo"

End Sub

Private Function ListViewAdd(UList1 As String, UList2 As String, UList3 As String)

Dim ListViewAddEntry As ListItem

Set ListViewAddEntry = ListView1.ListItems.Add(, , , , UList1)
ListViewAddEntry.SubItems(1) = UList2
ListViewAddEntry.SubItems(2) = UList3

End Function

Private Sub Command1_Click()

' Hide/Show columns on demand
If ListViewColumn = False Then
ListViewColumn = True
ListView1.HideColumnHeaders = ListViewColumn
ListView1.ColumnHeaders.Item(1).Width = 256
ListView1.ColumnHeaders.Item(3).Width = 0
ListView1.Width = 2500
ListView1.BackColor = &H80000008
ListView1.ForeColor = &H80000005
ListView1.Appearance = ccFlat
ListView1.BorderStyle = ccNone

Else
ListViewColumn = False
ListView1.HideColumnHeaders = ListViewColumn
ListView1.ColumnHeaders.Item(1).Width = 1000
ListView1.ColumnHeaders.Item(3).Width = 2000
ListView1.Width = 4095
ListView1.BackColor = &H80000005
ListView1.ForeColor = &H80000008
ListView1.Appearance = cc3D
ListView1.BorderStyle = ccFixedSingle
End If

End Sub
