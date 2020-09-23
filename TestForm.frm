VERSION 5.00
Object = "*\AListBoxOcx.vbp"
Begin VB.Form Form1 
   Caption         =   "Example"
   ClientHeight    =   1815
   ClientLeft      =   2325
   ClientTop       =   3045
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   Begin VB.CommandButton Command2 
      Caption         =   "Down"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Up"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin Project2.CustomListBox CustomListBox1 
      Height          =   1575
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      _extentx        =   4471
      _extenty        =   2355
      backcolor       =   16777215
      fontinfo        =   "TestForm.frx":0000
      forecolor       =   16777215
      graphical       =   -1  'True
      picture         =   "TestForm.frx":0028
      scrollbarbackcolor=   12632256
      scrollbarbordercolor=   8421504
      selboxcolor     =   12632064
      sorted          =   0   'False
   End
   Begin VB.CommandButton remove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Text            =   "Some Text"
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2460
      Picture         =   "TestForm.frx":311A
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1860
      Picture         =   "TestForm.frx":39E4
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1320
      Picture         =   "TestForm.frx":42AE
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   780
      Picture         =   "TestForm.frx":4B78
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "TestForm.frx":5442
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is just the test form.

Option Explicit

Private Sub Command1_Click()
    CustomListBox1.MoveUp
End Sub

Private Sub add_Click()
    Static PicIndex As Integer
    If Text1.Text = "" Then
        MsgBox "Please enter some text.", vbInformation, "Alert"
        Exit Sub
    End If
    
    PicIndex = PicIndex + 1
    If PicIndex > 4 Then
        PicIndex = 0
    End If
    
    CustomListBox1.AddItem Text1.Text, PicIndex
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
    CustomListBox1.Movedown
End Sub

Private Sub Command3_Click()
    CustomListBox1.Clear
End Sub

Private Sub CustomListBox1_DblClick()
    MsgBox "You clicked: " & CustomListBox1.List(CustomListBox1.ListIndex)
End Sub

Private Sub Form_Load()
    ' Add image to use in list
    CustomListBox1.Addimage Image1.Picture ' this will end up being index 0
    CustomListBox1.Addimage Image2.Picture ' index 1
    CustomListBox1.Addimage Image3.Picture ' index 2
    CustomListBox1.Addimage Image4.Picture ' index 3
    CustomListBox1.Addimage Image5.Picture ' index 4
End Sub

Private Sub remove_Click()
    If CustomListBox1.ListIndex = -1 Then Exit Sub
    CustomListBox1.RemoveItem CustomListBox1.ListIndex
End Sub
