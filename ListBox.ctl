VERSION 5.00
Begin VB.UserControl CustomListBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   Begin VB.PictureBox ImageSizer 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1920
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox TheList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   60
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   1
      Top             =   60
      Width           =   2820
   End
   Begin VB.PictureBox ScrollBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3285
      Picture         =   "ListBox.ctx":0000
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   15
      Width           =   150
   End
   Begin VB.Shape VScrollBar 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1575
      Left            =   3240
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "CustomListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Custom Listbox Control 1.1
' By Patrick Gillespie (patorjk@aol.com)
' 8.16.00
' http://www.patorjk.com/

' This is an example on how to create your own listbox. This listbox has most of the
' features of a normal listbox, except you can also use picture backgrounds in it.
' This example is still being improved on, so if you find any errors or have any
' suggestions please email me.

Option Explicit

Dim ListItems() As String
Dim ListCount As Integer
Dim SortList As Boolean
Dim SelItem As Integer
Dim OldSelItem As Integer
Dim ListItemHeight As Long

Dim SelColor As Long

Dim CanScroll As Boolean

' for scrollbar
Dim OffSetY As Integer
Dim TopY As Integer

' picture info
Dim ThePic As Picture, IsGraphical As Boolean
Dim PicWidth As Integer, PicHeight As Integer
Dim PicArray() As StdPicture, PicCount As Integer
Dim PicIndex() As Integer

Private WithEvents TheFont As StdFont
Attribute TheFont.VB_VarHelpID = -1

' The events
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Get Graphical() As OLE_OPTEXCLUSIVE
    ' This is called when we want to know the state of the Graphical property
    Graphical = IsGraphical
End Property

Public Property Let Graphical(ByVal TheOpinion As OLE_OPTEXCLUSIVE)
    ' This sets the Graphical property
    IsGraphical = TheOpinion
    Call DrawListBox
    PropertyChanged "Graphical"
End Property

Public Property Get Picture() As Picture
    ' This is when we want to know the picture being stored
    Set Picture = ThePic
End Property

Public Property Set Picture(ByVal LaPic As Picture)
    ' This sets the Picture property
    Set ThePic = LaPic
    Call DrawListBox
    PropertyChanged "Picture"
End Property

Public Property Get FontInfo() As StdFont
    ' Get the font information
    Set FontInfo = TheFont
End Property

Public Property Set FontInfo(NewFont As StdFont)
    ' Set the new font information and then redraw
    Set TheFont = NewFont
    Set TheList.Font = NewFont
    Call DrawListBox
    
    PropertyChanged "FontInfo"
End Property

Public Property Get SelBoxColor() As OLE_COLOR
    ' Gets current color
    SelBoxColor = SelColor
End Property

Public Property Let SelBoxColor(ByVal NewColor As OLE_COLOR)
    ' Sets color
    SelColor = NewColor
    ' redraw list
    DrawListBox
    
    PropertyChanged "SelBoxColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ' Gets current color
    ForeColor = TheList.ForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    ' Sets color
    TheList.ForeColor = NewColor
    ' redraw list with color
    DrawListBox
    
    PropertyChanged "ForeColor"
End Property

Public Property Get ScrollBarBorderColor() As OLE_COLOR
    ' Gets current color
    ScrollBarBorderColor = VScrollBar.BorderColor
End Property

Public Property Let ScrollBarBorderColor(ByVal NewColor As OLE_COLOR)
    ' Sets color
    VScrollBar.BorderColor = NewColor
    
    PropertyChanged "ScrollBarBorderColor"
End Property

Public Property Get ScrollBarBackColor() As OLE_COLOR
    ' Gets current color
    ScrollBarBackColor = VScrollBar.BackColor
End Property

Public Property Let ScrollBarBackColor(ByVal NewColor As OLE_COLOR)
    ' Sets color
    VScrollBar.BackColor = NewColor
    PropertyChanged "ScrollBarBackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    ' Gets current color
    BackColor = TheList.BackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    ' Sets color
    TheList.BackColor = NewColor
    PropertyChanged "BackColor"
End Property

Public Property Get Sorted() As OLE_OPTEXCLUSIVE
    ' This is called when we want to know the state of the Graphical property
    Sorted = SortList
End Property

Public Property Let Sorted(ByVal TheOpinion As OLE_OPTEXCLUSIVE)
    ' This sets the Graphical property
    SortList = TheOpinion
    PropertyChanged "Sorted"
End Property

Private Sub ScrollBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Scrolllistbasedonkey(KeyCode)
End Sub

Private Sub ScrollBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Set up variables in case the list needs to be scrolled (see mouse move event)
    Dim CurPos As POINTAPI
    Call GetCursorPos(CurPos)
    OffSetY = CurPos.y
    TopY = ScrollBox.Top
End Sub

Private Sub ScrollBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Move the list depending on the location of the scrollbox
    Dim YPos As Integer, CurPos As POINTAPI
    If Button = 1 And CanScroll = True Then
        Call GetCursorPos(CurPos)
        YPos = TopY - (OffSetY - CurPos.y)
        If YPos < 1 Then
            ScrollBox.Top = 1
            Call ScrollList(ScrollBox.Top)
        ElseIf YPos > (VScrollBar.Height - ScrollBox.ScaleHeight - 1) Then
            ScrollBox.Top = VScrollBar.Height - ScrollBox.Height - 1
            Call ScrollList(ScrollBox.Top)
        Else
            ScrollBox.Top = YPos
            Call ScrollList(ScrollBox.Top)
        End If
    ElseIf CanScroll = False Then
        ScrollBox.Top = 1
    End If
End Sub

Private Sub TheList_Click()
    ' Raise the Click event
    RaiseEvent Click
End Sub

Private Sub TheList_DblClick()
    ' Raise the DblClick event
    RaiseEvent DblClick
End Sub

Private Sub TheList_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Scrolllistbasedonkey(KeyCode)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TheList_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TheList_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TheList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim YPos As Integer
    Call SetUpList
    
    YPos = CInt(y)
    SelItem = Int(YPos / ListItemHeight)
    Call DrawSelItem
    
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub TheList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' This code is excuted if the user is holding the left mouse button down and
    ' moving the mouse.
    Dim CurPos As POINTAPI, ListWndSize As RECT
    If Button = 1 Then
        Dim YPos As Integer
        Call SetUpList
        Call GetWindowRect(UserControl.hwnd, ListWndSize)
        Call GetCursorPos(CurPos)
        
        If CurPos.y < ListWndSize.Top Then
            ' The cursor is above the listbox
            Call Timeout(0.1)
            'Call MoveUp
        ElseIf CurPos.y > ListWndSize.Bottom Then
            ' The cursor is below the listbox
            Call Timeout(0.1)
            'Call MoveDown
        Else
            ' The cursor is on the listbox so just select a new item if
            ' one is moved over.
            YPos = CInt(y)
            SelItem = Int(YPos / ListItemHeight)
        End If
        
        Call DrawSelItem
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub DrawSelItem()
    ' This sub draws the selected item box around the selected item
    Dim Y1 As Long, Y2 As Long
    If SelItem <= ListCount And SelItem >= 0 And SelItem <> OldSelItem Then
        ' Draw Selected Box
        Y1 = CLng(SelItem * ListItemHeight)
        Y2 = CLng((SelItem + 1) * ListItemHeight)
        Call DrawRectangle(TheList.hdc, 0, Y1, TheList.ScaleWidth, Y2, SelColor)
        Call UpdateItem(SelItem)
        ' Clear Away Old Select Box
        If OldSelItem <> -1 Then
            Y1 = CLng(OldSelItem * ListItemHeight)
            Y2 = CLng((OldSelItem + 1) * ListItemHeight)
            If IsGraphical = True Then
                ' if a graphic is being used then clear away the old select box with
                ' the graphic image that goes in it's place
                Call DrawImageRect(0, Y1, TheList.ScaleWidth, Y2)
            Else
                Call DrawRectangle(TheList.hdc, 0, Y1, TheList.ScaleWidth, Y2, TheList.BackColor)
            End If
            Call UpdateItem(OldSelItem)
        End If
        OldSelItem = SelItem
    End If
End Sub

Private Sub DrawImageRect(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    ' This sub draws what the image under a selected item would look like
    Dim PicsDownY1 As Integer, PicsDownY2 As Integer, Offset As Long
    Dim StartCut As Long, StopCut As Long, i As Integer, i2 As Integer
    PicsDownY1 = Int(Y1 / PicHeight)
    PicsDownY2 = Int(Y2 / PicHeight)

    ' Set up canvas to draw on
    ImageSizer.Cls
    ImageSizer.AutoRedraw = True
    ImageSizer.Width = TheList.ScaleWidth
    ImageSizer.Height = Y2 - Y1

    StartCut = Y1 - (PicsDownY1 * PicHeight)
    If PicsDownY2 = PicsDownY1 Then
        StopCut = Y2 - (PicsDownY2 * PicHeight)
        ' Draw Top - No middle or bottom area needed
        For i = 0 To X2 Step PicWidth
            ImageSizer.PaintPicture ThePic, i, 0, (i + 1) * PicWidth, StopCut - StartCut, 0, StartCut, (i + 1) * PicWidth, StopCut - StartCut
        Next
        ImageSizer.Picture = ImageSizer.Image
        ImageSizer.AutoRedraw = False
    Else
        ' DrawTop
        StopCut = PicHeight
        For i = 0 To X2 Step PicWidth
            ImageSizer.PaintPicture ThePic, i, 0, (i + 1) * PicWidth, StopCut - StartCut, 0, StartCut, (i + 1) * PicWidth, StopCut - StartCut
        Next
        ' Draw Middle
        Offset = StopCut - StartCut
        For i = 0 To (PicsDownY2 - PicsDownY1) - 2
            For i2 = 0 To X2 Step PicWidth
                ImageSizer.PaintPicture ThePic, i2, Offset + (i * PicHeight), PicWidth, PicHeight, 0, 0, PicWidth, PicHeight
            Next
        Next
        ' Draw Bottom
        StartCut = 0
        StopCut = Y2 - (PicsDownY2 * PicHeight)
        For i = 0 To X2 Step PicWidth
            ImageSizer.PaintPicture ThePic, i, Offset + ((PicsDownY2 - PicsDownY1) - 1) * PicHeight, i + PicWidth, StopCut - StartCut, 0, StartCut, i + PicWidth, StopCut - StartCut
        Next
        ImageSizer.Picture = ImageSizer.Image
        ImageSizer.AutoRedraw = False
    End If
    TheList.PaintPicture ImageSizer.Picture, 0, Y1
    ImageSizer.Picture = LoadPicture("")
End Sub

Private Sub TheList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
    Set TheFont = New StdFont
    TheList.Top = 0
    TheFont.Name = "Verdana"
    TheFont.Size = 7
    ListCount = -1
    PicCount = -1
    SelItem = -1
    OldSelItem = -1
    SetUpList
    UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
    ' This is called only when the control is first created.
    ' It's not in the initialize event because at that point in time the control
    ' has not yet been placed on the form.
    Set FontInfo = TheFont
    SelColor = &HC0C000
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Scrolllistbasedonkey(KeyCode)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim XPos As Integer, YPos As Integer
    XPos = CInt(x)
    YPos = CInt(y)
    If XPos > TheList.ScaleWidth Then
        If CanScroll = True Then
            If YPos < 1 Then
                ScrollBox.Top = 1
                Call ScrollList(ScrollBox.Top)
            ElseIf YPos > (VScrollBar.Height - ScrollBox.ScaleHeight - 1) Then
                ScrollBox.Top = VScrollBar.Height - ScrollBox.Height - 1
                Call ScrollList(ScrollBox.Top)
            Else
                ScrollBox.Top = YPos
                Call ScrollList(ScrollBox.Top)
            End If
        ElseIf CanScroll = False Then
            ScrollBox.Top = 1
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' This event is called every time the control is created except for the
    ' first time you put it on the form.
    TheList.BackColor = PropBag.ReadProperty("BackColor", RGB(255, 255, 255))
    TheList.ForeColor = PropBag.ReadProperty("ForeColor", 0)
    VScrollBar.BackColor = PropBag.ReadProperty("ScrollBarBackColor", &HC0C0C0)
    VScrollBar.BorderColor = PropBag.ReadProperty("ScrollBarBorderColor", &H808080)
    SelBoxColor = PropBag.ReadProperty("SelBoxColor", &HC0C000)
    SortList = PropBag.ReadProperty("Sorted", False)
    IsGraphical = PropBag.ReadProperty("Graphical", True)
    Set ThePic = PropBag.ReadProperty("Picture", LoadPicture(""))
    Set FontInfo = PropBag.ReadProperty("FontInfo", UserControl.Font)
    
    If ThePic = UserControl.Picture Then
        IsGraphical = False
    End If
End Sub

Private Sub UserControl_Resize()
    ' Set things up right, and redraw the list if needed
    Static OldWidth As Long, OldHeight As Long
    
    UserControl.ScaleMode = 3
    If UserControl.ScaleWidth > 50 Then
        TheList.Width = UserControl.ScaleWidth - VScrollBar.Width - 1
        TheList.Left = 0
        VScrollBar.Left = TheList.Width + 1
        ScrollBox.Left = VScrollBar.Left + 1
    End If
    
    If UserControl.ScaleHeight > 50 Then
        VScrollBar.Height = UserControl.ScaleHeight
    End If

    Call SetUpList
    
    If UserControl.ScaleWidth > OldWidth And OldWidth <> 0 Then
        Call DrawListBox
    Else
        If UserControl.ScaleHeight > OldHeight And OldHeight <> 0 Then
            Call DrawListBox
        End If
    End If
    
    OldWidth = UserControl.ScaleWidth
    OldHeight = UserControl.ScaleHeight
End Sub

Public Sub AddItem(NewItem As String, Optional PictureIndex As Integer = -1)
    ' Adds an item to the list array and then redraws the list
    ListCount = ListCount + 1
    ReDim Preserve ListItems(ListCount) As String
    ListItems(ListCount) = NewItem
    ReDim Preserve PicIndex(ListCount) As Integer
    If PictureIndex > PicCount Then
        ' they put in an invalid pictureindex
        PicIndex(ListCount) = -1
    Else
        PicIndex(ListCount) = PictureIndex
    End If
    
    If SortList = True Then
        Call SortTheList
    Else
        Call DrawListBox
    End If
    Call ScrollScrollBox(TheList.Top * -1)
End Sub

Public Sub AddImage(LaPic As Picture)
    PicCount = PicCount + 1
    ReDim Preserve PicArray(PicCount) As StdPicture
    Set PicArray(PicCount) = LaPic
End Sub

Public Sub SortTheList()
    ' This sub sorts the list a to z
    Dim i As Integer, i2 As Integer, Hold As String
    For i = 0 To ListCount
        For i2 = 0 To ListCount
            If i <> i2 Then
                If LCase$(ListItems(i)) < LCase$(ListItems(i2)) Then
                    Hold = ListItems(i)
                    ListItems(i) = ListItems(i2)
                    ListItems(i2) = Hold
                End If
            End If
        Next
    Next
    Call DrawListBox
End Sub

Private Sub DrawListBox()
    ' Draw the list box
    Dim i As Integer, SidePicHeight As Integer, SidePicWidth As Integer
    
    Call SetUpList
    
    ' clear and draw items on list
    TheList.Cls
    
    ' draw image background if set to
    If IsGraphical = True Then
        Call DrawImageBG
    End If
    
    ' Note: List items are drawn in two places, this sub, and in the updateitem sub
    For i = 0 To ListCount
        If PicIndex(i) = -1 Then
            ' Item doesn't have a picture next to it
            TheList.CurrentX = 3
            TheList.CurrentY = (ListItemHeight * (i))
            TheList.Print ListItems(i)
        Else
            ' Item does have a picture next to it
            TheList.CurrentX = 3
            TheList.CurrentY = (ListItemHeight * (i))
            SidePicHeight = ListItemHeight - 2
            SidePicWidth = ListItemHeight - 2
            TheList.PaintPicture PicArray(PicIndex(i)), TheList.CurrentX, TheList.CurrentY + 1, SidePicWidth, SidePicHeight, 0, 0
            TheList.CurrentX = 3 + SidePicWidth + 3
            TheList.CurrentY = (ListItemHeight * (i))
            TheList.Print ListItems(i)
        End If
    Next
    
    OldSelItem = -1
    DrawSelItem
End Sub

Private Sub DrawImageBG()
    ' This sub draws the background
    Dim x As Integer, y As Integer
    On Error Resume Next
    Call GetImageSize
    For x = 0 To TheList.ScaleWidth Step PicWidth
        For y = 0 To TheList.ScaleHeight Step PicHeight
            TheList.PaintPicture ThePic, x, y
        Next
    Next
End Sub

Private Sub GetImageSize()
    ' This sub gets the size of the image to use for the back of the listbox
    Set ImageSizer.Picture = ThePic
    PicWidth = ImageSizer.ScaleWidth
    PicHeight = ImageSizer.ScaleHeight
    ImageSizer.Picture = LoadPicture("")
End Sub

Private Sub SetUpList()
    ' set the listbox up
    ListItemHeight = CInt(TheList.TextHeight("M") + 2)
    ' set list height
    If (ListItemHeight * (ListCount + 1)) > UserControl.ScaleHeight - 1 Then
        TheList.Height = ListItemHeight * (ListCount + 1) + 1
        CanScroll = True
    Else
        TheList.Height = UserControl.Height
        CanScroll = False
    End If
End Sub

Private Sub UpdateItem(index As Integer, Optional ClearItem As Boolean = False)
    Dim SidePicHeight As Integer, SidePicWidth As Integer
    Call SetUpList
    If ClearItem = True Then
        TheList.Line (0, OldSelItem * ListItemHeight)-(TheList.ScaleWidth, (OldSelItem + 1) * ListItemHeight), TheList.BackColor, BF
    End If
    If PicIndex(index) = -1 Then
        ' Item doesn't have a picture next to it
        TheList.CurrentX = 3
        TheList.CurrentY = (ListItemHeight * (index))
        TheList.Print ListItems(index)
    Else
        ' Item does have a picture next to it
        TheList.CurrentX = 3
        TheList.CurrentY = (ListItemHeight * (index))
        SidePicHeight = ListItemHeight - 2
        SidePicWidth = ListItemHeight - 2
        TheList.PaintPicture PicArray(PicIndex(index)), TheList.CurrentX, TheList.CurrentY + 1, SidePicWidth, SidePicHeight, 0, 0
        TheList.CurrentX = 3 + SidePicWidth + 3
        TheList.CurrentY = (ListItemHeight * (index))
        TheList.Print ListItems(index)
    End If
End Sub

Private Sub ScrollList(Pos As Long)
    ' This sub scrolls the list
    Dim SpaceToScroll As Long, TheStep As Double, BarLength As Integer
    Pos = Pos - 1
    BarLength = VScrollBar.Height - ScrollBox.ScaleHeight
    SpaceToScroll = (ListItemHeight * (ListCount + 1)) - (UserControl.ScaleHeight - 1)
    TheStep = SpaceToScroll / BarLength
    TheList.Top = 0 - (Pos * TheStep)
End Sub

Private Sub ScrollScrollBox(ListPos As Long)
    ' This sub moves the scrollbox depending on what position you want to
    ' move the list to.
    Dim SpaceToScroll As Long, TheStep As Double, BarLength As Integer
    Dim ScrollBoxTop As Integer
    BarLength = VScrollBar.Height - ScrollBox.ScaleHeight - 1
    SpaceToScroll = (ListItemHeight * (ListCount + 1)) - (UserControl.ScaleHeight - 1)
    TheStep = SpaceToScroll / BarLength
    ScrollBoxTop = (ListPos / TheStep) + 1

    If SelItem = ListCount And CanScroll = True Then
        ScrollBox.Top = VScrollBar.Height - ScrollBox.ScaleHeight - 1
    Else
        ScrollBox.Top = ScrollBoxTop
    End If
    
    Call ScrollList(ScrollBox.Top)
    ' TheList.Top = 0 - ListPos
End Sub

Private Sub Scrolllistbasedonkey(KeyCode As Integer)
    ' This sub is called when the user presses a key on the listbox
    If KeyCode = vbKeyUp Then
        MoveUp
    ElseIf KeyCode = vbKeyDown Then
        MoveDown
    End If
End Sub

Public Function LCount() As Integer
    ' Returns the number of items in the list
    LCount = ListCount + 1
End Function

Public Function Listindex() As Integer
    ' Returns the selected item
    Listindex = SelItem
End Function

Public Function List(index As Integer) As String
    ' Returns an item in the list array depending on the index you put in
    List = ListItems(index)
End Function

Public Sub RemoveItem(index As Integer)
    ' This sub removes an item from a listbox
    Dim i As Integer
    ' Exit sub if item is not on the list
    If index = -1 Or index > ListCount Then Exit Sub
    ' Set selitem status
    If index = SelItem Then
        SelItem = -1
        OldSelItem = -1
    ElseIf SelItem > index Then
        SelItem = SelItem - 1
        OldSelItem = OldSelItem - 1
    End If
    
    ListCount = ListCount - 1
    For i = index To ListCount
        ListItems(i) = ListItems(i + 1)
        PicIndex(i) = PicIndex(i + 1)
    Next
    If ListCount <> -1 Then
        ReDim Preserve ListItems(ListCount) As String
        ReDim Preserve PicIndex(ListCount) As Integer
    End If
    Call SetUpList
    If CanScroll = True Then
        Call ScrollList(ScrollBox.Top)
    Else
        ScrollBox.Top = 1
        TheList.Top = 0
    End If
    Call DrawListBox
End Sub

Public Sub Clear()
    ' Clears all the items out of the list box
    ListCount = -1
    ReDim ListItems(0) As String
    ReDim PicIndex(0) As Integer
    Call SetUpList
    If CanScroll = True Then
        Call ScrollList(ScrollBox.Top)
    Else
        ScrollBox.Top = 1
        TheList.Top = 0
    End If
    Call DrawListBox
End Sub

Public Sub MoveUp()
    ' Moves the selected item up one
    Dim ItemTop As Long
    If SelItem = -1 Or SelItem = 0 Then Exit Sub
    ' Move the selected item index up one
    SelItem = SelItem - 1

    ItemTop = (SelItem) * ListItemHeight
    If ItemTop < Abs(TheList.Top) Then
        Call ScrollScrollBox(ItemTop)
    End If
    
    Call DrawSelItem
End Sub

Public Sub MoveDown()
    ' Moves the selected item down one
    Dim ItemBottom As Long
    If SelItem = -1 Or SelItem = ListCount Then Exit Sub
    ' Move the selected item index up one
    SelItem = SelItem + 1
    
    ItemBottom = (SelItem + 1) * ListItemHeight
    If ItemBottom > Abs(TheList.Top) + UserControl.ScaleHeight Then
        Call ScrollScrollBox(ItemBottom - UserControl.ScaleHeight)
    End If
    
    Call DrawSelItem
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' This event is called when the control needs to save the values of the
    ' properties (like before you go into run mode).
    PropBag.WriteProperty "BackColor", TheList.BackColor
    PropBag.WriteProperty "FontInfo", FontInfo
    PropBag.WriteProperty "ForeColor", TheList.ForeColor
    PropBag.WriteProperty "Graphical", IsGraphical
    PropBag.WriteProperty "Picture", ThePic
    PropBag.WriteProperty "ScrollBarBackColor", VScrollBar.BackColor
    PropBag.WriteProperty "ScrollBarBorderColor", VScrollBar.BorderColor
    PropBag.WriteProperty "SelBoxColor", SelColor
    PropBag.WriteProperty "Sorted", SortList
End Sub
