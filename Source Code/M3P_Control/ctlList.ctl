VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   KeyPreview      =   -1  'True
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ToolboxBitmap   =   "ctlList.ctx":0000
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3600
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Text"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SortString"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "ctlList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Const DT_RIGHT = &H2
Const DT_LEFT = &H0
Const DT_CENTER = &H1
Const DT_NOPREFIX = &H800
Const colTime As Long = 50

Dim colText As Long
Enum PicAlignment
    picTopLeft = 0
    picTopRight = 1
    picBottomLeft = 2
    picBottomRight = 3
    picCenter = 4
    picTitle = 5
    picStrech = 6
End Enum

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event Resize()

Dim currentPlay As Long
Dim CurrentIndex As Long

Type ListStyle
    Picture As String
    BackColor As Long
    ForeColor As Long
    PlayColor As Long
    PlayForeColor As Long
    SelectColor As Long
    SelectBorderColor As Long
    PlayBold As Boolean
    ShowNumber As Boolean
End Type

Dim ePicAlign As PicAlignment
Dim tListStyle As ListStyle
Dim intStart As Integer
Dim intEnd As Integer

Public Function AddItem(Optional Key As Long, Optional text As String, Optional time As String) As ListItem
    On Error Resume Next
    Dim i As Long
    Dim tmp As Boolean
    
    If Key = vbNull Then Key = lvwList.ListItems.Count + 1
    If text = "" Then text = "Uncorrect text"
    If time = "" Then time = "00:00"
    
    lvwList.Sorted = False
    
    lvwList.ListItems.Add lvwList.ListItems.Count + 1, , text
    lvwList.ListItems(lvwList.ListItems.Count).SubItems(1) = time
    lvwList.ListItems(lvwList.ListItems.Count).SubItems(3) = Key
    If tListStyle.ShowNumber Then
        ShowNumber (lvwList.ListItems.Count)
        For i = 1 To lvwList.ListItems.Count
            Call ReNumber(i)
        Next i
    End If
    Set AddItem = lvwList.ListItems(lvwList.ListItems.Count)
End Function
Public Sub DisplayList()
    Call DrawList
    Dim z As Long
    Dim r As Long
    z = lvwList.GetFirstVisible.index
    For r = z To z + CLng(ItemPerPage)
        If r > lvwList.ListItems.Count Then Exit For
        If lvwList.ListItems(r).Selected = True Then
            Call DrawSelect(r)
        End If
    Next r
    Call Play
End Sub
Public Property Get Number() As Boolean
    Number = tListStyle.ShowNumber
End Property
Public Property Let Number(ByVal NewValue As Boolean)
    tListStyle.ShowNumber = NewValue
    Dim i As Long
    If lvwList.ListItems.Count > 0 Then
        If Not NewValue Then
            For i = 1 To lvwList.ListItems.Count
                lvwList.ListItems(i).text = Mid(lvwList.ListItems(i).text, 6)
            Next i
        Else
            For i = 1 To lvwList.ListItems.Count
                If Mid(lvwList.ListItems(i).text, 1, 3) = Format(i, "000") Then lvwList.ListItems(i).text = Mid(lvwList.ListItems(i).text, 6)
                Call ShowNumber(i)
                Call ReNumber(i)
            Next i
        End If
    End If
    Call DisplayList
End Property
Public Property Get PlayBold() As Boolean
    PlayBold = tListStyle.PlayBold
End Property
Public Property Let PlayBold(ByVal NewValue As Boolean)
    tListStyle.PlayBold = NewValue
End Property
Public Property Get Picture() As String
    Picture = tListStyle.Picture
End Property
Public Property Let Picture(ByVal NewVal As String)
    On Error Resume Next
    If NewVal <> "" Then
        tListStyle.Picture = NewVal
        Set pic.Picture = LoadPicture(tListStyle.Picture)
        Call drawPic
    Else
        Set pic.Picture = Nothing
        Set UserControl.Picture = Nothing
        pic.BackColor = tListStyle.BackColor
    End If
    PropertyChanged "Picture"
End Property
Public Property Get PicAlign() As PicAlignment
    PicAlign = ePicAlign
End Property
Public Property Let PicAlign(NewVal As PicAlignment)
    ePicAlign = NewVal
    Call drawPic
    PropertyChanged "PicAlign"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
    UserControl.MousePointer = vNewValue
    PropertyChanged "MousePointer"
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal Img As Picture)
    Set UserControl.MouseIcon = Img
    PropertyChanged "MouseIcon"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = tListStyle.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    tListStyle.BackColor = NewValue
    UserControl.BackColor = tListStyle.BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = tListStyle.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    tListStyle.ForeColor = NewValue
    PropertyChanged "ForeColor"
End Property
Public Property Get PlayColor() As OLE_COLOR
    PlayColor = tListStyle.PlayColor
End Property
Public Property Let PlayColor(ByVal NewValue As OLE_COLOR)
    tListStyle.PlayColor = NewValue
    PropertyChanged "PlayColor"
End Property
Public Property Get PlayForeColor() As OLE_COLOR
    PlayForeColor = tListStyle.PlayForeColor
End Property
Public Property Let PlayForeColor(ByVal NewValue As OLE_COLOR)
    tListStyle.PlayForeColor = NewValue
    PropertyChanged "PlayForeColor"
End Property
Public Property Get SelectColor() As OLE_COLOR
    SelectColor = tListStyle.SelectColor
End Property
Public Property Let SelectBorderColor(ByVal NewValue As OLE_COLOR)
    tListStyle.SelectBorderColor = NewValue
    PropertyChanged "SelectBorderColor"
End Property
Public Property Get SelectBorderColor() As OLE_COLOR
    SelectColor = tListStyle.SelectBorderColor
End Property
Public Property Let SelectColor(ByVal NewValue As OLE_COLOR)
    tListStyle.SelectColor = NewValue
    PropertyChanged "SelectColor"
End Property
Public Property Set Font(ByVal NewFont As StdFont)
    Set UserControl.Font = NewFont
    Set lvwList.Font = NewFont
    PropertyChanged "Font"
End Property
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Get EnsureVisible(index As Long) As Boolean
    EnsureVisible = lvwList.ListItems(index).EnsureVisible
End Property
Public Sub ItemVisible(index As Long)
    On Error Resume Next
    If index = 0 Or index > lvwList.ListItems.Count Then Exit Sub
    lvwList.ListItems(index).EnsureVisible
    Call DisplayList
End Sub
Public Property Get IsSort() As Boolean
    IsSort = lvwList.Sorted
End Property
Public Property Let IsSort(NewVal As Boolean)
    Dim i As Long
    Dim tmp As Long
    If currentPlay <> 0 Then tmp = lvwList.ListItems(currentPlay).SubItems(3)
    
    
    lvwList.Sorted = NewVal
    lvwList.SortKey = 2
    
    For i = 1 To lvwList.ListItems.Count
        If lvwList.ListItems(i).SubItems(3) = tmp Then
            currentPlay = i
            Exit For
        End If
    Next i
    
    If tListStyle.ShowNumber Then
        For i = 1 To lvwList.ListItems.Count
            Call ReNumber(i)
        Next i
    End If
    Call DisplayList
    PropertyChanged "IsSort"
End Property
Public Property Get IsSortStyle() As Boolean
    IsSortStyle = IIf(lvwList.SortOrder = lvwAscending, True, False)
End Property
Public Property Let IsSortStyle(NewVal As Boolean)
    Dim i As Long
    Dim tmp As Long
    If currentPlay <> 0 Then tmp = lvwList.ListItems(currentPlay).SubItems(3)
    
    Call IIf(NewVal, lvwList.SortOrder = lvwAscending, lvwList.SortOrder = lvwDescending)
   
    For i = 1 To lvwList.ListItems.Count
        If lvwList.ListItems(i).SubItems(3) = tmp Then
            currentPlay = i
            Exit For
        End If
    Next i
    
    If tListStyle.ShowNumber Then
        For i = 1 To lvwList.ListItems.Count
            Call ReNumber(i)
        Next i
    End If
    Call DisplayList
    PropertyChanged "IsSortStyle"
End Property
Public Property Get ListItemCount() As Long
    ListItemCount = lvwList.ListItems.Count
End Property




Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    currentPlay = CurrentIndex
    RaiseEvent DblClick
End Sub
Public Sub AddSortString(index As Long, strSort As String)
    On Error Resume Next
    lvwList.ListItems(index).SubItems(2) = strSort
End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If KeyCode = 40 Or KeyCode = 38 Then
        Select Case Shift
            Case 0 'None
                For i = 1 To lvwList.ListItems.Count
                    lvwList.ListItems(i).Selected = False
                Next i
                intStart = 0
                intEnd = 0
                CurrentIndex = IIf(KeyCode = 40, CurrentIndex + 1, CurrentIndex - 1)
                lvwList.ListItems(CurrentIndex).Selected = True
            Case 1 'shift
                intStart = CurrentIndex
                CurrentIndex = IIf(KeyCode = 40, CurrentIndex + 1, CurrentIndex - 1)
                intEnd = CurrentIndex
                If intStart < intEnd Then
                    For i = intStart To intEnd Step 1
                        lvwList.ListItems(i).Selected = True
                    Next i
                Else
                    For i = intStart To intEnd Step -1
                        lvwList.ListItems(i).Selected = True
                    Next i
                End If
        End Select
        Call DisplayList
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If lvwList.ListItems.Count = 0 Then
        Exit Sub
    Else
        Dim firstVisible As Long
        Dim tmp As Long
        Dim i As Long
        
        If Button = vbLeftButton Or Button = vbRightButton Then
            firstVisible = lvwList.GetFirstVisible.index
            tmp = firstVisible + CInt(Y \ lvwList.ListItems(1).Height)
            If tmp > 0 And tmp <= lvwList.ListItems.Count Then
                Select Case Shift
                    Case 0
                        For i = 1 To lvwList.ListItems.Count
                            lvwList.ListItems(i).Selected = False
                        Next i
                        intStart = 0
                        intEnd = 0
                        CurrentIndex = tmp
                        lvwList.ListItems(CurrentIndex).Selected = True
                    Case 1 'shift
                        intEnd = tmp
                        intStart = CurrentIndex
                        If intStart < intEnd Then
                            For i = intStart To intEnd Step 1
                                lvwList.ListItems(i).Selected = True
                            Next i
                        Else
                            For i = intStart To intEnd Step -1
                                lvwList.ListItems(i).Selected = True
                            Next i
                        End If
                    Case 2
                        lvwList.ListItems(tmp).Selected = Not lvwList.ListItems(tmp).Selected
                        intStart = 0
                        intEnd = 0
                        If lvwList.ListItems(tmp).Selected = True Then
                            CurrentIndex = tmp
                        End If
                End Select
                Call DisplayList
            End If
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Public Function TextHeight() As Long
    On Error Resume Next
    If ItemPerPage <> 0 Then
        TextHeight = lvwList.ListItems(1).Height
    End If

End Function
Private Sub DrawSelect(index As Long)
    On Error Resume Next
    Dim rec As RECT
    
    If lvwList.ListItems.Count = 0 Then Exit Sub
    If index = 0 Or index > lvwList.ListItems.Count Then Exit Sub
    
    UserControl.ForeColor = tListStyle.ForeColor
    UserControl.FontBold = False
    
    rec.top = lvwList.ListItems(index).top
    rec.Left = 1
    rec.Bottom = rec.top + lvwList.ListItems(index).Height
    rec.Right = rec.Left + colText
    
    UserControl.Line (0, rec.top - 1)-(lvwList.Width - 1, rec.Bottom - 1), tListStyle.SelectColor, BF
    UserControl.Line (0, rec.top - 1)-(lvwList.Width - 1, rec.Bottom - 1), tListStyle.SelectBorderColor, B
    'draw text display
    Call DrawText(UserControl.hDC, lvwList.ListItems(index).text, Len(lvwList.ListItems(index).text), rec, DT_LEFT)
    'draw time lenght
    rec.Left = colText
    rec.Right = colText + colTime '- 2
    Call DrawText(UserControl.hDC, lvwList.ListItems(index).SubItems(1), Len(lvwList.ListItems(index).SubItems(1)), rec, DT_RIGHT)
End Sub
Public Sub DrawList()
    On Error Resume Next
    Dim rec As RECT
    Dim x As Long
    Dim Y As Long
    
    UserControl.Cls
    If lvwList.ListItems.Count = 0 Then Exit Sub
    UserControl.ForeColor = tListStyle.ForeColor
    UserControl.FontBold = False
    Y = lvwList.GetFirstVisible.index
    For x = Y To Y + ItemPerPage
        If x > lvwList.ListItems.Count Then Exit For
        If lvwList.ListItems(x).Selected = False Then
        rec.top = lvwList.ListItems(x).top
        rec.Left = 0
        rec.Bottom = rec.top + lvwList.ListItems(x).Height
        rec.Right = rec.Left + colText
        'draw text display
        Call DrawText(UserControl.hDC, lvwList.ListItems(x).text, Len(lvwList.ListItems(x).text), rec, DT_LEFT)
        'draw time lenght
        rec.Left = colText
        rec.Right = colText + colTime
        Call DrawText(UserControl.hDC, lvwList.ListItems(x).SubItems(1), Len(lvwList.ListItems(x).SubItems(1)), rec, DT_RIGHT)
        End If
    Next x
End Sub
Public Property Get CurrentItem() As Long
    On Error Resume Next
    CurrentItem = lvwList.ListItems(CurrentIndex).index
End Property
Public Property Get CurrentPlayItem() As Long
    On Error Resume Next
    CurrentPlayItem = lvwList.ListItems(currentPlay).index
End Property
Public Property Let CurrentPlayItem(NewVal As Long)
    On Error Resume Next
    currentPlay = NewVal
    Call Play
End Property
Public Property Get Key(index As Long) As Long
    On Error Resume Next
    Key = lvwList.ListItems(index).SubItems(3)
End Property
Public Property Let Key(index As Long, NewVal As Long)
    On Error Resume Next
     lvwList.ListItems(index).SubItems(3) = NewVal
End Property
Public Property Get ListItemSelect(index As Long) As Boolean
    ListItemSelect = lvwList.ListItems(index).Selected
End Property
Public Property Let ListItemSelect(index As Long, NewVal As Boolean)
    lvwList.ListItems(index).Selected = NewVal
    Call DisplayList
    PropertyChanged "ListItemSelect"
End Property
Public Property Get ListItemTime(index As Long) As String
    On Error Resume Next
    ListItemTime = lvwList.ListItems(index).SubItems(1)
End Property
Public Property Let ListItemTime(index As Long, NewVal As String)
    On Error Resume Next
    lvwList.ListItems(index).SubItems(1) = NewVal
    Call DisplayList
    PropertyChanged "ListItemTime"
End Property

Public Property Get ListItemText(index As Long) As String
    On Error Resume Next
    ListItemText = lvwList.ListItems(index).text
End Property
Public Property Let ListItemText(index As Long, NewVal As String)
    On Error Resume Next
    lvwList.ListItems(index).text = NewVal
    Call DisplayList
    PropertyChanged "ListItemText"
End Property

Private Sub Play()
    Dim rec As RECT
    
    If currentPlay = 0 Or currentPlay > lvwList.ListItems.Count Then Exit Sub
    
    UserControl.ForeColor = tListStyle.PlayForeColor
    UserControl.FontBold = tListStyle.PlayBold
    
    rec.top = lvwList.ListItems(currentPlay).top
    rec.Left = 0
    rec.Bottom = rec.top + lvwList.ListItems(currentPlay).Height
    rec.Right = rec.Left + colText
    
    UserControl.Line (0, rec.top - 1)-(lvwList.Width, rec.Bottom - 1), tListStyle.PlayColor, BF
    'draw text display
    Call DrawText(UserControl.hDC, lvwList.ListItems(currentPlay).text, Len(lvwList.ListItems(currentPlay).text), rec, DT_LEFT)
    'draw time lenght
    rec.Left = colText
    rec.Right = colText + colTime
    Call DrawText(UserControl.hDC, lvwList.ListItems(currentPlay).SubItems(1), Len(lvwList.ListItems(currentPlay).SubItems(1)), rec, DT_RIGHT)
End Sub
Public Function FindItem(strTextFind As String) As ListItem
    On Error GoTo ErrHandle
    
    FindItem = lvwList.FindItem(strTextFind)
ErrHandle:
    FindItem = -1
    Exit Function
End Function
Public Function ItemPerPage() As Integer
    On Error Resume Next
    If lvwList.ListItems.Count = 0 Then ItemPerPage = 0: Exit Function
    ItemPerPage = CInt(lvwList.Height / lvwList.ListItems(1).Height)
End Function
Public Function ItemHeight() As Integer
    On Error Resume Next
    If lvwList.ListItems.Count = 0 Then ItemHeight = UserControl.TextHeight("|"): Exit Function
    ItemHeight = CInt(lvwList.ListItems(1).Height)
End Function
Public Sub RemoveItem(index As Long)
    On Error Resume Next
    Dim i As Long
    
    If index = 0 Or index > lvwList.ListItems.Count Then Exit Sub
    lvwList.ListItems.Remove index
    
    If tListStyle.ShowNumber Then
        For i = 1 To lvwList.ListItems.Count
            Call ReNumber(i)
        Next i
    End If
    
    If currentPlay > index Then currentPlay = currentPlay - 1
    If CurrentIndex > index Then currentPlay = CurrentIndex - 1
    
    Call DisplayList
End Sub
Public Sub ClearItem()
    On Error Resume Next
    Dim i As Long
    ReDim List(0)
    lvwList.ListItems.Clear
    CurrentIndex = 0
    currentPlay = 0
    UserControl.Cls
End Sub
Public Sub SelectItem(strType As String)
    Dim x As Integer
    Select Case LCase(strType)
        Case "all"
            For x = 1 To lvwList.ListItems.Count
                lvwList.ListItems(x).Selected = True
            Next x
        Case "none"
            For x = 1 To lvwList.ListItems.Count
                lvwList.ListItems(x).Selected = False
            Next x
        Case "invert"
            For x = 1 To lvwList.ListItems.Count
                If lvwList.ListItems(x).Selected = False Then
                    lvwList.ListItems(x).Selected = True
                Else
                    lvwList.ListItems(x).Selected = False
                End If
            Next x
    End Select
    Call DisplayList
End Sub
Private Sub ShowNumber(lngNumber As Long)
    lvwList.ListItems(lngNumber).text = Format(lngNumber, "000") & ". " & lvwList.ListItems(lngNumber).text
End Sub
Private Sub ReNumber(lngNumber As Long)
    lvwList.ListItems(lngNumber).text = Format(lngNumber, "000") & "." & Mid(lvwList.ListItems(lngNumber).text, InStr(1, lvwList.ListItems(lngNumber).text, ". ", vbBinaryCompare) + 1)
End Sub

Private Sub UserControl_Resize()
    Call drawPic
    lvwList.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    colText = UserControl.ScaleWidth - colTime
    If lvwList.ListItems.Count > 0 Then
        Call DisplayList
    End If
    RaiseEvent Resize
End Sub
Private Sub drawPic()
    On Error Resume Next
    
    Set UserControl.Picture = Nothing
    
    UserControl.Cls
    
    Select Case ePicAlign
        Case Is = picBottomLeft
            UserControl.PaintPicture pic.Picture, 0, UserControl.ScaleHeight - pic.Height, pic.ScaleWidth, pic.ScaleHeight, 0, 0
        Case Is = picBottomRight
            UserControl.PaintPicture pic.Picture, UserControl.ScaleWidth - pic.Width, UserControl.ScaleHeight - pic.Height, pic.ScaleWidth, pic.ScaleHeight, 0, 0
        Case Is = picCenter
            UserControl.PaintPicture pic.Picture, (UserControl.ScaleWidth - pic.Width) / 2, (UserControl.ScaleHeight - pic.Height) / 2, pic.Width, pic.Height, 0, 0, pic.Width, pic.Height, vbSrcCopy
        Case Is = picStrech
            UserControl.PaintPicture pic.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0, 0
        Case Is = picTitle
            Dim x As Long, Y As Long
                For x = 0 To UserControl.ScaleWidth Step pic.Width
                    For Y = 0 To UserControl.ScaleHeight Step pic.Height
                        UserControl.PaintPicture pic.Picture, x, Y, pic.Width, pic.Height, 0, 0
                    Next Y
                Next x
        Case Is = picTopLeft
            UserControl.PaintPicture pic.Picture, 0, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0
        Case Is = picTopRight
            UserControl.PaintPicture pic.Picture, UserControl.ScaleWidth - pic.Width, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0
    End Select
    Set UserControl.Picture = UserControl.Image
End Sub

