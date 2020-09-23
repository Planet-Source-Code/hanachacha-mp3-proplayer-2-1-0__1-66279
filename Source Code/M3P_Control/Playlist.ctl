VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Playlist 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   ToolboxBitmap   =   "Playlist.ctx":0000
   Begin MSComctlLib.ListView lvwPL 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SortString"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "Playlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
                                        (ByVal hDC As Long, ByVal lpStr As String, _
                                        ByVal nCount As Long, lpRect As RECT, _
                                        ByVal wFormat As Long) As Long
Private Const DT_RIGHT = &H2
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Type ListInfor
    strText As String
    strTime As String
    strSort As String
    Selected As Boolean
End Type
Dim NowPlaying() As ListInfor


Private Type DemensionText
    BackColor As Long
    ForeColor As Long
    PlayBackColor As Long
    PlayForeColor As Long
    SelectedColor As Long
    ShowNumber As Boolean
End Type

Dim tPlaylistConfig As DemensionText
Dim bolRDown As Boolean
Dim i As Long

Private Const ColTime As Long = 50
Dim ColText As Long
Dim currentPlayIndex As Long
Dim CurrentIndex As Long
Dim intStart As Integer
Dim intEnd As Integer

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()
Public Event DblClick()


Private Sub DrawPlayIndex() 'Use for playing file
    On Error Resume Next
    Dim rec As RECT
    
    If currentPlayIndex = 0 Then
        Exit Sub
    End If
        
    UserControl.FontBold = True
    UserControl.ForeColor = tPlaylistConfig.PlayForeColor
    rec.Top = lvwPL.ListItems(currentPlayIndex).Top
    rec.Left = 0
    rec.Bottom = rec.Top + lvwPL.ListItems(currentPlayIndex).Height
    rec.Right = ColText
    UserControl.Line (0, rec.Top)-(rec.Right + ColTime, rec.Bottom), tPlaylistConfig.PlayBackColor, BF
    Call DrawText(UserControl.hDC, NowPlaying(lvwPL.ListItems(currentPlayIndex).Text).strText, Len(NowPlaying(lvwPL.ListItems(currentPlayIndex).Text).strText), rec, DT_LEFT)
    rec.Left = ColText
    rec.Right = ColText + ColTime
    Call DrawText(UserControl.hDC, NowPlaying(lvwPL.ListItems(currentPlayIndex).Text).strTime, Len(NowPlaying(lvwPL.ListItems(currentPlayIndex).Text).strTime), rec, DT_RIGHT)
End Sub
Private Sub DrawList() 'Support MultiSelect
    On Error Resume Next
    Dim rec As RECT
    Dim lngColor As Long
    Dim X As Long
    Dim Y As Long
    
    UserControl.Cls
    UserControl.BackColor = tPlaylistConfig.BackColor
    UserControl.ForeColor = tPlaylistConfig.ForeColor
    UserControl.FontBold = False
    Y = lvwPL.GetFirstVisible.index
    For X = Y To Y + ItemPage
        If X > lvwPL.ListItems.Count Then Exit For
        If lvwPL.ListItems(X).Selected = False Then
            rec.Top = lvwPL.ListItems(X).Top
            rec.Left = 0
            rec.Bottom = rec.Top + lvwPL.ListItems(X).Height
            rec.Right = rec.Left + ColText
            'draw text display
            Call DrawText(UserControl.hDC, NowPlaying(lvwPL.ListItems(X).Text).strText, Len(NowPlaying(lvwPL.ListItems(X).Text).strText), rec, DT_LEFT)
            'draw time lenght
            rec.Left = ColText
            rec.Right = ColText + ColTime
            Call DrawText(UserControl.hDC, NowPlaying(lvwPL.ListItems(X).Text).strTime, Len(NowPlaying(lvwPL.ListItems(X).Text).strTime), rec, DT_RIGHT)
        End If
    Next X
End Sub
Private Sub DrawSelected(index As Long)
    On Error Resume Next
    
    Dim rec As RECT
    Dim lngColor As Long
    Dim i As Long
    Dim DT As Long
    
    If index = 0 Or index > lvwPL.ListItems.Count Then Exit Sub
    If lvwPL.ListItems(index).Selected = False Then Exit Sub
    lngColor = tPlaylistConfig.SelectedColor
    UserControl.ForeColor = tPlaylistConfig.ForeColor
    UserControl.FontBold = False
    'Cal rec
    rec.Top = lvwPL.ListItems(index).Top
    rec.Left = 0
    rec.Bottom = rec.Top + lvwPL.ListItems(index).Height
    rec.Right = rec.Left + ColText
    'draw text display
    UserControl.Line (0, rec.Top)-(lvwPL.Width, rec.Bottom), lngColor, BF
    Call DrawText(UserControl.hDC, NowPlaying(lvwPL.ListItems(index).Text).strText, Len(NowPlaying(lvwPL.ListItems(index).Text).strText), rec, DT_LEFT)
    'draw time lenght
    rec.Left = ColText
    rec.Right = rec.Left + ColTime
    Call DrawText(UserControl.hDC, NowPlaying(lvwPL.ListItems(index).Text).strTime, Len(NowPlaying(lvwPL.ListItems(index).Text).strTime), rec, DT_RIGHT)
End Sub


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        tPlaylistConfig.BackColor = .ReadProperty("BackColor", vbWhite)
        tPlaylistConfig.ForeColor = .ReadProperty("BackColor", vbBlack)
        tPlaylistConfig.PlayBackColor = .ReadProperty("PlayBackColor", vbBlue)
        tPlaylistConfig.PlayForeColor = .ReadProperty("PlayForeColor", vbWhite)
        tPlaylistConfig.SelectedColor = .ReadProperty("SelectColor", vbGreen)
    End With
End Sub

Private Sub UserControl_Resize()
    lvwPL.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    ColText = lvwPL.Width - ColTime
    If lvwPL.ListItems.Count > 0 Then
        lvwPL.Height = (ItemPage * (lvwPL.ListItems(1).Height))
        DrawAll
    End If
End Sub
Private Sub UserControl_DblClick()
    On Error Resume Next
    If bolRDown = False Then
        currentPlayIndex = CurrentIndex
    End If
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 40 Then
        If CurrentIndex < lvwPL.ListItems.Count Then
            lvwPL.ListItems(CurrentIndex).Selected = False
            CurrentIndex = CurrentIndex + 1
            lvwPL.ListItems(CurrentIndex).Selected = True
        End If
    End If
    If KeyCode = 38 Then
        If CurrentIndex > 1 Then
            lvwPL.ListItems(CurrentIndex).Selected = False
            CurrentIndex = CurrentIndex - 1
            lvwPL.ListItems(CurrentIndex).Selected = True
        End If
    End If
    DrawAll
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        currentPlayIndex = CurrentIndex
    End If
    DrawAll
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If lvwPL.ListItems.Count = 0 Then
        Exit Sub
    Else
        Dim firstVisible As Long
        Dim tmp As Long
        firstVisible = lvwPL.GetFirstVisible.index
        If Button = vbRightButton Then
            bolRDown = True
        End If
        If Button = vbLeftButton Then
            bolRDown = False
            tmp = firstVisible + Y \ lvwPL.ListItems(1).Height
            If tmp > 0 And tmp <= lvwPL.ListItems.Count Then
                Select Case Shift
                    Case 0
                        For i = 1 To lvwPL.ListItems.Count
                            lvwPL.ListItems(i).Selected = False
                            NowPlaying(NowIndex(i)).Selected = False
                        Next i
                        intStart = 0
                        intEnd = 0
                        CurrentIndex = tmp
                        lvwPL.ListItems(CurrentIndex).Selected = True
                        NowPlaying(NowIndex(i)).Selected = True
                    Case 1 'shift
                        intEnd = tmp
                        intStart = CurrentIndex
                        If intStart < intEnd Then
                            For i = intStart To intEnd Step 1
                                lvwPL.ListItems(i).Selected = True
                                NowPlaying(NowIndex(i)).Selected = True
                            Next i
                        Else
                            For i = intStart To intEnd Step -1
                                lvwPL.ListItems(i).Selected = True
                                NowPlaying(NowIndex(i)).Selected = True
                            Next i
                        End If
                    Case 2
                        lvwPL.ListItems(tmp).Selected = Not lvwPL.ListItems(tmp).Selected
                        NowPlaying(NowIndex(tmp)).Selected = Not NowPlaying(NowIndex(tmp)).Selected
                        intStart = 0
                        intEnd = 0
                        If lvwPL.ListItems(tmp).Selected = True Then
                            CurrentIndex = tmp
                            lvwPL.ListItems(CurrentIndex).Selected = True
                            NowPlaying(NowIndex(CurrentIndex)).Selected = True
                        End If
                End Select
            End If
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim firstVisible As Long
    Dim tmp As Integer
       
    If lvwPL.ListItems.Count = 0 Then
        Exit Sub
    Else
        firstVisible = lvwPL.GetFirstVisible.index
        tmp = firstVisible + Y \ lvwPL.ListItems(1).Height
        If Button = vbRightButton Then
            bolRDown = False
            If tmp > 0 And tmp <= lvwPL.ListItems.Count Then
                CurrentIndex = tmp
                For i = 1 To lvwPL.ListItems.Count
                    lvwPL.ListItems(i).Selected = False
                Next i
                lvwPL.ListItems(CurrentIndex).Selected = True
                NowPlaying(NowIndex(CurrentIndex)).Selected = True
            End If
        End If
        If Button = vbLeftButton Then
            bolRDown = False
        End If
        If CurrentIndex = currentPlayIndex Then
            lvwPL.ListItems(CurrentIndex).Selected = False
            NowPlaying(NowIndex(CurrentIndex)).Selected = False
        End If
        DrawAll
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Public Property Set Font(ByVal NewFont As StdFont)
    Set UserControl.Font = NewFont
    Set lvwPL.Font = NewFont
    PropertyChanged "Font"
End Property
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    tPlaylistConfig.BackColor = NewColor
    PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = tPlaylistConfig.BackColor
End Property
Public Property Let PlayBackColor(ByVal NewColor As OLE_COLOR)
    tPlaylistConfig.PlayBackColor = NewColor
    UserControl.BackColor = NewColor
    PropertyChanged "PlayBackColor"
End Property
Public Property Get PlayBackColor() As OLE_COLOR
    PlayBackColor = tPlaylistConfig.PlayBackColor
End Property
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    tPlaylistConfig.ForeColor = NewColor
    PropertyChanged "ForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = tPlaylistConfig.ForeColor
End Property
Public Property Let PlayForeColor(ByVal NewColor As OLE_COLOR)
    tPlaylistConfig.PlayForeColor = NewColor
    PropertyChanged "PlayForeColor"
End Property
Public Property Get PlayForeColor() As OLE_COLOR
    PlayForeColor = tPlaylistConfig.PlayForeColor
End Property
Public Property Let SelectColor(ByVal NewColor As OLE_COLOR)
    tPlaylistConfig.SelectedColor = NewColor
    PropertyChanged "SelectColor"
End Property
Public Property Get SelectColor() As OLE_COLOR
    SelectColor = tPlaylistConfig.SelectedColor
End Property
Public Property Let Size(ByVal NewSize As Long)
    UserControl.Font.Size = NewSize
    PropertyChanged "Size"
End Property
Public Property Let ShowNumber(ByVal bol As Boolean)
    tPlaylistConfig.ShowNumber = bol
    PropertyChanged "ShowNumber"
    If bol Then
        For i = 1 To UBound(NowPlaying)
            Call Number(i)
        Next i
    End If
    DrawAll
End Property
Public Sub Add(Text As String, Time As String)
    ReDim Preserve NowPlaying(lvwPL.ListItems.Count + 1)
    
    lvwPL.ListItems.Add lvwPL.ListItems.Count + 1, , UBound(NowPlaying)
    NowPlaying(UBound(NowPlaying)).strText = Text
    NowPlaying(UBound(NowPlaying)).strTime = Time
    If tPlaylistConfig.ShowNumber Then Number (UBound(NowPlaying))
    DoSort
    If tPlaylistConfig.ShowNumber Then
        For i = 1 To UBound(NowPlaying)
            Call ReNumber(i)
        Next i
    End If
    DrawAll
End Sub
Public Sub EnsureVisible(index As Long)
    On Error Resume Next
    If index < 1 Or index > lvwPL.ListItems.Count Then Exit Sub
    lvwPL.ListItems(index).EnsureVisible
    DrawAll
End Sub
Public Function ItemPage() As Integer
    On Error Resume Next
    If lvwPL.ListItems.Count = 0 Then Exit Function
    ItemPage = lvwPL.Height / lvwPL.ListItems(1).Height
End Function

Private Sub DrawAll()
    DrawList
    Dim z As Long
    Dim r As Long
    z = (lvwPL.GetFirstVisible.index)
    If z <> 0 Then
        For r = z To z + ItemPage
            If r > lvwPL.ListItems.Count Then Exit For
            If lvwPL.ListItems(r).Selected = True Then
                Call DrawSelected(r)
            End If
        Next r
    End If
    DrawPlayIndex
End Sub
Public Sub DoSort()
    lvwPL.Sorted = False
    For i = 1 To lvwPL.ListItems.Count
        lvwPL.ListItems(i).SubItems(1) = NowPlaying(lvwPL.ListItems(i).Text).strSort
    Next i
    lvwPL.SortKey = 1
    lvwPL.Sorted = True
    DrawAll
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", tPlaylistConfig.BackColor, vbWhite)
        Call .WriteProperty("ForeColor", tPlaylistConfig.ForeColor, vbBlack)
        Call .WriteProperty("PlayBackColor", tPlaylistConfig.PlayBackColor, vbBlue)
        Call .WriteProperty("PlayForeColor", tPlaylistConfig.PlayForeColor, vbWhite)
        Call .WriteProperty("SelectColor", tPlaylistConfig.SelectedColor, vbGreen)
    End With
End Sub
Public Function ItemText(index As Long) As String
    On Error Resume Next
    ItemText = NowPlaying(index).strText
End Function
Public Sub ItemSetText(index As Long, str As String)
    On Error Resume Next
    NowPlaying(index).strText = str
    DrawAll
End Sub
Public Function ItemTime(index As Long) As String
    On Error Resume Next
    ItemTime = NowPlaying(index).strTime
End Function
Public Sub ItemSetTime(index As Long, str As String)
    On Error Resume Next
    NowPlaying(index).strTime = str
    DrawAll
End Sub
Public Sub ItemSetSort(index As Long, strSort As String)
    On Error Resume Next
    NowPlaying(index).strSort = strSort
End Sub
Public Function ItemGetSelect(index As Long) As Boolean
    ItemGetSelect = NowPlaying(index).Selected
End Function
Public Sub ItemSelect(index As Long)
    lvwPL.ListItems(index).Selected = True
    NowPlaying(NowIndex(index)).Selected = True
End Sub
Public Function NowIndex(index As Long) As Long
    If index < 1 Or index > lvwPL.ListItems.Count Then NowIndex = 0: Exit Function
    NowIndex = CLng(lvwPL.ListItems(index).Text)
End Function
Public Function CurrentItem() As Long
    CurrentItem = CurrentIndex
End Function
Public Function CurrentPlayItem() As Long
    CurrentPlayItem = currentPlayIndex
End Function
Public Property Let SortOrder(ByVal newValue As Integer)
    lvwPL.SortOrder = newValue
    PropertyChanged "SortOrder"
End Property
Public Property Get SortOrder() As Integer
     SortOrder = lvwPL.SortOrder
End Property

Public Sub Remove(index As Long)
    Dim X As Long, Y As Long
    If index < 1 And index > lvwPL.ListItems.Count Then Exit Sub
    X = lvwPL.ListItems(index).Text
    For Y = X To UBound(NowPlaying)
        NowPlaying(Y) = NowPlaying(Y + 1)
    Next Y
    ReDim Preserve NowPlaying(UBound(NowPlaying))
    lvwPL.ListItems.Remove index
    DrawAll
End Sub
Public Sub Clear()
    On Error Resume Next
    ReDim NowPlaying(0)
    lvwPL.ListItems.Clear
    UserControl.Cls
End Sub
Public Function Count() As Long
    On Error Resume Next
    Count = UBound(NowPlaying)
End Function
Private Sub Number(iNumber As Long)
    NowPlaying(lvwPL.ListItems(iNumber).Text).strText = Format(iNumber, "000") & ". " & NowPlaying(lvwPL.ListItems(iNumber).Text).strText
End Sub
Private Sub ReNumber(iNumber As Long)
    NowPlaying(iNumber).strText = Format(iNumber, "000") & "." & Mid(NowPlaying(iNumber).strText, InStr(1, NowPlaying(iNumber).strText, ". ", vbBinaryCompare) + 1)
End Sub
