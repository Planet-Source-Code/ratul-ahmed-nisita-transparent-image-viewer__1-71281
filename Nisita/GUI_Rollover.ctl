VERSION 5.00
Begin VB.UserControl GUI_Rollover 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ControlContainer=   -1  'True
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   ToolboxBitmap   =   "GUI_Rollover.ctx":0000
   Begin VB.Image Image1 
      Height          =   15
      Left            =   225
      Top             =   1470
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "GUI_Rollover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private m_ImageDisabled      As Picture
Private m_ImageDown          As Picture
Private m_ImageHover         As Picture
Private m_ImageMask          As Picture
Private m_ImageNormal        As Picture
Private m_ImageSelected      As Picture
Private m_ImageSelectedHover As Picture

Private m_Enabled            As Boolean
Private m_MaskColor          As OLE_COLOR
Private m_Selected           As Boolean
Private m_Selectable         As Boolean

Event OnMouseClick()
Event OnMouseEnter()
Event OnMouseLeave()
Event OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event onMouseSelectionChange()

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ((X < 0) Or (X > UserControl.ScaleWidth) Or (Y < 0) Or (Y > UserControl.ScaleHeight)) Then
        '//Mouse Leave
        If Not Enabled Then Exit Sub
        If selected Then
            UserControl.Picture = ImageSelected
        Else
            UserControl.Picture = ImageNormal
        End If
        RaiseEvent OnMouseLeave
        Call ReleaseCapture
    ElseIf GetCapture <> UserControl.hWnd Then
        '//Mouse Hover
        If Not Enabled Then Exit Sub
        If selected Then
            UserControl.Picture = ImageSelectedHover
        Else
            UserControl.Picture = ImageHover
        End If
        RaiseEvent OnMouseEnter
        Call SetCapture(UserControl.hWnd)
    Else
        '//Mouse Move
        RaiseEvent OnMouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//Mouse Click
    If Not Enabled Then Exit Sub
    If selectable Then
        selected = Not selected
        RaiseEvent onMouseSelectionChange
    End If
    If selected Then
        UserControl.Picture = ImageSelectedHover
    Else
        UserControl.Picture = ImageHover
    End If
    RaiseEvent OnMouseClick
    Call SetCapture(UserControl.hWnd)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Enabled Then Exit Sub
    UserControl.Picture = ImageDown
End Sub

Private Sub RefreshImage()
    If selected Then
        UserControl.Picture = ImageSelected
    Else
        UserControl.Picture = ImageNormal
    End If
    UserControl.MaskPicture = ImageMask
    SizeControl
End Sub

Private Sub SizeControl()
    If Image1.Width > 1 And Image1.Height > 1 Then
        UserControl.Size Image1.Width * 15, Image1.Height * 15
    End If
End Sub


'// Start Custom Properties ---------------------------------------------------------------------
'//
'// Enabled
Public Property Let Enabled(ByVal value As Boolean)
    m_Enabled = value
    If value = False Then
        UserControl.Picture = ImageDisabled
    Else
        UserControl.Picture = ImageNormal
    End If
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
'
'// Selectable
Public Property Let selectable(ByVal value As Boolean)
    m_Selectable = value
End Property
Public Property Get selectable() As Boolean
    selectable = m_Selectable
End Property
'
'// Selected
Public Property Let selected(ByVal value As Boolean)
    m_Selected = value
End Property
Public Property Get selected() As Boolean
    selected = m_Selected
End Property
'
'// MaskColor
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property
Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
End Property
'
'// ImageDisabled
Public Property Get ImageDisabled() As Picture
    Set ImageDisabled = m_ImageDisabled
End Property
Public Property Set ImageDisabled(ByVal New_ImageDisabled As Picture)
    Set m_ImageDisabled = New_ImageDisabled
    PropertyChanged "ImageDisabled"
End Property
Public Property Let ImageDisabled(ByVal New_ImageDisabled As Picture)
    Set m_ImageDisabled = New_ImageDisabled
    PropertyChanged "ImageDisabled"
End Property
'
'// ImageDown
Public Property Get ImageDown() As Picture
    Set ImageDown = m_ImageDown
End Property
Public Property Set ImageDown(ByVal New_ImageDown As Picture)
    Set m_ImageDown = New_ImageDown
    PropertyChanged "ImageDown"
End Property
Public Property Let ImageDown(ByVal New_ImageDown As Picture)
    Set m_ImageDown = New_ImageDown
    PropertyChanged "ImageDown"
End Property
'
'// ImageMask
Public Property Get ImageMask() As Picture
    Set ImageMask = m_ImageMask
End Property
Public Property Set ImageMask(ByVal New_ImageMask As Picture)
    Set m_ImageMask = New_ImageMask
    Set UserControl.MaskPicture = m_ImageMask
    RefreshImage
    PropertyChanged "ImageMask"
End Property
Public Property Let ImageMask(ByVal New_ImageMask As Picture)
    Set m_ImageMask = New_ImageMask
    Set UserControl.MaskPicture = m_ImageMask
    RefreshImage
    PropertyChanged "ImageMask"
End Property
'
'// ImageHover
Public Property Get ImageHover() As Picture
    Set ImageHover = m_ImageHover
End Property
Public Property Set ImageHover(ByVal New_ImageHover As Picture)
    Set m_ImageHover = New_ImageHover
    PropertyChanged "ImageHover"
End Property
Public Property Let ImageHover(ByVal New_ImageHover As Picture)
    Set m_ImageHover = New_ImageHover
    PropertyChanged "ImageHover"
End Property
'
'// ImageNormal
Public Property Get ImageNormal() As Picture
    Set ImageNormal = m_ImageNormal
End Property
Public Property Set ImageNormal(ByVal New_ImageNormal As Picture)
    Set m_ImageNormal = New_ImageNormal
    Image1.Picture = ImageNormal
    RefreshImage
    PropertyChanged "ImageNormal"
End Property
Public Property Let ImageNormal(ByVal New_ImageNormal As Picture)
    Set m_ImageNormal = New_ImageNormal
    Image1.Picture = ImageNormal
    RefreshImage
    PropertyChanged "ImageNormal"
End Property
'
'// ImageSelected
Public Property Get ImageSelected() As Picture
    Set ImageSelected = m_ImageSelected
End Property
Public Property Set ImageSelected(ByVal New_ImageSelected As Picture)
    Set m_ImageSelected = New_ImageSelected
    PropertyChanged "ImageSelected"
    RefreshImage
End Property
Public Property Let ImageSelected(ByVal New_ImageSelected As Picture)
    Set m_ImageSelected = New_ImageSelected
    PropertyChanged "ImageSelected"
    RefreshImage
End Property
'
'// ImageSelectedHover
Public Property Get ImageSelectedHover() As Picture
    Set ImageSelectedHover = m_ImageSelectedHover
End Property
Public Property Set ImageSelectedHover(ByVal New_ImageSelectedHover As Picture)
    Set m_ImageSelectedHover = New_ImageSelectedHover
    PropertyChanged "ImageSelectedHover"
End Property
Public Property Let ImageSelectedHover(ByVal New_ImageSelectedHover As Picture)
    Set m_ImageSelectedHover = New_ImageSelectedHover
    PropertyChanged "ImageSelectedHover"
End Property
'
'// End Custom Properties --------------------------------------------------------------------------

Private Sub UserControl_InitProperties()
    'Set Default settings when control is created on the form
    Enabled = True
    MaskColor = &HFF00FF
    selectable = True
    selected = False
End Sub

Private Sub UserControl_Resize()
    SizeControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    m_Selectable = PropBag.ReadProperty("Selectable", True)
    m_Selected = PropBag.ReadProperty("Selected", False)
    Set m_ImageNormal = PropBag.ReadProperty("ImageNormal", Nothing)
    Set m_ImageHover = PropBag.ReadProperty("ImageHover", Nothing)
    Set m_ImageDown = PropBag.ReadProperty("ImageDown", Nothing)
    Set m_ImageDisabled = PropBag.ReadProperty("ImageDisabled", Nothing)
    Set m_ImageMask = PropBag.ReadProperty("ImageMask", Nothing)
    Set m_ImageSelected = PropBag.ReadProperty("ImageSelected", Nothing)
    Set m_ImageSelectedHover = PropBag.ReadProperty("ImageSelectedHover", Nothing)
    RefreshImage
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", m_Enabled, True
    PropBag.WriteProperty "MaskColor", m_MaskColor, &HFF00FF
    PropBag.WriteProperty "Selectable", m_Selectable, True
    PropBag.WriteProperty "Selected", m_Selected, False
    PropBag.WriteProperty "ImageNormal", m_ImageNormal, Nothing
    PropBag.WriteProperty "ImageHover", m_ImageHover, Nothing
    PropBag.WriteProperty "ImageDown", m_ImageDown, Nothing
    PropBag.WriteProperty "ImageDisabled", m_ImageDisabled, Nothing
    PropBag.WriteProperty "ImageMask", m_ImageMask, Nothing
    PropBag.WriteProperty "ImageSelected", m_ImageSelected, Nothing
    PropBag.WriteProperty "ImageSelectedHover", m_ImageSelectedHover, Nothing
End Sub

Private Sub UserControl_Terminate()
    Set m_ImageNormal = Nothing
    Set m_ImageNormal = Nothing
    Set m_ImageHover = Nothing
    Set m_ImageDown = Nothing
    Set m_ImageDisabled = Nothing
    Set m_ImageMask = Nothing
End Sub

