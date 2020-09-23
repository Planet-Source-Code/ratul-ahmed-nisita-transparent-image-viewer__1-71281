VERSION 5.00
Begin VB.Form Frmmain 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Nisita"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000B&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Nisita"
   Picture         =   "FrmMain.frx":15162
   ScaleHeight     =   6480
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fal 
      Height          =   285
      Left            =   -120
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   7920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open With"
      Height          =   2415
      Left            =   1200
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox Fpath 
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   6255
      End
      Begin VB.TextBox delt 
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Text            =   """"
         Top             =   840
         Width           =   6255
      End
      Begin VB.TextBox cmst 
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   0
         Pattern         =   "*gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.wmf;*.emf;*.png;*.tga"
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox txtline 
         Height          =   285
         Left            =   6840
         TabIndex        =   14
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox selected 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox oldline 
         Height          =   285
         Left            =   6840
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5245
      Left            =   240
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   490
      Width           =   8100
   End
   Begin VB.CheckBox chkBiLinear 
      Caption         =   "Quality Sizing"
      Height          =   240
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Stretch Quality Option"
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1485
   End
   Begin Nisita_Simple.vbkToolTip vbkToolTip1 
      Left            =   5040
      Top             =   6000
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin Nisita_Simple.GUI_Rollover NextB 
      Height          =   330
      Left            =   8200
      TabIndex        =   5
      Top             =   6020
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      Selectable      =   0   'False
      ImageNormal     =   "FrmMain.frx":1DDA8
      ImageHover      =   "FrmMain.frx":2140C
      ImageDown       =   "FrmMain.frx":24B55
      ImageDisabled   =   "FrmMain.frx":281B9
      ImageMask       =   "FrmMain.frx":2B81D
      ImageSelected   =   "FrmMain.frx":2EE81
      ImageSelectedHover=   "FrmMain.frx":324E5
   End
   Begin Nisita_Simple.GUI_Rollover BackB 
      Height          =   330
      Left            =   7820
      TabIndex        =   4
      Top             =   6020
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      Selectable      =   0   'False
      ImageNormal     =   "FrmMain.frx":35B49
      ImageHover      =   "FrmMain.frx":391CD
      ImageDown       =   "FrmMain.frx":3C93A
      ImageDisabled   =   "FrmMain.frx":3FFBE
      ImageMask       =   "FrmMain.frx":43642
      ImageSelected   =   "FrmMain.frx":46CC6
      ImageSelectedHover=   "FrmMain.frx":4A34A
   End
   Begin Nisita_Simple.GUI_Rollover Fulls 
      Height          =   330
      Left            =   7340
      TabIndex        =   3
      Top             =   6040
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      Selectable      =   0   'False
      ImageNormal     =   "FrmMain.frx":4D9CE
      ImageHover      =   "FrmMain.frx":50EE2
      ImageDown       =   "FrmMain.frx":54633
      ImageDisabled   =   "FrmMain.frx":57B47
      ImageMask       =   "FrmMain.frx":5B05B
      ImageSelected   =   "FrmMain.frx":5E56F
      ImageSelectedHover=   "FrmMain.frx":61A83
   End
   Begin Nisita_Simple.GUI_Rollover Comp 
      Height          =   330
      Left            =   6860
      TabIndex        =   2
      Top             =   6030
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      Selectable      =   0   'False
      ImageNormal     =   "FrmMain.frx":64F97
      ImageHover      =   "FrmMain.frx":684B7
      ImageDown       =   "FrmMain.frx":6BAE1
      ImageDisabled   =   "FrmMain.frx":6F001
      ImageMask       =   "FrmMain.frx":72521
      ImageSelected   =   "FrmMain.frx":75A41
      ImageSelectedHover=   "FrmMain.frx":78F61
   End
   Begin Nisita_Simple.GUI_Rollover OpenB 
      Height          =   330
      Left            =   6380
      TabIndex        =   1
      Top             =   6030
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      Selectable      =   0   'False
      ImageNormal     =   "FrmMain.frx":7C481
      ImageHover      =   "FrmMain.frx":7F9C5
      ImageDown       =   "FrmMain.frx":83045
      ImageDisabled   =   "FrmMain.frx":86589
      ImageMask       =   "FrmMain.frx":89ACD
      ImageSelected   =   "FrmMain.frx":8D011
      ImageSelectedHover=   "FrmMain.frx":90555
   End
   Begin Nisita_Simple.GUI_Rollover But_Exit 
      Height          =   300
      Left            =   8160
      TabIndex        =   0
      Top             =   15
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   529
      ImageNormal     =   "FrmMain.frx":93A99
      ImageHover      =   "FrmMain.frx":96F2C
      ImageDown       =   "FrmMain.frx":9A3E8
      ImageDisabled   =   "FrmMain.frx":9D87B
      ImageMask       =   "FrmMain.frx":A0D0E
      ImageSelected   =   "FrmMain.frx":A41A1
      ImageSelectedHover=   "FrmMain.frx":A7634
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   45
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   5295
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Format : Select a file first"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6200
      Width           =   5535
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================
'==Name: Nisita Simaple image viewer       =
'==Coder: Ratul Ahmed                      =
'==Thanx to :LaVolpe                       =
'===========================================


Option Explicit
' SAMPLE FORM ONLY, used to expose many of the c32bppDIB class options/capabilities

' Unicode-aware Open/Save Dialog box
' ////////////////////////////////////////////////////////////////
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> [Image Configuration Here] <<<
Private Type OPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     Flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (lpString As Any) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Const OFN_DONTADDTORECENT As Long = &H2000000
Private Const OFN_ENABLESIZING As Long = &H800000
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_NOCHANGEDIR As Long = &H8
' ////////////////////////////////////////////////////////////////
Private cImage As c32bppDIB
Private cShadow As c32bppDIB
Private m_GDItoken As Long

Private Sub chkBiLinear_Click()
chkBiLinear.Value = 1
cImage.HighQualityInterpolation = chkBiLinear.Value
RefreshImage
End Sub

Private Sub RefreshImage()
    Dim newWidth As Long, newHeight As Long
    Dim mirrorOffsetX As Long, mirrorOffsetY As Long
    Dim negAngleOffset As Long
    Dim X As Long, Y As Long
    Dim ShadowOffset As Long
    Dim LightAdjustment As Single
    
    mirrorOffsetX = 1
    mirrorOffsetY = 1
   
    ShadowOffset = Val(0) + 2   ' set shadow's blur depth as needed
    cImage.ScaleImage Picture1.ScaleWidth, Picture1.ScaleHeight, newWidth, newHeight, _
    scaleDownAsNeeded
    X = (Picture1.ScaleWidth - newWidth) \ 2
    Y = (Picture1.ScaleHeight - newHeight) \ 2
    Picture1.Cls
    If Not cShadow Is Nothing Then
        Picture1.CurrentX = 20
        Picture1.CurrentY = 5
        Picture1.Print "See c32bppDIB.CreateDropShadow for more ": _
        Picture1.CurrentX = 20
        Picture1.Print "Color, Opacity, Blur Effect, ": Picture1.CurrentX = 20
        Picture1.Print "  and X,Y Position are adjustable"
    End If
    cImage.Render Picture1.hDC, X + newWidth \ 2, Y + newHeight \ 2, _
    newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
        Val(100), , , , Val(-1), LightAdjustment, , True
   Picture1.Refresh
    
End Sub

Private Sub ShowImage(Optional bRefresh As Boolean = True, Optional DragDropCutPast _
As Boolean)

    Dim sSource As Variant
    Dim cX As Long, cY As Long
    Dim pto As Variant
  Me.Visible = False
  frmwarn.Visible = True
                sSource = OpenSaveFileDialog(False, "Select Image")
                Text1 = sSource
                pto = Text1
                fal = Text1
                Text2 = Mid(Text1.Text, 1, InStrRev(Trim(Text1.Text), "\"))
                lblName = "Name : " & Mid(Text1.Text, InStrRev(Text1.Text, "\") + 1)
                cImage.LoadPicture_File pto, 256, 256
                RefreshImage

    Select Case cImage.ImageType ' want to know source image's format?
        Case imgNone, imgError:     lblType.Caption = "Image was not loaded"
        Case imgBitmap:             lblType.Caption = "Format: Standard Bitmap or JPG"
        Case imgEMF:                lblType.Caption = "Format: Extended Windows Metafile"
        Case imgWMF:                lblType.Caption = "Format: Standard Windows Metafile"
        Case imgIcon:               lblType.Caption = "Format: Standard Icon"
        Case imgBmpARGB:            lblType.Caption = "Format: 32bpp Bitmap with ARGB"
        Case imgBmpPARGB:           lblType.Caption = "Format: 32bpp Bitmap with pARGB"
        Case imgCursor:             lblType.Caption = "Format: Standard Cursor"
        Case imgCursorARGB:         lblType.Caption = "Format: Alpha Cursor"
        Case imgIconARGB:           lblType.Caption = "Format: Alpha Icon"
        Case imgPNG:                lblType.Caption = "Format: PNG"
        Case imgPNGicon:            lblType.Caption = "Format: PNG in Vista Icon"
        Case imgGIF
            If cImage.Alpha = True Then
                                    lblType.Caption = "Format: Transparent GIF"
            Else
                                    lblType.Caption = "Format: GIF"
            End If
        Case imgTGA
            If cImage.Alpha = True Then
                                    lblType.Caption = "Format: Transparent TGA"
            Else
                                    lblType.Caption = "Format: TGA (Targa)"
            End If
        Case Else:                  lblType.Caption = "..."
    End Select
    
    If cImage.ImageType > imgNone Then
        lblType.Caption = lblType.Caption & " {" & cImage.Width & " x " & cImage.Height & "}"
    End If
    
    If Not cShadow Is Nothing Then
        CreateNewShadowClass
    Else
        If bRefresh Then RefreshImage
    End If
   frmwarn.Visible = False
   Me.Visible = True
   
End Sub
Private Function OpenSaveFileDialog(bSave As Boolean, DialogTitle As String, Optional DefaultExt As String, Optional SingleFilter As Boolean) As String

    ' using API version vs commondialog enables Unicode filenames to be passed to c32bppDIB classes
    Dim ofn As OPENFILENAME
    Dim rtn As Long
    Dim bUnicode As Boolean
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hwnd
        .hInstance = App.hInstance
        If SingleFilter Then
            Select Case DefaultExt
            Case "png"
                .lpstrFilter = "PNG" & vbNullChar & "*.png" & vbNullChar
            Case "jpg"
                .lpstrFilter = "JPG" & vbNullChar & "*.jpg" & vbNullChar
            Case "tga"
                .lpstrFilter = "TGA (Targa)" & vbNullChar & "*.tga" & vbNullChar
            End Select
        Else
            .lpstrFilter = "Image Files" & vbNullChar & "*gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.wmf;*.emf;*.png;*.tga"
            If cImage.isGDIplusEnabled Then
                .lpstrFilter = .lpstrFilter & ";*.tiff"
            End If
            .lpstrFilter = .lpstrFilter & vbNullChar & "Bitmaps" & vbNullChar & "*.bmp" & vbNullChar & "GIFs" & vbNullChar & "*.gif" & vbNullChar & _
                            "Icons/Cursors" & vbNullChar & "*.ico;*.cur" & vbNullChar & "JPGs" & vbNullChar & "*.jpg;*.jpeg" & vbNullChar & _
                            "Meta Files" & vbNullChar & "*.wmf;*.emf" & vbNullChar & "PNGs" & vbNullChar & "*.png" & vbNullChar & "TGAs (Targa)" & vbNullChar & "*.tga" & vbNullChar
            If cImage.isGDIplusEnabled Then
                .lpstrFilter = .lpstrFilter & "TIFFs" & vbNullChar & "*.tiff" & vbNullChar
            End If
            .lpstrFilter = .lpstrFilter & "All Files" & vbNullChar & "*.*" & vbNullChar
        End If
        .lpstrDefExt = DefaultExt
        .lpstrFile = String$(256, 0)
        .nMaxFile = 256
        .nMaxFileTitle = 256
        .lpstrTitle = DialogTitle
        .Flags = OFN_LONGNAMES Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_DONTADDTORECENT _
                Or OFN_NOCHANGEDIR
        ' ^^ don't want to change paths otherwise VB IDE locks folder until IDE is closed
        If bSave Then
            .Flags = .Flags Or OFN_CREATEPROMPT Or OFN_OVERWRITEPROMPT
        Else
            .Flags = .Flags Or OFN_FILEMUSTEXIST
        End If
    
        bUnicode = Not (IsWindowUnicode(GetDesktopWindow) = 0&)
        If bUnicode Then
            .lpstrInitialDir = StrConv(.lpstrInitialDir, vbUnicode)
            .lpstrFile = StrConv(.lpstrFile, vbUnicode)
            .lpstrFilter = StrConv(.lpstrFilter, vbUnicode)
            .lpstrTitle = StrConv(.lpstrTitle, vbUnicode)
            .lpstrDefExt = StrConv(.lpstrDefExt, vbUnicode)
        End If
        .lpstrFileTitle = .lpstrFile
    End With
    
    If bUnicode Then
        If bSave Then
            rtn = GetSaveFileNameW(ofn)
        Else
            rtn = GetOpenFileNameW(ofn)
        End If
        If rtn > 0& Then
            If bUnicode Then
                rtn = lstrlenW(ByVal ofn.lpstrFile)
                OpenSaveFileDialog = StrConv(Left$(ofn.lpstrFile, rtn * 2), vbFromUnicode)
            End If
        End If
    Else
        If bSave Then
            rtn = GetSaveFileName(ofn)
        Else
            rtn = GetOpenFileName(ofn)
        End If
        If rtn > 0& Then
            rtn = lstrlen(ofn.lpstrFile)
            OpenSaveFileDialog = Left$(ofn.lpstrFile, rtn)
        End If
    End If

ExitRoutine:
End Function
Private Sub CreateNewShadowClass()
    Dim blurDepth As Long
    Dim Color As Long
    blurDepth = 2
    Color = 2
    Set cShadow = cImage.CreateDropShadow(blurDepth, Color)
    cShadow.gdiToken = m_GDItoken   ' assign shared token if one exists
    RefreshImage
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> [Image Configuration end Here] <<<



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> [ Form Configurations Here ] <<<
Private Sub Form_Load()
Picture1.AutoRedraw = True
    Set cImage = New c32bppDIB
    cImage.InitializeDIB ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, vbPixels), ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, vbPixels)
    cImage.CreateCheckerBoard 20, vbWhite
    cImage.Render Picture1.hDC
    cImage.DestroyDIB
    Picture1.Picture = Picture1.Image
    Show
    
cmst = Command
RemoveString cmst, delt
Dim pto As Variant
pto = Fpath
Text1 = Fpath
fal = Text1
Text2 = Mid(Text1.Text, 1, InStrRev(Trim(Text1.Text), "\"))
lblName = "Name : " & Mid(Text1.Text, InStrRev(Text1.Text, "\") + 1)
cImage.LoadPicture_File pto, 256, 256
RefreshImage
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    ' when you create a token to be shared, you must
    ' destroy it in the Unload or Terminate event
    ' and also reset gdiToken property for each existing class
    If m_GDItoken Then
        If Not cShadow Is Nothing Then cShadow.gdiToken = 0&
        If Not cImage Is Nothing Then
            cImage.gdiToken = 0&
            cImage.DestroyGDIplusToken m_GDItoken
        End If
    End If
    End
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> [ Form Configurations end Here ] <<<
Public Sub RemoveString(Entire As String, Word As String)
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
        I = InStr(1, Entire, Word)
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, I - 1)
            Entire = LeftPart & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    Fpath = Entire
    
End Sub




'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>[ Button Configurations Here ] <<<
Private Sub Comp_OnMouseClick()
Picture1_Click
Picture1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then NextB_OnMouseClick
If KeyCode = vbKeyReturn Then Picture1_DblClick
Picture1.SetFocus
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
vbkToolTip1.Fire Label1, "About", _
"Nisita Image viewer © Ratul Ahmed                  Anchor Soft 2008" _
 , 181, 50

End Sub

Private Sub ckk()
    Select Case cImage.ImageType ' want to know source image's format?
        Case imgNone, imgError:     lblType.Caption = "Image was not loaded"
        Case imgBitmap:             lblType.Caption = "Format: Standard Bitmap or JPG"
        Case imgEMF:                lblType.Caption = "Format: Extended Windows Metafile"
        Case imgWMF:                lblType.Caption = "Format: Standard Windows Metafile"
        Case imgIcon:               lblType.Caption = "Format: Standard Icon"
        Case imgBmpARGB:            lblType.Caption = "Format: 32bpp Bitmap with ARGB"
        Case imgBmpPARGB:           lblType.Caption = "Format: 32bpp Bitmap with pARGB"
        Case imgCursor:             lblType.Caption = "Format: Standard Cursor"
        Case imgCursorARGB:         lblType.Caption = "Format: Alpha Cursor"
        Case imgIconARGB:           lblType.Caption = "Format: Alpha Icon"
        Case imgPNG:                lblType.Caption = "Format: PNG"
        Case imgPNGicon:            lblType.Caption = "Format: PNG in Vista Icon"
        Case imgGIF
            If cImage.Alpha = True Then
                                    lblType.Caption = "Format: Transparent GIF"
            Else
                                    lblType.Caption = "Format: GIF"
            End If
        Case imgTGA
            If cImage.Alpha = True Then
                                    lblType.Caption = "Format: Transparent TGA"
            Else
                                    lblType.Caption = "Format: TGA (Targa)"
            End If
        Case Else:                  lblType.Caption = "..."
    End Select
End Sub

Private Sub OpenB_OnMouseClick()
On Error Resume Next
ShowImage
Picture1.SetFocus
selected = ""
End Sub
Private Sub But_Exit_OnMouseClick()
Unload Me
End Sub
Private Sub BackB_OnMouseClick()
On Error Resume Next
If txtline = oldline - 1 Then
lblType.Caption = "Format : Select a file first"
Else
txtline = txtline + 1
selected.Text = LineText(Text3, txtline)

Dim pto As Variant
pto = Text2 & "\" & selected
cImage.LoadPicture_File pto, 256, 256
RefreshImage
End If
Picture1.SetFocus
ckk
End Sub
Private Sub Fulls_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 vbkToolTip1.Fire Comp, "Full Size", "Show image In Full Size" _
 , 120, 45
End Sub

Private Sub OpenB_OnMouseMove(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
 vbkToolTip1.Fire OpenB, "Open File", "Open Image Files" _
 & vbNewLine & "Png,bmp,icon,jpg," & vbNewLine & "gif And more ...", 100, 75
End Sub
Private Sub BackB_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 vbkToolTip1.Fire BackB, "Previous Image", "Show you Previous image" _
 , 130, 45
End Sub

Private Sub But_Exit_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 vbkToolTip1.Fire But_Exit, "Exit", "    Kill Nisita" _
 , 70, 45
End Sub

Private Sub Comp_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 vbkToolTip1.Fire Comp, "Compact Size", "Show image In Copmact Size" _
 , 145, 45
End Sub
Private Sub NextB_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 vbkToolTip1.Fire NextB, "Next Image", "Show you Next image" _
 , 115, 45
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   
    If Button = 1 Then
    
        'Release capture
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        
    End If
End Sub

Private Sub Fulls_OnMouseClick()
Picture1_DblClick
Picture1.SetFocus
End Sub

Private Sub NextB_OnMouseClick()
On Error Resume Next
If txtline <= 1 Then
lblType.Caption = "Format : Select a file first"
Else
txtline = txtline - 1
selected.Text = LineText(Text3, txtline)
Dim pto As Variant
pto = Text2 & "\" & selected
fal = pto
cImage.LoadPicture_File pto, 256, 256
RefreshImage
End If
Picture1.SetFocus
ckk
End Sub

Private Sub Picture1_Click()
Dim pto As Variant
pto = fal

If Me.Height = Screen.Height Then
Me.Visible = False
Me.Height = 6480
Me.Width = 8610
Picture1.Height = 5245
Picture1.Width = 8100
Me.Left = 0
Me.Top = 0
Picture1.Left = 240
Picture1.Top = 490
Set cImage = New c32bppDIB
    cImage.InitializeDIB ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, vbPixels), ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, vbPixels)
    cImage.CreateCheckerBoard 20, vbWhite
    cImage.Render Picture1.hDC
    cImage.DestroyDIB
    Picture1.Picture = Picture1.Image
    Show
cImage.LoadPicture_File pto, 256, 256
RefreshImage
Me.Visible = True
Else
lblType.Caption = "Format : Select a file first"
Picture1.SetFocus
End If
ckk

End Sub

Private Sub Picture1_DblClick()
Dim pto As Variant
pto = fal
If Me.Height = Screen.Height Then
Me.Visible = False
Me.Height = 6480
Me.Width = 8610
Picture1.Height = 5245
Picture1.Width = 8100
Me.Left = 0
Me.Top = 0
Picture1.Left = 240
Picture1.Top = 490
Set cImage = New c32bppDIB
    cImage.InitializeDIB ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, vbPixels), ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, vbPixels)
    cImage.CreateCheckerBoard 20, vbWhite
    cImage.Render Picture1.hDC
    cImage.DestroyDIB
    Picture1.Picture = Picture1.Image
    Show
cImage.LoadPicture_File pto, 256, 256
RefreshImage
Me.Visible = True
Else
Me.Visible = False
Me.Height = Screen.Height
Me.Width = Screen.Width
Picture1.Height = Me.Height
Picture1.Width = Me.Width
Me.Left = 0
Me.Top = 0
Picture1.Left = 0
Picture1.Top = 0

    Set cImage = New c32bppDIB
    cImage.InitializeDIB ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, vbPixels), ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, vbPixels)
    cImage.CreateCheckerBoard 20, vbWhite
    cImage.Render Picture1.hDC
    cImage.DestroyDIB
    Picture1.Picture = Picture1.Image
    Show
cImage.LoadPicture_File pto, 256, 256
RefreshImage
Me.Visible = True
Picture1.SetFocus
End If
ckk
End Sub



Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then NextB_OnMouseClick
If KeyCode = vbKeyReturn Then Picture1_DblClick
If KeyCode = vbKeyPageDown Then NextB_OnMouseClick
If KeyCode = vbKeyPageUp Then BackB_OnMouseClick
If KeyCode = vbKeyLeft Then BackB_OnMouseClick
If KeyCode = vbKeyRight Then NextB_OnMouseClick
If KeyCode = vbKeyEscape Then Picture1_DblClick
Picture1.SetFocus
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Simple example of pasting file names
    If KeyCode = vbKeyV Then
        
        If (Shift And vbCtrlMask) = vbCtrlMask Then
        
            Dim Files() As String
            ' use class to return unicode filenames if applicable
            If cImage.GetPastedFileNames(Files()) > 0 Then
            
                If cImage.LoadPicture_File(Files(1), 256, 256) = True Then
                    If Not cShadow Is Nothing Then
                        CreateNewShadowClass
                    Else
                        RefreshImage
                    End If
                    
                Else
                    MsgBox "Failed to load the image file", vbInformation + vbOKOnly
                End If
                
            Else    ' didn't paste a file name, see if it is a clipboard image?
            
                If cImage.LoadPicture_ClipBoard = True Then
                    If Not cShadow Is Nothing Then
                        CreateNewShadowClass
                    Else
                        RefreshImage
                    End If
                    
                Else
                    MsgBox "Failed to load the clipboard image", vbInformation + vbOKOnly
                End If
                
            End If
        
        End If
    End If
    Picture1.SetFocus
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' simmple OLE drag/drop example
    
    ' use class to return unicode filenames if applicable
    If cImage.GetDroppedFileNames(Data) = True Then
        If cImage.LoadPicture_File(Data.Files(1), 256, 256) = True Then
            If Not cShadow Is Nothing Then
                CreateNewShadowClass
            Else
                RefreshImage
            End If
            
        End If
        
    End If
Picture1.SetFocus
End Sub

Private Sub selected_Change()
lblName.Caption = "Name : " & selected
End Sub

Private Sub Text2_Change()
On Error Resume Next
File1 = Text2
Dim X As Integer
Dim Abb As String
For X = 0 To File1.ListCount
        Abb = File1.List(X)
Text3 = Text3.Text & File1.List(X) & vbCrLf
Next X
End Sub

Private Sub Text3_Change()
Dim xxx As Integer
For xxx = 1 To LineCount(Text3) - 1
        txtline = xxx
        selected.Text = LineText(Text3, txtline)
Next xxx
oldline = txtline
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>[ Button Configurations Ends Here ] <<<


