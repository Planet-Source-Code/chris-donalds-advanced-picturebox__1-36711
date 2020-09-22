VERSION 5.00
Begin VB.UserControl cTransPictureBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   Picture         =   "cTransPicbox.ctx":0000
   PropertyPages   =   "cTransPicbox.ctx":00B4
   ScaleHeight     =   1680
   ScaleWidth      =   3045
   ToolboxBitmap   =   "cTransPicbox.ctx":00DA
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   1200
      ScaleHeight     =   555
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgSrc 
      Height          =   585
      Left            =   1800
      Top             =   660
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "cTransPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum STYLES
    [None] = 0
    [Repeat] = 1
    [Stretch] = 2
End Enum

Public Enum BORDER
    [None] = 0
    [Fixed Single] = 1
End Enum

Private Const m_def_Picture = "(none)"
Private Const m_def_TransColor = 0
Private Const m_def_BackColor = 0
Private Const m_def_Style = STYLES.[Repeat]
Private Const m_def_SyncTransColor = True

Private Const nPicture = "PictureFile"
Private Const nTransColor = "TransparentColor"
Private Const nBackColor = "BackColor"
Private Const nStyle = "Style"
Private Const nSyncTransColor = "SyncTransColor"
Private Const nBorderStyle = "BorderStyle"
Private Const nMouseIcon = "MouseIcon"
Private Const nMousePointer = "MousePointer"
Private Const nEnabled = "Enabled"
Private Const nWidth = "Width"
Private Const nHeight = "Height"

Dim pPicture As StdPicture
Dim pTransColor As OLE_COLOR
Dim pHwnd As Long
Dim pBackColor As OLE_COLOR
Dim pStyle As STYLES
Dim pSyncTransColor As Boolean

Public Event Click()
Public Event DblClick()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub About()
Attribute About.VB_Description = "Credits and copyrights"
Attribute About.VB_UserMemId = -552
    dlgAbout.Show vbModal
    Unload dlgAbout
    Set dlgAbout = Nothing
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    ' Set default control variables
    
    pTransColor = m_def_TransColor      ' Transparent color (mask color)
    pBackColor = m_def_BackColor        ' Background color
    pHwnd = hWnd                        ' Window handle of control
    pStyle = m_def_Style                ' Picture formating style
End Sub

Private Sub UserControl_InitProperties()
    ' Set initial visible properties
    
    Set PictureFile = Picture
    TransparentColor = m_def_TransColor
    BackColor = m_def_BackColor
    Style = m_def_Style
    SyncTransparentColor = m_def_SyncTransColor
End Sub

Public Property Get PictureFile() As StdPicture
Attribute PictureFile.VB_Description = "Sets/returns the picture used for this control"
    ' get the picture property of the control
    ' stored in the user control's picture property
    
    Set PictureFile = pPicture
End Property

Public Property Set PictureFile(new_pictureFile As StdPicture)
    ' This is where the bulk of the work is done
    
    drawImage new_pictureFile
    PropertyChanged "PictureFile"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, _
        ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
        ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, _
        ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
        ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, _
        ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
        ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Get properties from bag
    With PropBag
        Set PictureFile = .ReadProperty(nPicture, Picture)
        TransparentColor = .ReadProperty(nTransColor, m_def_TransColor)
        BackColor = .ReadProperty(nBackColor, m_def_BackColor)
        Style = .ReadProperty(nStyle, m_def_Style)
        SyncTransparentColor = .ReadProperty(nSyncTransColor, m_def_SyncTransColor)
        BorderStyle = .ReadProperty(nBorderStyle, BORDER.[Fixed Single])
        Set MouseIcon = .ReadProperty(nMouseIcon)
        MousePointer = .ReadProperty(nMousePointer)
        Enabled = .ReadProperty(nEnabled, True)
        Width = .ReadProperty(nWidth)
        Height = .ReadProperty(nHeight)
    End With
    drawImage pPicture, False
End Sub

Private Sub UserControl_Resize()
    drawImage pPicture, False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Write properties to bag
    With PropBag
        .WriteProperty nPicture, pPicture
        .WriteProperty nTransColor, pTransColor
        .WriteProperty nBackColor, pBackColor
        .WriteProperty nStyle, pStyle
        .WriteProperty nSyncTransColor, pSyncTransColor
        .WriteProperty nBorderStyle, BorderStyle, BORDER.[Fixed Single]
        .WriteProperty nMouseIcon, MouseIcon
        .WriteProperty nMousePointer, MousePointer
        .WriteProperty nEnabled, Enabled
        .WriteProperty nWidth, Width
        .WriteProperty nHeight, Height
    End With
End Sub

Public Property Let Style(ByVal new_style As STYLES)
Attribute Style.VB_Description = "Sets/returns the painting format"
    ' Set the style property
    Call validateStyle(new_style)
    If Err.Number <> 0 Then
        new_style = m_def_Style
        Err.Number = 0
    End If
    pStyle = new_style
    drawImage pPicture
    PropertyChanged "Style"
End Property

Public Property Get Style() As STYLES
    ' Get the style property
    Style = pStyle
End Property

Public Property Let TransparentColor(ByVal new_transcolor As OLE_COLOR)
Attribute TransparentColor.VB_Description = "Sets/returns the color that is hidden from view"
    ' set the transparent color value
    pTransColor = new_transcolor
    MaskColor = pTransColor     ' set the usercontrol's maskcolor value
    PropertyChanged "TransparentColor"
End Property

Public Property Get TransparentColor() As OLE_COLOR
    ' get the transparent color property
    TransparentColor = pTransColor
End Property

Public Property Let BackColor(ByVal new_backcolor As OLE_COLOR)
Attribute BackColor.VB_Description = "Sets/returns the background color for the control"
    ' set the back color value
    pBackColor = new_backcolor
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    ' get the back color property
    BackColor = pBackColor
End Property

Public Property Get SyncTransparentColor() As Boolean
Attribute SyncTransparentColor.VB_Description = "If true, forces the TransparentColor property to be the color of the picture at point 0,0"
    SyncTransparentColor = pSyncTransColor
End Property

Public Property Let SyncTransparentColor(ByVal new_sync As Boolean)
    If Ambient.UserMode Then
'        Err.Raise Number:=31013, _
'        Description:="Property is read-only at run time."
    Else
        Call validateSync(new_sync)
        If Err.Number <> 0 Then
            new_sync = m_def_SyncTransColor
            Err.Number = 0
        End If
    End If
    pSyncTransColor = new_sync
    PropertyChanged "SyncTransparentColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BORDER
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BORDER)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ' Validation is supplied by UserControl.
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Width
Public Property Get Width() As Long
    Width = UserControl.Width
End Property

Public Property Let Width(ByVal New_Width As Long)
    UserControl.Width() = New_Width
    PropertyChanged "Width"
End Property

Private Sub validateStyle(ByVal new_style As STYLES)
    ' Check the Styles property
    Select Case new_style
        Case Is < STYLES.[None], Is > STYLES.[Stretch]
            Err.Raise 380
    End Select
End Sub

Private Sub validateSync(ByVal new_sync As Boolean)
    ' Check the SyncTransColor property
    If new_sync <> True And new_sync <> False Then
        Err.Raise 380
    End If
End Sub

Private Sub paintImage()
    Dim X As Long, Y As Long
    Dim wSrc As Long, hSrc As Long
    Dim wDest As Long, hDest As Long
    
    wSrc = imgSrc.Width
    hSrc = imgSrc.Height
    
    With picDest
        Select Case pStyle
            Case STYLES.[None]
                Set .Picture = imgSrc.Picture
                .Width = wSrc
                .Height = hSrc
            Case STYLES.[Repeat]
                .Width = UserControl.Width
                .Height = UserControl.Height
                wDest = .Width
                hDest = .Height
                
                Y = 0
                Do While Y <= hDest
                    X = 0
                    Do While X <= wDest
                        .PaintPicture imgSrc.Picture, X, Y, wSrc, hSrc
                        X = X + wSrc
                    Loop
                    Y = Y + hSrc
                Loop
            Case STYLES.[Stretch]
                .Width = UserControl.Width
                .Height = UserControl.Height
                wDest = .Width
                hDest = .Height
                
                .PaintPicture imgSrc.Picture, 0, 0, wDest, hDest
        End Select
        .Refresh
    End With
End Sub

Private Sub drawImage(pic As Picture, Optional snapSize As Boolean = True)
    ' 1. If the pic is valid, store it in the imgSrc control
    ' 2. Depending on the pStyle setting,
    '       a. Paint (single, repeat or stretch)
    '          the imgSrc.picture into the picDest control
    ' 3. If the pSyncTransColor is set:
    '       a. Change the pTranscolor value to the hotpoint color
    '          (point 0,0 of the picDest)
    '       b. Change the usercontrol's maskcolor to the
    '          pTranscolor value
    ' 4. Change the usercontrol's maskpicture property to the
    '    picDest picture value
    
    Dim clr As Long
    
    Set imgSrc.Picture = pic
    If Not (pic Is Nothing) Then
        Call paintImage
        
        If pSyncTransColor Then
            ' get the hotpoint color and set pTranscolor to it
            clr = picDest.Point(0, 0)
            TransparentColor = clr
        End If
        
        If snapSize Then
            Width = picDest.Width
            Height = picDest.Height
        End If
        
        Set MaskPicture = picDest.Image
        Set Picture = picDest.Image
        
        Set pPicture = pic
    Else
        TransparentColor = m_def_TransColor
        Set MaskPicture = Nothing
        Set Picture = Nothing
        Set pPicture = Nothing
    End If
    Refresh
End Sub
