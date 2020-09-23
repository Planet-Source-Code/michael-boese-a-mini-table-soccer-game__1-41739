VERSION 5.00
Begin VB.UserControl SkinControl 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   LockControls    =   -1  'True
   ScaleHeight     =   1365
   ScaleWidth      =   2265
   ToolboxBitmap   =   "ctlSkinControl.ctx":0000
   Begin VB.PictureBox picSkin 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "SkinControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************
'
' SkinControl UserControl
'
' Transparente Steuerelemente auf der Basis einer Bitmap erstellen
' (Methode 1).
'
' © Copyright 2000 by Herfried K. Wagner für ActiveVB.de.
'
' Teile des Programms wurden von Dave Scarmozzino (www.TheScarms.com)
' entwickelt.
'
' Der  Autor übernimmt keine Haftung für Schäden, die durch dieses
' Programm verursacht wurden.
' Sie sind nicht berechtigt, diesen Code weiterzugeben, ausser in Form
' einer kompilierten Anwendung!
'
' Wenn Sie in diesem Beispiel Fehler entdecken, dann teilen Sie mir
' das bitte mit.
'
' e-Mail:  Hirf@ActiveVB.de     Homepage:  http://www.ActiveVB.de
'                                          http://www.beam.to/HirfHome
'
'**********************************************************************
Option Explicit

'
' Win32 API-Declarations.
'
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" _
    (ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" _
    (ByVal hDestRgn As Long, _
    ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, _
    ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Long) As Long

'
' Win32 API-Constants.
'
Private Const RGN_OR = 2

'
' Variablen.
'
' Lokale Kopien der Eigenschaften.
Private m_picPicture As Picture
Private m_olcMaskColor As OLE_COLOR
Private m_blnChangeMask As Boolean
Private m_blnEnabled As Boolean
Private m_mpcMousePointer As MousePointerConstants

' Interne Variablen.
Private m_lngHeight As Long
Private m_lngWidth As Long

' Ereignisse.
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
Public Event EnabledChange(blnEnabled As Boolean)

'
' Ereignisse.
'
Private Sub UserControl_Click()
    If m_blnEnabled = True Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If m_blnEnabled = True Then RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
    Shift As Integer)
    
    If m_blnEnabled = True Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If m_blnEnabled = True Then RaiseEvent _
        KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, _
    Shift As Integer)
    
    If m_blnEnabled = True Then RaiseEvent _
        KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    
    If m_blnEnabled = True Then RaiseEvent _
        MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    
    If m_blnEnabled = True Then RaiseEvent _
        MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    
    If m_blnEnabled = True Then _
        RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'
' UserControl-Funktionen.
'
Private Sub UserControl_Initialize()
    picSkin.AutoRedraw = True
    picSkin.AutoSize = True
End Sub

Private Sub UserControl_InitProperties()
    
    ' Nothing als Voreinstellung verwenden, um
    ' die Picture-Eigenschaft zu initialisieren, zu lesen
    ' und zu schreiben, damit keine .frx-Datei benötigt
    ' wird, wenn kein Bild verwendet wird.
    Set UserControl.Picture = Nothing
    
    UserControl.BackColor = vbWhite
    lblMessage.Caption = "(Select Picture)"
    lblMessage.Visible = True
    
    m_blnEnabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_picPicture = PropBag.ReadProperty("Picture", Nothing)
    m_blnEnabled = PropBag.ReadProperty("Enabled", True)
    m_olcMaskColor = PropBag.ReadProperty("MaskColor", vbBlack)
    m_blnChangeMask = PropBag.ReadProperty("ChangeMask", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_mpcMousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    
    UserControl.MousePointer = m_mpcMousePointer
    Call CreateSkin
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", m_picPicture, Nothing)
    Call PropBag.WriteProperty("Enabled", m_blnEnabled, True)
    Call PropBag.WriteProperty("MaskColor", m_olcMaskColor, vbBlack)
    Call PropBag.WriteProperty("ChangeMask", m_blnChangeMask, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_mpcMousePointer, vbDefault)
End Sub

'
' Eigenschaften.
'
Public Property Get hWnd() As Long   ' Read-Only.
    hWnd = UserControl.hWnd
End Property

Public Property Get hDC() As Long   ' Read-Only.
    hDC = UserControl.hDC
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Gibt die Farbe zurück, die als transparente Farbe in der Bitmep verwendet werden soll oder legt diese fest."
    MaskColor = m_olcMaskColor
End Property

Public Property Let MaskColor(ByVal olcNewMaskColor As OLE_COLOR)
    m_olcMaskColor = olcNewMaskColor
    Call CreateSkin
    
    PropertyChanged "MaskColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt zurück oder legt fest, ob das Steuerelement auf Ereignisse reagieren soll oder nicht."
    Enabled = m_blnEnabled
End Property

Public Property Let Enabled(ByVal blnNewEnabled As Boolean)
    m_blnEnabled = blnNewEnabled
    RaiseEvent EnabledChange(blnNewEnabled)
    
    PropertyChanged "Enabled"
End Property

Public Property Get ChangeMask() As Boolean
Attribute ChangeMask.VB_Description = "Gibt einen Wert zurück oder legt einen Wert fest, der angibt, ob beim Ändern der Picture-Eigenschaft die Form des Steuerelements neu festgelegt werden soll."
    ChangeMask = m_blnChangeMask
End Property

Public Property Let ChangeMask(ByVal blnNewChangeMask As Boolean)
    m_blnChangeMask = blnNewChangeMask
    Call CreateSkin
    
    PropertyChanged "ChangeMask"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Gibt die Bitmap zurück, die als Grafik für das Steuerelement verwendet werden soll oder legt diese fest."
    Set Picture = m_picPicture
End Property

Public Property Let Picture(ByVal picNewPicture As Picture)
    Set m_picPicture = picNewPicture
    Call CreateSkin
    
    PropertyChanged "Picture"
End Property

Public Property Set Picture(ByVal picNewPicture As Picture)
    Set m_picPicture = picNewPicture
    Call CreateSkin
    
    PropertyChanged "Picture"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = m_mpcMousePointer
End Property

Public Property Let MousePointer(mpcNewMousePointer As MousePointerConstants)
    m_mpcMousePointer = mpcNewMousePointer
    UserControl.MousePointer = m_mpcMousePointer
    
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MouseIcon(ByVal picNewPicture As Picture)
    Set UserControl.MouseIcon = picNewPicture
    
    PropertyChanged "MouseIcon"
End Property

Public Property Set MouseIcon(ByVal picNewPicture As Picture)
    Set UserControl.MouseIcon = picNewPicture
    
    PropertyChanged "MouseIcon"
End Property

'
' Interne Routinen.
'
'
' The optional last parameter allows you to specify the
' image's background color. If left blank, the
' color of the image's top left pixel is used.
'
Private Function RegionFromBitmap(picSource As PictureBox) As Long
    Dim lngReturn As Long
    Dim lngRgnTmp As Long
    Dim lngSkinRgn As Long
    Dim lngStart As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngBackColor As Long
    '
    ' Create a rectangular region.
    ' A region is a rectangle, polygon, or ellipse (or a
    ' combination of two or more of these shapes)
    ' that can be filled, painted, inverted, framed, and
    ' used to perform hit testing (testing for
    ' the cursor location).
    '
    lngSkinRgn = CreateRectRgn(0, 0, 0, 0)
    
    With picSource
        '
        ' Get the dimensions of the bitmap.
        '
        m_lngHeight = .Height / Screen.TwipsPerPixelY
        m_lngWidth = .Width / Screen.TwipsPerPixelX
        '
        ' If no background color is passed in, get the red, green,
        ' blue (RGB) color value of the top
        ' left pixel in the picturebox's device context (DC).
        '
        lngBackColor = m_olcMaskColor
        '
        ' Loop through the bitmap, row by row, examining
        ' each pixel. In each row, work from left to
        ' right comparing each pixel to the background color.
        '
        For lngRow = 0 To m_lngHeight - 1
            lngCol = 0
            Do While lngCol < m_lngWidth
                '
                ' Skip all pixels in a row with the same color
                ' as the background color.
                '
                Do While lngCol < m_lngWidth And GetPixel(.hDC, lngCol, _
                    lngRow) = lngBackColor
                    
                    lngCol = lngCol + 1
                Loop
                
                If lngCol < m_lngWidth Then
                    '
                    ' Get the start and end of the block of
                    ' pixels in the row that are not the same
                    ' color as the background.
                    '
                    lngStart = lngCol
                    Do While lngCol < m_lngWidth And GetPixel(.hDC, _
                        lngCol, lngRow) <> lngBackColor
                        
                        lngCol = lngCol + 1
                    Loop
                    If lngCol > m_lngWidth Then lngCol = m_lngWidth
                    '
                    ' Create a region equal in size to the line
                    ' of pixels that don't match the
                    ' background color. Combine this region
                    ' with our final region.
                    '
                    lngRgnTmp = CreateRectRgn(lngStart, lngRow, _
                        lngCol, lngRow + 1)
                    lngReturn = CombineRgn(lngSkinRgn, _
                        lngSkinRgn, lngRgnTmp, RGN_OR)
                    Call DeleteObject(lngRgnTmp)
                End If
            Loop
        Next lngRow
    End With
    RegionFromBitmap = lngSkinRgn
End Function

Private Sub CreateSkin()
    Dim lngRegion As Long
    
    ' Wenn keine Grafik in der Picture-Eigenschaft
    ' estgelegt wurde.
    If m_picPicture Is Nothing Then
        Dim lngSkinRgn As Long
        
        Set UserControl.Picture = Nothing
        
        If Ambient.UserMode = True Then   ' In der EXE.
            UserControl.BackColor = vbButtonFace
            If lblMessage.Visible = True Then _
                lblMessage.Visible = False
            
            ' Wenn in Runtime-Modus keine Grafik,
            ' Steuerelement "löschen".
            lngSkinRgn = CreateRectRgn(0, 0, 0, 0)
        Else   ' In IDE.
            UserControl.BackColor = vbWhite
            lblMessage.Visible = True
            lblMessage.Caption = "(Select Picture)"
            lngSkinRgn = CreateRectRgn(0, 0, UserControl.Width, _
                UserControl.Height)
        End If
        
        Call SetWindowRgn(UserControl.hWnd, lngSkinRgn, True)
        
        Exit Sub
    End If
    
    ' Wenn nicht in IDE und ChangeMask = False
    ' und eine Grafik vorhanden, dann nicht die Grösse des
    ' Steuerelements ändern sondern nur die Grafik.
    If Ambient.UserMode = True And m_blnChangeMask = False And _
        UserControl.Picture <> 0 Then _
        UserControl.Picture = m_picPicture: Exit Sub
    
    ' Zurücksetzen, wenn die Einstellungen wegen Fehlen
    ' der Grafik in IDE geändert wurden.
    UserControl.BackColor = vbButtonFace
    lblMessage.Visible = False
    
    With UserControl
        .picSkin.Picture = m_picPicture
        '
        ' Size the PictureBox.
        '
        .Width = .picSkin.Width
        .Height = .picSkin.Height
        '
        ' Based on the picture, create a region
        ' for Windows to use for our PictureBox and tell
        ' Windows not to paint anything outside this region.
        '
        lngRegion = RegionFromBitmap(UserControl.picSkin)
        Call SetWindowRgn(.hWnd, lngRegion, True)
        .picSkin.Picture = LoadPicture("")
        '
        ' Load the skin image into the form's background.
        '
        .Picture = m_picPicture
    End With
End Sub
