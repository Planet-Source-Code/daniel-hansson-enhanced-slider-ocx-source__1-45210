VERSION 5.00
Begin VB.UserControl Slider 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ScaleHeight     =   3630
   ScaleWidth      =   3405
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   2040
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   9
      Top             =   2400
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicEmptyFill2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   2040
      Picture         =   "UserControl1.ctx":02D6
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   8
      Top             =   1200
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2160
      Picture         =   "UserControl1.ctx":0354
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   7
      Top             =   840
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicFullFill2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   2160
      Picture         =   "UserControl1.ctx":06DE
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   6
      Top             =   2160
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicHandle2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Picture         =   "UserControl1.ctx":075C
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   150
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicHandle 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   720
      Picture         =   "UserControl1.ctx":09FE
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   4
      Top             =   360
      Width           =   285
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicFullFill 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   360
      Picture         =   "UserControl1.ctx":0C98
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   3
      Top             =   240
      Width           =   15
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1440
      Picture         =   "UserControl1.ctx":0D2A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      Top             =   240
      Width           =   210
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicEmptyFill 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1080
      Picture         =   "UserControl1.ctx":10DC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   1
      Top             =   240
      Width           =   15
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "UserControl1.ctx":116E
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   0
      Top             =   240
      Width           =   165
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'declare BitBlt to copy parts of pictures to other pictures
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'same as BitBlt onlt it will stretch too
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'list of options
Public Enum eHorVert
    Horizontal
    vertical
End Enum

'events
Public Event DblClick(Value As Long)
Public Event Scroll()
Public Event Change()
Public Event Click()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'values
Dim CurX As Single
Dim CurY As Single

Dim iValue As Long
Dim iMax As Long
Dim iMin As Long

Dim TempVal As Single

Dim bEnabled As Boolean
Dim bSnap As Boolean
Dim UseVal As Boolean

Dim sHorVert As eHorVert

'Colors
Dim lBgColor As Long
Dim lFgColor As Long
Dim lFgColor2 As Long
Dim FcRed As Single
Dim FcGreen As Single
Dim FcBlue As Single
Dim FcRed2 As Single
Dim FcGreen2 As Single
Dim FcBlue2 As Single

Public Property Let Style(eStyle As eHorVert)
    sHorVert = eStyle
    If sHorVert = Horizontal Then
        UserControl.Width = UserControl.Height
    Else
        UserControl.Height = UserControl.Width
    End If
    UserControl_Resize
    UseVal = False
    Draw
End Property

Public Property Get Style() As eHorVert
    Style = sHorVert
End Property

Public Property Let Enabled(BolEnabled As Boolean)
    bEnabled = BolEnabled
End Property

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Public Property Let Snap(BolSnap As Boolean)
    bSnap = BolSnap
End Property

Public Property Get Snap() As Boolean
    Snap = bSnap
End Property

Public Property Let BackColor(lColor As OLE_COLOR)
    lBgColor = lColor
    SetBgColor
    UseVal = False
    Draw
    RaiseEvent Change
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = lBgColor
End Property

Public Property Let ForeColor(lColor As OLE_COLOR)
    lFgColor = lColor
    FcRed = GetRed(lColor)
    FcGreen = GetGreen(lColor)
    FcBlue = GetBlue(lColor)
    UseVal = False
    Draw
    RaiseEvent Change
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lFgColor
End Property

Public Property Let ForeColor2(lColor As OLE_COLOR)
    lFgColor2 = lColor
    FcRed2 = GetRed(lColor)
    FcGreen2 = GetGreen(lColor)
    FcBlue2 = GetBlue(lColor)
    UseVal = False
    Draw
    RaiseEvent Change
End Property

Public Property Get ForeColor2() As OLE_COLOR
    ForeColor2 = lFgColor2
End Property

Public Property Let Value(intValue As Long) 'when changing value, this will run
    iValue = intValue
    If iValue < iMin Then iValue = iMin
    If iValue > iMax Then iValue = iMax
    UseVal = False
    Draw
    RaiseEvent Change
End Property

Public Property Get Value() As Long 'when gettin value, this will run
    Value = iValue
End Property

Public Property Let Max(intMax As Long)
    iMax = intMax
    If iMax <= iMin Then iMax = iMin + 1
    If iValue > iMax Then iValue = iMax
    UseVal = False
    Draw
    RaiseEvent Change
End Property

Public Property Get Max() As Long
    Max = iMax
End Property

Public Property Let Min(intMin As Long)
    iMin = intMin
    If iMin >= iMax Then iMin = iMax - 1
    If iValue < iMin Then iValue = iMin
    UseVal = False
    Draw
    RaiseEvent Change
End Property

Public Property Get Min() As Long
    Min = iMin
End Property

Private Sub UserControl_Click()
    RaiseEvent Click 'when clicking, raise the clicking event so it can be used in the form
End Sub

Private Sub UserControl_DblClick()
    If bEnabled = True Then
        If sHorVert = Horizontal Then
            'calc value
            iValue = Round2(((CurX - PicHandle.Width) / Screen.TwipsPerPixelX) / ((UserControl.ScaleWidth / Screen.TwipsPerPixelX) - PicLeft.ScaleWidth - PicRight.ScaleWidth - PicHandle.ScaleWidth) * (iMax - iMin))
            iValue = iValue + iMin
            If CurX < (PicHandle.ScaleWidth / 2 + PicLeft.ScaleWidth - 1) * Screen.TwipsPerPixelX Then CurX = (PicHandle.ScaleWidth / 2 + PicLeft.ScaleWidth - 1) * Screen.TwipsPerPixelX
            If CurX > UserControl.ScaleWidth - (PicHandle.ScaleWidth / 2 + PicRight.ScaleWidth) * Screen.TwipsPerPixelX Then CurX = UserControl.ScaleWidth - (PicHandle.ScaleWidth / 2 + PicRight.ScaleWidth) * Screen.TwipsPerPixelX
            TempVal = CurX / Screen.TwipsPerPixelX - PicHandle.ScaleWidth / 2
        Else
            'calc value
            iValue = Round2(((CurY - PicHandle2.Height) / Screen.TwipsPerPixelY) / ((UserControl.ScaleHeight / Screen.TwipsPerPixelY) - PicDown.ScaleHeight - PicUp.ScaleHeight - PicHandle2.ScaleHeight) * (iMax - iMin))
            iValue = iValue + iMin
            iValue = (iMax + iMin) - iValue '- (iMax - iMin)
            If CurY < (PicHandle2.ScaleHeight / 2 + PicDown.ScaleHeight - 1) * Screen.TwipsPerPixelY Then CurY = (PicHandle2.ScaleHeight / 2 + PicDown.ScaleHeight - 1) * Screen.TwipsPerPixelY
            If CurY > UserControl.ScaleHeight - (PicHandle2.ScaleHeight / 2 + PicUp.ScaleHeight) * Screen.TwipsPerPixelY Then CurY = UserControl.ScaleHeight - (PicHandle2.ScaleHeight / 2 + PicUp.ScaleHeight) * Screen.TwipsPerPixelY
            TempVal = UserControl.ScaleHeight / Screen.TwipsPerPixelY - (CurY / Screen.TwipsPerPixelY - PicHandle2.ScaleHeight / 2) - 4
        End If
        UseVal = True
        'check that the value isnt too big/small
        If iValue < iMin Then iValue = iMin
        If iValue > iMax Then iValue = iMax
        Draw
        RaiseEvent Scroll
        RaiseEvent Change
    End If
    
    RaiseEvent DblClick(iValue)
End Sub

Private Sub UserControl_Initialize()
    'set standard values
    iValue = 50
    iMax = 100
    iMin = -100
    lBgColor = vbBlack
    lFgColor = RGB(200, 200, 0)
    lFgColor2 = RGB(200, 0, 0)
    bEnabled = True
    UseVal = False
    bSnap = False
    sHorVert = Horizontal
    Draw
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        iValue = iValue - 1
    ElseIf KeyCode = vbKeyRight Then
        iValue = iValue + 1
    End If
        
    If iValue < iMin Then iValue = iMin
    If iValue > iMax Then iValue = iMax
    UseVal = False
    Draw
    RaiseEvent Change
    RaiseEvent Scroll
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And bEnabled = True Then
        If sHorVert = Horizontal Then
            'calc value
            iValue = Round2(((X - PicHandle.Width) / Screen.TwipsPerPixelX) / ((UserControl.ScaleWidth / Screen.TwipsPerPixelX) - PicLeft.ScaleWidth - PicRight.ScaleWidth - PicHandle.ScaleWidth) * (iMax - iMin))
            iValue = iValue + iMin
            If X < (PicHandle.ScaleWidth / 2 + PicLeft.ScaleWidth - 1) * Screen.TwipsPerPixelX Then X = (PicHandle.ScaleWidth / 2 + PicLeft.ScaleWidth - 1) * Screen.TwipsPerPixelX
            If X > UserControl.ScaleWidth - (PicHandle.ScaleWidth / 2 + PicRight.ScaleWidth) * Screen.TwipsPerPixelX Then X = UserControl.ScaleWidth - (PicHandle.ScaleWidth / 2 + PicRight.ScaleWidth) * Screen.TwipsPerPixelX
            TempVal = X / Screen.TwipsPerPixelX - PicHandle.ScaleWidth / 2
        Else
            'calc value
            iValue = Round2(((Y - PicHandle2.Height) / Screen.TwipsPerPixelY) / ((UserControl.ScaleHeight / Screen.TwipsPerPixelY) - PicDown.ScaleHeight - PicUp.ScaleHeight - PicHandle2.ScaleHeight) * (iMax - iMin))
            iValue = iValue + iMin
            iValue = (iMax + iMin) - iValue '- (iMax - iMin)
            If Y < (PicHandle2.ScaleHeight / 2 + PicDown.ScaleHeight - 1) * Screen.TwipsPerPixelY Then Y = (PicHandle2.ScaleHeight / 2 + PicDown.ScaleHeight - 1) * Screen.TwipsPerPixelY
            If Y > UserControl.ScaleHeight - (PicHandle2.ScaleHeight / 2 + PicUp.ScaleHeight) * Screen.TwipsPerPixelY Then Y = UserControl.ScaleHeight - (PicHandle2.ScaleHeight / 2 + PicUp.ScaleHeight) * Screen.TwipsPerPixelY
            TempVal = UserControl.ScaleHeight / Screen.TwipsPerPixelY - (Y / Screen.TwipsPerPixelY - PicHandle2.ScaleHeight / 2) - 4
        End If
        UseVal = True
        'check that the value isnt too big/small
        If iValue < iMin Then iValue = iMin
        If iValue > iMax Then iValue = iMax
        Draw
        RaiseEvent Scroll
        RaiseEvent Change
    End If
    CurX = X
    CurY = Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag) 'when reading properties
    Value = PropBag.ReadProperty("Value", 50) 'show the value in properties window
    Max = PropBag.ReadProperty("Max", 100) 'show the max value in properties window
    Min = PropBag.ReadProperty("Min", -100) 'show the min value in properties window
    BackColor = PropBag.ReadProperty("BackColor", vbBlack)
    ForeColor = PropBag.ReadProperty("ForeColor", RGB(200, 200, 0))
    ForeColor2 = PropBag.ReadProperty("ForeColor2", RGB(200, 0, 0))
    Enabled = PropBag.ReadProperty("Enabled", True)
    Snap = PropBag.ReadProperty("Snap", False)
    Style = PropBag.ReadProperty("Style", Horizontal)
    UseVal = False
    Draw
End Sub

Private Sub UserControl_Resize()
    'keep slider within good dimensions
    If sHorVert = Horizontal Then
        If UserControl.ScaleWidth / Screen.TwipsPerPixelX < PicLeft.ScaleWidth + PicRight.ScaleWidth + PicHandle.ScaleWidth * 2 Then
            UserControl.Width = (PicLeft.ScaleWidth + PicRight.ScaleWidth + PicHandle.ScaleWidth * 2) * Screen.TwipsPerPixelX
        End If
        UserControl.Height = PicLeft.Height
    Else
        If UserControl.ScaleHeight / Screen.TwipsPerPixelY < PicDown.ScaleHeight + PicUp.ScaleHeight + PicHandle2.ScaleHeight * 2 Then
            UserControl.Height = (PicDown.ScaleHeight + PicUp.ScaleHeight + PicHandle2.ScaleHeight * 2) * Screen.TwipsPerPixelY
        End If
        UserControl.Width = PicDown.Width
    End If
    UseVal = False
    Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag) 'when writing properties
    PropBag.WriteProperty "Value", Value, 50 'store new value when changed
    PropBag.WriteProperty "Max", Max, 100 'store new max value when changed
    PropBag.WriteProperty "Min", Min, -100 'store new min value when changed
    PropBag.WriteProperty "BackColor", BackColor, vbBlack
    PropBag.WriteProperty "ForeColor", ForeColor, RGB(200, 200, 0)
    PropBag.WriteProperty "ForeColor2", ForeColor2, RGB(200, 0, 0)
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Snap", Snap, False
    PropBag.WriteProperty "Style", Style, Horizontal
End Sub

Private Sub Draw()
    If sHorVert = Horizontal Then
        If iMax - iMin <> 0 Then
            Dim TempX As Long
            'check the xvalue to draw the handle
            If UseVal = False Or Snap = True Then
                TempX = (iValue / (iMax - iMin) * (UserControl.ScaleWidth - PicLeft.Width - PicRight.Width - PicHandle.Width)) / Screen.TwipsPerPixelX
                TempX = TempX + PicHandle.ScaleWidth / 2 + 1
                TempX = TempX - iMin / (iMax - iMin) * (UserControl.ScaleWidth - PicLeft.Width - PicRight.Width - PicHandle.Width) / Screen.TwipsPerPixelX
            ElseIf UseVal = True And Snap = False Then
                TempX = TempVal
            End If
            'clear old slider
            UserControl.Cls
            
            SetFgColor
            
            'stretch the 'filled' pic from 0 to tempx (where the handle is)
            StretchBlt UserControl.hdc, 0, 0, TempX, PicLeft.ScaleHeight, PicFullFill.hdc, 0, 0, PicEmptyFill.ScaleWidth, PicEmptyFill.ScaleHeight, vbSrcCopy
            'stretch the 'empty' pic from tempx to the end
            StretchBlt UserControl.hdc, TempX, 0, UserControl.Width / Screen.TwipsPerPixelX - TempX, PicLeft.ScaleHeight, PicEmptyFill.hdc, 0, 0, PicEmptyFill.ScaleWidth, PicEmptyFill.ScaleHeight, vbSrcCopy
            
            'draw the left side
            BitBlt UserControl.hdc, 0, 0, PicLeft.ScaleWidth, PicLeft.ScaleHeight, PicLeft.hdc, 0, 0, vbSrcCopy
            'draw the right side
            BitBlt UserControl.hdc, UserControl.Width / Screen.TwipsPerPixelX - PicRight.ScaleWidth, 0, PicRight.ScaleWidth, PicLeft.ScaleHeight, PicRight.hdc, 0, 0, vbSrcCopy
            'draw the handle
            BitBlt UserControl.hdc, TempX, 4, PicHandle.ScaleWidth, PicHandle.ScaleHeight, PicHandle.hdc, 0, 0, vbSrcCopy
            'refresh
            UserControl.Refresh
        End If
    Else
        If iMax - iMin <> 0 Then
            Dim TempY As Long
            'check the yvalue to draw the handle
            If UseVal = False Or Snap = True Then
                TempY = (iValue / (iMax - iMin) * (UserControl.ScaleHeight - PicDown.Height - PicUp.Height - PicHandle2.Height)) / Screen.TwipsPerPixelY
                TempY = TempY + PicHandle2.ScaleHeight / 2 + 1 + PicHandle2.ScaleHeight
                TempY = TempY - iMin / (iMax - iMin) * (UserControl.ScaleHeight - PicDown.Height - PicUp.Height - PicHandle2.Height) / Screen.TwipsPerPixelY
            ElseIf UseVal = True And Snap = False Then
                TempY = TempVal
            End If
            'clear old slider
            UserControl.Cls
            
            SetFgColor
            
            'stretch the 'empty' pic from 0 to TempY (where the handle is)
            StretchBlt UserControl.hdc, 0, 0, PicDown.ScaleWidth, UserControl.Height / Screen.TwipsPerPixelX - TempY, PicEmptyFill2.hdc, 0, 0, PicEmptyFill2.ScaleWidth, PicEmptyFill2.ScaleHeight, vbSrcCopy
            'stretch the 'filled' pic from TempY to the end
            StretchBlt UserControl.hdc, 0, UserControl.Height / Screen.TwipsPerPixelX - TempY, PicUp.ScaleWidth, UserControl.Height / Screen.TwipsPerPixelX, PicFullFill2.hdc, 0, 0, PicFullFill2.ScaleWidth, PicFullFill2.ScaleHeight, vbSrcCopy
            
            'draw the left side
            BitBlt UserControl.hdc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hdc, 0, 0, vbSrcCopy
            'draw the right side
            BitBlt UserControl.hdc, 0, UserControl.Height / Screen.TwipsPerPixelY - PicDown.ScaleHeight, PicDown.ScaleWidth, PicDown.ScaleHeight, PicDown.hdc, 0, 0, vbSrcCopy
            'draw the handle
            BitBlt UserControl.hdc, 4, UserControl.Height / Screen.TwipsPerPixelX - TempY, PicHandle2.ScaleWidth, PicHandle2.ScaleHeight, PicHandle2.hdc, 0, 0, vbSrcCopy
            'refresh
            UserControl.Refresh
        End If
    End If
End Sub

Private Sub SetBgColor()
    Dim Y As Long
    
    For Y = 6 To 12
        'If sHorVert = Horizontal Then
            SetPixelV PicEmptyFill.hdc, 0, Y, lBgColor 'fill with new color
        'Else
            SetPixelV PicEmptyFill2.hdc, Y, 0, lBgColor 'fill with new color
        'End If
    Next
End Sub

Private Sub SetFgColor()
    Dim Y As Long
    Dim NewRed As Single
    Dim NewGreen As Single
    Dim NewBlue As Single
    Dim NewColor As Long
    
    If UseVal = False Or Snap = True Then
        NewRed = FcRed * ((iMax - iMin - (iValue - Min)) / (iMax - iMin)) + FcRed2 * ((iValue - iMin) / (iMax - iMin))
        NewGreen = FcGreen * ((iMax - iMin - (iValue - Min)) / (iMax - iMin)) + FcGreen2 * ((iValue - iMin) / (iMax - iMin))
        NewBlue = FcBlue * ((iMax - iMin - (iValue - Min)) / (iMax - iMin)) + FcBlue2 * ((iValue - iMin) / (iMax - iMin))
    ElseIf UseVal = True And Snap = False Then
        Dim TempSize As Single
        Dim TempSize2 As Single
        
        If sHorVert = Horizontal Then
            TempSize = TempVal / (UserControl.ScaleWidth / Screen.TwipsPerPixelX - PicLeft.ScaleWidth - PicRight.ScaleWidth)
            TempSize2 = 1 - TempSize
        Else
            TempSize = TempVal / (UserControl.ScaleHeight / Screen.TwipsPerPixelY - PicDown.ScaleHeight - PicUp.ScaleHeight)
            TempSize2 = 1 - TempSize
        End If

        NewRed = FcRed * TempSize2 + FcRed2 * TempSize
        NewGreen = FcGreen * TempSize2 + FcGreen2 * TempSize
        NewBlue = FcBlue * TempSize2 + FcBlue2 * TempSize
    End If
    
    If NewRed < 0 Then NewRed = 0
    If NewGreen < 0 Then NewGreen = 0
    If NewBlue < 0 Then NewBlue = 0
    If NewRed > 255 Then NewRed = 0
    If NewGreen > 255 Then NewGreen = 0
    If NewBlue > 255 Then NewBlue = 0
    
    NewColor = RGB(NewRed, NewGreen, NewBlue)
    
    For Y = 6 To 12
        'SetPixelV PicFullFill.hdc, 0, Y, lFgColor 'fill with new color
        If sHorVert = Horizontal Then
            SetPixelV PicFullFill.hdc, 0, Y, NewColor
        Else
            SetPixelV PicFullFill2.hdc, Y, 0, NewColor
        End If
    Next
End Sub

Private Function Round2(Number As Long) As Long 'corrected round function
    If Number - Round(Number) >= 0.5 Then Number = Round(Number) + 1
    If Number - Round(Number) < 0.5 Then Number = Round(Number)
    Round2 = Number
End Function

Private Function GetRed(Color As Long) As Integer
    GetRed = Color And 255
End Function

Private Function GetGreen(Color As Long) As Integer
    GetGreen = (Color And 65280) \ 256
End Function

Private Function GetBlue(Color As Long) As Integer
    GetBlue = (Color And 16711680) \ 65535
End Function















