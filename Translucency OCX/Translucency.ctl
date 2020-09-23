VERSION 5.00
Begin VB.UserControl Translucency 
   BackColor       =   &H80000007&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   495
   ToolboxBitmap   =   "Translucency.ctx":0000
   Begin VB.PictureBox picCapture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   870
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Translucency.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Translucency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'credits
'Unknown Author for the idea of BitBlt-ing the portion of the screen to the form
'
'known errors...
'=======================================================================================================
'(:-0   (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0
'=======================================================================================================

'when we edit the parent form's properties and try running the application, the form is not translucent
'the forms' translucency is not perfect when the form resizes or moves or changes state

'=======================================================================================================
'(:-0   (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0    (:-0
'=======================================================================================================

'=======================================================================================================
'some enhancements u can try
'1) to make the translucency redraw as the parent form is dragged
'        subclass the form to know when the parent form is dragged, and use translucency1.drawTranslucency
'2) to make the translucency work on resize
'       write translucency1.drawTranslucency in the resize event of parent form
'=======================================================================================================
'_______________________________________________________________________________________________________
'guys, pls report the bugs...
'i guess i need feedback to do more work...
'if the source gave u some spark of ideas or taught you new things,
'i consider the work was not futile...

'also if u think this ain't that bad for a beginner, pls vote for me...
'
'thanks and luv,
'Praveen
'
'_______________________________________________________________________________________________________
'Praveenc_1999@ yahoo.com
'=======================================================================================================

Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal dreamAKA As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateDC Lib "Gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32.dll" (ByVal hdc As Long) As Long
'Default Property Values:

Const m_def_BlendColor = vbBlack
'Property Variables:

Dim m_BlendColor As OLE_COLOR
Private Const SM_CYCAPTION = 4
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33

Public Function drawTranslucency()
'UserControl.Parent.Cls
'UserControl.Parent.Refresh
    If Ambient.UserMode Then
        BlendIT UserControl.Parent, picCapture
    End If
End Function

Private Sub BlendIT(frm As Object, Pic As PictureBox)

  Dim titleBarheight As Integer
  Dim xDeviation As Integer
  Dim yDeviation As Integer
  Dim windowFrameHeight As Integer
  Dim windowframewidth As Integer
  Dim BlendVal As Long
  Dim hDCscr As Long
    
    'we need to set the size of the picture
    picCapture.Move 5000, 5000, frm.Width, frm.Height

    'crucial
    frm.AutoRedraw = True

    'now the calculations that make the code work with diff. border styles

    If frm.BorderStyle <> 0 Then
        titleBarheight = GetSystemMetrics(SM_CYCAPTION)
        windowFrameHeight = GetSystemMetrics(SM_CYFRAME)
        windowframewidth = GetSystemMetrics(SM_CXFRAME)
        yDeviation = titleBarheight + windowFrameHeight
        xDeviation = windowframewidth
      Else
        yDeviation = 0
        xDeviation = 0
    End If

    'set the back color of the form to the blended color..
    'this gives the effect of the pale shade color

    frm.BackColor = m_BlendColor
    'ok now getting the picture
  
    'first clean the board before u draw
    Pic.Cls
    
    'get the DeviceContext for the screen
    hDCscr = CreateDC("DISPLAY", "", "", 0)
    
    'get the picture of the screen at forms' position, width and height, and draw it to the form
    BitBlt Pic.hdc, 0, 0, frm.Width, frm.Height, hDCscr, frm.Left / Screen.TwipsPerPixelX + xDeviation, frm.Top / Screen.TwipsPerPixelX + yDeviation, vbSrcCopy 'vbSrcErase
    frm.Cls
    
    'take this for granted.. don't ask what is this
    'this idea came from the AlphaBlending-The Mystical way by Alexander Anikin
    
    BlendVal = 11796480
    
    'Blend the picture on the picturebox to the form
    AlphaBlend frm.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, Pic.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, BlendVal
    
    DeleteDC hDCscr

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BlendColor = PropBag.ReadProperty("BlendColor", m_def_BlendColor)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BlendColor() As OLE_COLOR

    BlendColor = m_BlendColor

End Property

Public Property Let BlendColor(ByVal New_BlendColor As OLE_COLOR)

    m_BlendColor = New_BlendColor
    PropertyChanged "BlendColor"

End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_BlendColor = m_def_BlendColor

End Sub

Private Sub UserControl_Resize()

    Image1.Move 0, 0
    Width = Image1.Width
    Height = Image1.Height

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BlendColor", m_BlendColor, m_def_BlendColor)
    
End Sub
