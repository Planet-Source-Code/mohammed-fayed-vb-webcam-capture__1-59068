VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "ÈÑãÌÉ :  ãÍãÏ ÓãíÑ ÅÈÑÇåíã ÝÇíÏ"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   RightToLeft     =   -1  'True
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   300
      RightToLeft     =   -1  'True
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   11
      Top             =   1710
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   90
      Pattern         =   "*.bmp"
      TabIndex        =   7
      Top             =   2610
      Visible         =   0   'False
      Width           =   1425
   End
   Begin WebCam.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Options"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16576
      cFHover         =   16576
      cBhover         =   32768
      cGradient       =   32768
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   12648384
   End
   Begin VB.Timer Timer1 
      Left            =   158
      Top             =   4260
   End
   Begin WebCam.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "size"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16576
      cFHover         =   16576
      cBhover         =   32768
      cGradient       =   32768
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   12648384
   End
   Begin WebCam.lvButtons_H lvButtons_H3 
      Height          =   465
      Left            =   5490
      TabIndex        =   2
      Top             =   4980
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Caption         =   "Start"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16576
      cFHover         =   16576
      cBhover         =   32768
      cGradient       =   32768
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin WebCam.lvButtons_H lvButtons_H4 
      Height          =   465
      Left            =   3765
      TabIndex        =   3
      Top             =   4980
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Caption         =   "Stop"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16576
      cFHover         =   16576
      cBhover         =   32768
      cGradient       =   32768
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   12648384
   End
   Begin WebCam.lvButtons_H lvButtons_H5 
      Height          =   465
      Left            =   2010
      TabIndex        =   4
      Top             =   4980
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Caption         =   "Save Pic."
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16576
      cFHover         =   16576
      cBhover         =   32768
      cGradient       =   32768
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin WebCam.lvButtons_H lvButtons_H6 
      Height          =   465
      Left            =   300
      TabIndex        =   5
      Top             =   4980
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16576
      cFHover         =   16576
      cBhover         =   32768
      cGradient       =   32768
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   6900
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   60
      Width           =   225
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   420
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   30
      Width           =   6195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Each Picture saved in folder ..\myPic"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2220
      TabIndex        =   8
      Top             =   5550
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2190
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   1665
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5280
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   180
      Picture         =   "Form1.frx":9F6F
      Top             =   2160
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
'
'   collected,Converted and Edited by :
'        Mohammed Samir Fayed
'              10/2004
'
'******************************************

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

    Private m_TimeToCapture_milliseconds As Integer
    
    Private m_Width As Long
    Private m_Height As Long
    
    Private mCapHwnd As Long
   
    Private bStopped As Boolean

Private Sub Form_Load()
On Error Resume Next
Label1.BackStyle = 0 ' CheckBox
Label4.BackStyle = 0 ' Form Caption
Label5.BackStyle = 0 ' Form Close

    m_TimeToCapture_milliseconds = 100
    m_Width = 352
    m_Height = 288
    bStopped = True
    mCapHwnd = 0
    
End Sub

Public Sub Start()
    On Error Resume Next
    If mCapHwnd <> 0 Then Exit Sub
    FrameNum = 0
    
    Timer1.Interval = m_TimeToCapture_milliseconds

    ' for safety, call stop, just in case we are already running
    Me.Timer1.Enabled = False

    ' setup a capture window
    mCapHwnd = capCreateCaptureWindowA("WebCap", 0, 0, 0, m_Width, m_Height, Me.hwnd, 0)
    DoEvents
    
    ' connect to the capture device
    Call SendMessage(mCapHwnd, WM_CAP_CONNECT, 0, 0)
    DoEvents
    
    Call SendMessage(mCapHwnd, WM_CAP_SET_PREVIEW, 0, 0)

    ' set the timer information
    bStopped = False
    Me.Timer1.Enabled = True
        

End Sub
    
Public Sub StopWork()
    On Error Resume Next
    ' stop the timer
    bStopped = True
    Timer1.Enabled = False

    ' disconnect from the video source
    DoEvents

    Call SendMessage(mCapHwnd, WM_CAP_DISCONNECT, 0, 0)
    mCapHwnd = 0

End Sub


Private Sub Label1_Click()
On Error Resume Next
    Image2.Visible = Not Image2.Visible
    
    If Image2.Visible = True Then
        Image1.Width = 352
        Image1.Height = 288
        Image1.Stretch = True
    Else
        Image1.Stretch = False
    End If
    
End Sub


Private Sub Label3_Click()

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim lngReturnValue As Long
    If Button = 1 Then
        'Release capture
        Call ReleaseCapture
        'Send a 'left mouse button down on caption'-message to our form
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub Label5_Click()
    ' From Close
    Call lvButtons_H6_Click
End Sub

Private Sub lvButtons_H1_Click()
On Error Resume Next
  If mCapHwnd = 0 Then Exit Sub

    Call SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
    DoEvents
    
End Sub

Private Sub lvButtons_H2_Click()
On Error Resume Next
    
    If mCapHwnd = 0 Then Exit Sub

    Call SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0)
    DoEvents

End Sub

Private Sub lvButtons_H3_Click()
  Start
  lvButtons_H1.Enabled = True
  lvButtons_H2.Enabled = True
  lvButtons_H4.Enabled = True
  lvButtons_H3.Enabled = False
  
End Sub

Private Sub lvButtons_H4_Click()
    StopWork
    lvButtons_H1.Enabled = False
    lvButtons_H2.Enabled = False
    lvButtons_H4.Enabled = False
    lvButtons_H3.Enabled = True
End Sub

Private Sub lvButtons_H5_Click()
On Error Resume Next
DoEvents
If Dir(App.Path & "\myPic", vbDirectory) = "" Then MkDir (App.Path & "\myPic")
File1.Path = App.Path & "\myPic"
'File1.Pattern = "*.bmp"
File1.Pattern = "*.jpg"
File1.Refresh

Dim Maxnum As Integer, ii As Integer
For ii = 0 To File1.ListCount - 1
    If Left(File1.List(ii), 1) = "p" Then
        If CInt(Mid(File1.List(ii), 2, Len(File1.List(ii)) - 4)) > Maxnum Then
            Maxnum = CInt(Mid(File1.List(ii), 2, Len(File1.List(ii)) - 4))
        End If
    End If
Next

    'SavePicture Image1.Picture, App.Path & "\myPic\p" & Maxnum + 1 & ".bmp"
    
    Picture1.Picture = Image1.Picture
    SAVEJPEG App.Path & "\myPic\p" & Maxnum + 1 & ".jpg", 100, Me.Picture1
  DoEvents
End Sub

Private Sub lvButtons_H6_Click()
 Timer1.Enabled = False
    If mCapHwnd <> 0 Then StopWork
    Unload Me
    End

End Sub

Private Sub Timer1_Timer()
On Error Resume Next

    ' pause the timer
    Timer1.Enabled = False

    ' get the next frame;
    Call SendMessage(mCapHwnd, WM_CAP_GET_FRAME, 0, 0)

    ' copy the frame to the clipboard
    Call SendMessage(mCapHwnd, WM_CAP_COPY, 0, 0)

    ' For some reason, the API is not resizing the video
    ' feed to the width and height provided when the video
    ' feed was started, so we must resize the image here
    ' Image1.Stretch = True
            
    ' get from the clipboard
    Image1.Picture = Clipboard.GetData
         
         
    ' restart the timer
    DoEvents
    If Not bStopped Then
        Timer1.Enabled = True
    End If

End Sub
