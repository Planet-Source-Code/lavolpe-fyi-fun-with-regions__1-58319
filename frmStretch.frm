VERSION 5.00
Begin VB.Form frmStretch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stretching a Region"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Resize Region"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   4200
      Width           =   855
   End
   Begin VB.PictureBox picXform 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2280
      ScaleHeight     =   2175
      ScaleWidth      =   4140
      TabIndex        =   0
      Top             =   1080
      Width           =   4140
   End
   Begin VB.Label Label2 
      Caption         =   $"frmStretch.frx":0000
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Resize above region.  Enter any whole/decimal value between -95% to +100%"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2235
      Left            =   120
      Top             =   1080
      Width           =   2070
   End
   Begin VB.Label Label2 
      Caption         =   $"frmStretch.frx":00A2
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "frmStretch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This example uses the StretchRegion function.  That function uses an API
' to treat a region like StretchBlt treats bitmaps.  I find the results
' are poor when reducing regions but ok when enlarging regions.

' IMHO, there is no real good way to reduce regions when those regions may contain
' lines, edges or whatever that consist of one or two pixel widths

' APIs used for this example...
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Command1_Click()
' validate textbox
If IsNumeric(Text1) = False Then
    MsgBox "Only numeric values acceptable", vbInformation + vbOKOnly
    Exit Sub
End If
If Val(Text1.Text) < -95 Or Val(Text1.Text) > 100 Then
    MsgBox "For this example, only values between -95 and +100 are acceptable", vbInformation + vbOKOnly
    Exit Sub
End If

Dim X As Single, hRgn As Long, newRgn As Long
Dim hBrush As Long

' the passed value cannot be zero; otherwise you risk a null region being returned
' We will modify the text box to a value  =>.05 and <= 2.0
X = Val(Text1)

' xScale is percentage of increase or decrease in width
' yScale is percentage of increase or decrease in height
' (i.e., 1.5 for 50% increase and 0.5 for 50% decrease)
If X < 0 Then
    X = 1 - Abs(X / 100)
Else
    X = 1 + X / 100
End If

' get the shaped region
hRgn = CreateShapedRegion2(Image1.Picture)
If hRgn Then
    ' now stretch that region
    ' Note I am passing the same X value twice for proper scale.
    ' However, you can stretch by X and/or Y
    newRgn = StretchRegion(hRgn, X, X)
    ' destroy source region
    DeleteObject hRgn
    If newRgn Then
        ' resize destination
        picXform.Move picXform.Left, picXform.Top, Image1.Width * X, Image1.Height * X
        picXform.Cls
        ' set the window region: newRgn destroyed by Windows, not by us!
        SetWindowRgn picXform.hwnd, newRgn, True
        With picXform
            .PaintPicture Image1, 0, 0, .Width, .Height
            .Refresh
        End With
    End If
End If
End Sub

Private Sub Form_Load()
' I am removing the blue border around the test region.
' It is a 1-pixel line that is not stretched well. See notes at top of this
' module. If you want to see how bad shrinking affects single pixel lines
' rem out the entire section between the With:EndWith statements below

Dim tPic As StdPicture
Dim bluePix As Long, bkgPix As Long
Dim hBrush As Long, hRgn As Long

' get a copy of my test image
Set Image1.Picture = frmRotation.Image1.Picture
' close the form if we loaded it
If frmRotation.Visible = False Then Unload frmRotation

' remove the blue single-pixel border
With picXform
    ' resize picBox & paint the image
    .Move .Left, .Top, Image1.Width, Image1.Height
    .PaintPicture Image1, 0, 0, .Width, .Height
    ' get the blue color & background color of the image
    bluePix = GetPixel(.hdc, 0, 1)
    bkgPix = GetPixel(.hdc, 0, 0)
    ' return the anti-region & fill the blue with bkg color
    hRgn = CreateShapedRegion2(Image1.Picture, , bluePix, True)
    hBrush = CreateSolidBrush(bkgPix)
    FillRgn .hdc, hRgn, hBrush
    DeleteObject hBrush
    DeleteObject hRgn
    ' now copy the image into a stdPicture object
    ' Note: by providing a bad parameter (-1), source image is
    ' returned as a stdPicture. Explained in function's remarks
    hRgn = RotateImageRegion(.Image.Handle, , , , -1, tPic)
    If hRgn Then    ' success
        ' delete the region; not used & reassign the Image1 control
        DeleteObject hRgn
        Set Image1.Picture = tPic
    Else
        ' if this failed, then the Image1 picture's won't be changed
    End If
    .Cls
End With

Call Command1_Click
End Sub
