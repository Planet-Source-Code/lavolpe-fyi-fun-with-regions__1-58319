VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFloodFill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Swapping Example"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdColor 
      Caption         =   "Start Over"
      Height          =   390
      Index           =   0
      Left            =   4350
      TabIndex        =   6
      Top             =   405
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   6150
      Top             =   345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "New Color >>"
      Height          =   390
      Index           =   1
      Left            =   2265
      TabIndex        =   3
      Top             =   405
      Width           =   1335
   End
   Begin VB.PictureBox picFF 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   105
      ScaleHeight     =   3795
      ScaleWidth      =   6540
      TabIndex        =   0
      Top             =   900
      Width           =   6600
   End
   Begin VB.Label Label2 
      Caption         =   "Click on image color below to be replaced"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   405
      Width           =   1665
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   3675
      TabIndex        =   4
      Top             =   435
      Width           =   240
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   1875
      TabIndex        =   2
      Top             =   435
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Example of using regions to swap colors for an entire image (pretty quick)"
      Height          =   255
      Left            =   210
      TabIndex        =   1
      Top             =   75
      Width           =   6465
   End
End
Attribute VB_Name = "frmFloodFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' USING A REGION TO FLOOD FILL A DC

' This example simply swaps one color with another using a shaped region
' and the FillRgn API.

' APIs used for this sample form
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub cmdColor_Click(Index As Integer)

If Index Then
    ' allow user to select a replacement color
    With dlgColor
        .Color = lblColor(1).BackColor
        .CancelError = True
        .Flags = cdlCCRGBInit
    End With
    On Error GoTo ExitRoutine
    dlgColor.ShowColor
    lblColor(1).BackColor = dlgColor.Color
    
' The region returned by the function should always be
' deleted when no longer needed; with an exception > whenever you
' apply a region using SetWindowRgn API, per MSDN, Windows owns that
' region and you are not to play with it any longer.
    
   
    Dim antiRgn As Long, hBrush As Long
    
    ' create the replacement color brush
    hBrush = CreateSolidBrush(lblColor(1).BackColor)
    ' create a region of only the color to be replaced
    antiRgn = CreateShapedRegion2(picFF.Image, 0, lblColor(0).BackColor, True)
    If antiRgn Then
        ' now use that region with FillRgn API
        FillRgn picFF.hdc, antiRgn, hBrush
        DeleteObject antiRgn
        lblColor(0).BackColor = lblColor(1).BackColor
    Else
        ' no antiRegion. Error occurred or the target color does not
        ' exist in the bitmap so no replacement color can be applied
    End If
    DeleteObject hBrush
    
Else

    ' Start Over: replace image with source from main form
    Call Form_Load

End If

ExitRoutine:
End Sub

Private Sub Form_Load()
' get play image from main form
picFF.Cls
With frmShapedRgns.Picture1.Picture
    .Render picFF.hdc, 0, 0, ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
        0, .Height, .Width, -.Height, ByVal 0&
End With
lblColor(0).BackColor = GetPixel(picFF.hdc, 0, 0)
End Sub

Private Sub picFF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    ' return color form picture box
    lblColor(0).BackColor = GetPixel(picFF.hdc, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
End If
End Sub
