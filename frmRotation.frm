VERSION 5.00
Begin VB.Form frmRotation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Region Rotation - Non Image-Dependent"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Index           =   1
      Left            =   2280
      ScaleHeight     =   2235
      ScaleWidth      =   2070
      TabIndex        =   3
      Top             =   1320
      Width           =   2070
   End
   Begin VB.ComboBox cboRotation 
      Height          =   315
      ItemData        =   "frmRotation.frx":0000
      Left            =   2310
      List            =   "frmRotation.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2100
   End
   Begin VB.Image Image1 
      Height          =   2235
      Left            =   120
      Picture         =   "frmRotation.frx":0079
      Top             =   1320
      Width           =   2070
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Show as Rotated by"
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   915
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   $"frmRotation.frx":F2DB
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This sample hilights the use of the function RotateSimpleRegion.
' That function rotates any region and does not require an image
' to be supplied. Ideally used when the region being rotated is
' not filled by some picture or when the rotated image is simply
' outlined or filled using APIs like FrameRgn or FillRgn.

' The routines are very quick.

' APIs used for this example...
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Sub cboRotation_Click()

' lock until drawing is complete...
cboRotation.Locked = True

Dim testRgn As Long, newRgn As Long
Dim hBrush As Long

' get the sample region.
testRgn = CreateShapedRegion2(Image1.Picture)

If testRgn Then ' successful

    Picture1(1).Cls
    ' size the destination appropriately
    With Image1
        If cboRotation.ListIndex = 1 Or cboRotation.ListIndex = 3 Then
            ' 90/270 degree rotation: swap width/height sizes
            Picture1(1).Move Picture1(1).Left, .Top, .Height, .Width
        Else
            ' all other options: use same width/height sizes
            Picture1(1).Move Picture1(1).Left, .Top, .Width, .Height
        End If
    End With
    
    ' create a colored brush for drawing
    hBrush = CreateSolidBrush(vbRed)
    
    ' call function to rotate the region based on the testRgn.
    ' Note that the function will also accept a Window handle should
    ' you want to rotate the region of a Window vs cached region.
    newRgn = RotateSimpleRegion(testRgn, False, cboRotation.ListIndex - 1)
    
    If newRgn Then
        ' successful: fill rgn with color so you can see the results
        FillRgn Picture1(1).hdc, newRgn, hBrush
        DeleteObject newRgn ' remove memory object
    Else
        MsgBox "Routines couldn't create the rotated region!", vbInformation + vbOKOnly, "Oops"
    End If
    DeleteObject testRgn ' remove memory object
    DeleteObject hBrush ' remove memory object
    
End If

' finish up
Picture1(1).Refresh
cboRotation.Locked = False
End Sub

Private Sub Form_Load()
'Set Image1.Picture = frmShapedRgns.Picture1.Picture
cboRotation.ListIndex = 0
End Sub
