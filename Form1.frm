VERSION 5.00
Begin VB.Form frmShapedRgns 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shaped Regions & AntiRegions"
   ClientHeight    =   3570
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAntiRgn 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show Anti-Region Shaped Form also"
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   3075
      Value           =   1  'Checked
      Width           =   3825
   End
   Begin VB.ComboBox cboRotation 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1785
      List            =   "Form1.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2685
      Width           =   2100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close Test Forms"
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   990
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Shaped Regions"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   405
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   2235
      Left            =   1800
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":0079
      ScaleHeight     =   2235
      ScaleWidth      =   2070
      TabIndex        =   0
      Top             =   405
      Width           =   2070
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Show as Rotated by"
      Height          =   210
      Left            =   75
      TabIndex        =   6
      Top             =   2760
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2 Regions (forms) created: normal and anti-region"
      Height          =   240
      Left            =   165
      TabIndex        =   4
      Top             =   135
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag && drop any bmp, jpg or gif into the test frame. It may resize off this from, but it doesn't matter. Jpgs are worse"
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   1650
      Width           =   1710
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Play Time"
      Begin VB.Menu mnuSample 
         Caption         =   "&Clip Regions vs Masks"
         Index           =   0
      End
      Begin VB.Menu mnuSample 
         Caption         =   "&Anti Region"
         Index           =   1
      End
      Begin VB.Menu mnuSample 
         Caption         =   "More Region &Rotations"
         Index           =   2
      End
      Begin VB.Menu mnuSample 
         Caption         =   "&Stretch Regions"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmShapedRgns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Sub Command1_Click()
' create the shaped window and the anti-region

' Note: JPGs are the worse, GIFs are better & Bitmaps are the best
' image files to use because colors are more uniform between the
' original source image and the eventual saved file copy.

' One last note: the region returned by the function should always be
' deleted when no longer needed; with an exception > whenever you
' apply a region using SetWindowRgn API, per MSDN, Windows owns that
' region and you are not to play with it any longer.

Dim testForms(0 To 1) As FrmTest
Dim formLoop As Integer
Dim windowRgn As Long
Dim testPic As StdPicture
Dim testFormCx As Long
Dim testFormCy As Long

Call Command2_Click ' unload any open test forms

For formLoop = 0 To chkAntiRgn.Value
    Set testForms(formLoop) = New FrmTest
    ' call function to return the region
    If cboRotation.ListIndex > 0 Then
        windowRgn = RotateImageRegion(Picture1.Picture.Handle, 0, , formLoop = 1, cboRotation.ListIndex - 1, testPic)
    Else
        windowRgn = CreateShapedRegion2(Picture1.Picture.Handle, 0, , formLoop = 1)
        Set testPic = Picture1
    End If
    ' position and show the shaped window
    If windowRgn Then
        testFormCx = ScaleX(testPic.Width, vbHimetric, vbPixels) * Screen.TwipsPerPixelX
        testFormCy = ScaleY(testPic.Height, vbHimetric, vbPixels) * Screen.TwipsPerPixelY
        With testForms(formLoop)
            .Move (Screen.Width - testFormCx) \ 2, _
                (Screen.Height - testFormCy) \ 2, _
                testFormCx, testFormCy
            .AutoRedraw = True
            SetWindowRgn .hwnd, windowRgn, True
            Set .Picture = testPic
            ' using SetWindowRgn, so we don't use DeleteObject on the region
            .Show
        End With
    End If
Next

If windowRgn Then
    If chkAntiRgn.Value Then
        Label1.Caption = "They are stacked on each other. You can click and drag the test forms around anywhere on the visible areas."
    Else
        Label1.Caption = "You can click and drag the test form around anywhere on its visible areas."
    End If
    Command2.Enabled = True
End If

Set testForms(0) = Nothing
Set testForms(1) = Nothing
End Sub

Private Sub Command2_Click()
' close test form(s)

Dim I As Integer
For I = Forms.Count - 1 To 0 Step -1
    If Forms(I).Name = "FrmTest" Then Unload Forms(I)
Next
Command2.Enabled = False
Label1.Caption = "Drag && drop any bmp, jpg or gif into the test frame. It may resize off this from, but it doesn't matter. Jpgs are worse"
End Sub

Private Sub Form_Load()
cboRotation.ListIndex = 0
Command2.Enabled = False
chkAntiRgn = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Command2_Click
Unload frmFloodFill
End Sub

Private Sub mnuSample_Click(Index As Integer)
Select Case Index
Case 0: frmAni.Show
Case 1: frmFloodFill.Show
Case 2: frmRotation.Show
Case 3: frmStretch.Show
Case Else
End Select
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
' drag & drop to use selected file

If Data.Files.Count Then
    On Error Resume Next
    ' use only the first file if multiple files were dropped
    Set Picture1.Picture = LoadPicture(Data.Files(1))
    If Err Then
        MsgBox "Failed to load that file. Try another", vbInformation + vbOKOnly
        Err.Clear
    Else
'        ' update the flood fill example form with selected picture
'        frmFloodFill.picFF.Cls
'        With Picture1.Picture
'            .Render frmFloodFill.picFF.hdc, 0, 0, ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
'                0, .Height, .Width, -.Height, ByVal 0&
'        End With
'        ' show it again if user closed it
'        If frmFloodFill.Visible = False Then frmFloodFill.Show
    End If
End If
End Sub
