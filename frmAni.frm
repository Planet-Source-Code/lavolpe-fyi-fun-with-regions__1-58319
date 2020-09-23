VERSION 5.00
Begin VB.Form frmAni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Masking vs Clipping -- A Unique Comparison"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerRegions 
      Left            =   3375
      Top             =   960
   End
   Begin VB.Timer timerMasked 
      Left            =   3375
      Top             =   540
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Start Simple Animation Example"
      Height          =   615
      Left            =   3990
      TabIndex        =   1
      Top             =   465
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   45
      Picture         =   "frmAni.frx":0000
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   465
      Width           =   3900
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   4575
      Picture         =   "frmAni.frx":10134
      Top             =   3345
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAni.frx":12D76
      Height          =   1695
      Left            =   4050
      TabIndex        =   3
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAni.frx":12E49
      Height          =   450
      Left            =   90
      TabIndex        =   2
      Top             =   15
      Width           =   5955
   End
End
Attribute VB_Name = "frmAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MORE FUN WITH REGIONS......

' This is not to initiate a debate on masking vs clipping regions. It is
' only to raise some eyebrows on the possibility of using shaped regions
' instead of masking in a limited manner.

' This example is using an image and simply moving it to random locations
' on the DC. The 1st 10 frames are shown using on the fly masking routines;
' the last 10 frames are used without masking but using a clipping region
' in place of the mask.

' This example may not apply if using a range of individual images, like
' animated GIFs do. This is because you would need to create, on the fly
' or caching, a shaped region for each image. If that is the case, then
' creating a mask on the fly is far more efficient & about as fast!
' Comparisons are coming up.

' Masking (drawing a transparent colored bitmap), requires some work.
' Generally the routine is a bit difficult, but not too much. It requires
' creating several DC's, a few bitmaps, some graphics magic using
' BitBlt with various flags like vbSrcCopy, vbSrcAnd, vbSrcPaint and/or
' other settings. So: a lot of code, cpu cycles and resources
' used to draw every transparent bitmap or frame.

' However, by using a CACHED shaped region for the image, there is
' no comparison.  This is much faster for single images and comes
' at a small price: caching the region which means one more GDI object
' taken from the GDI heap. However, you don't need any transparency
' functions either. Hmmm? Another scenario could be used where you
' create the shaped region on the fly as needed -- depends on how often
' this region would be created I guess.

' SINGLE IMAGE COMPARSION
'========================
' Example using a single image & running 500 iterations...
'   Creating Mask on the fly: about 219 ms average every 500 iterations
'   Clipping: about 16 ms average every 500 iterations

' MULTIPLE IMAGE COMPARSION
'==========================
' Example using 8 images & running 100 iterations (800 total iterations)....
'   Creating Mask on the fly: 1,141 ms
'   Creating Shaped Region on the fly: 5,515 ms
'   However, if caching the 8 regions: 825 ms
'       ^^ faster than masking but keeps more GDI objects (regions) in memory

' IMHO: Use masking routines over clipping regions just
' for the simple fact that you won't need to cache an extra GDI object.

Option Explicit

' APIs used in these examples
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InvertRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private tmpDC As Long   ' temp DC for back buffering
Private tmpBmp As Long  ' temp Bmp to hold frame before drawing
Private dRect As RECT   ' location of drawing on dc
Private rgnMask As Long ' shaped region used vs masks
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cmdAnimate_Click()
If tmpDC = 0 Then   ' haven't done this yet, do it one time only
    Dim hBmp As Long
    ' create the temporary DC & bmp for backbuffering
    ' used when drawing with both masks & regions
    tmpDC = CreateCompatibleDC(Me.hdc)
    hBmp = CreateCompatibleBitmap(Me.hdc, Image1.Width \ Screen.TwipsPerPixelX, Image1.Height \ Screen.TwipsPerPixelY)
    tmpBmp = SelectObject(tmpDC, hBmp)
    ' create the shaped region used as a mask
    rgnMask = CreateShapedRegion2(Image1.Picture.Handle)
End If
Picture1.Cls
' start the masked drawing first
timerMasked.Interval = 150
Randomize Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
If tmpDC Then
    ' delete some memory objects
    DeleteObject SelectObject(tmpDC, tmpBmp)
    DeleteDC tmpDC
    DeleteObject rgnMask
End If
End Sub

Private Sub ShowNextFrame(maskID As Integer)
    
    Dim x As Long, y As Long
    Dim sRect As RECT
    
    ' replace original frame from backbuffer
    ErasePrevImage False
    
    'get next random X,Y coords
    x = CLng(Rnd * (Picture1.Width - Image1.Width))
    y = CLng(Rnd * (Picture1.Height - Image1.Height))
    
    'set up rectangle to identify where drawing will take place
    ' convert coords to pixels vs twips
    dRect.Left = x \ Screen.TwipsPerPixelX
    dRect.Top = y \ Screen.TwipsPerPixelY
    dRect.Right = (x + Image1.Width) \ Screen.TwipsPerPixelX - 1
    dRect.Bottom = (y + Image1.Height) \ Screen.TwipsPerPixelY - 1
    
    ' save the destination area to the backbuffer
    BitBlt tmpDC, 0, 0, dRect.Right - dRect.Left + 1, dRect.Bottom - dRect.Top + 1, Picture1.hdc, dRect.Left, dRect.Top, vbSrcCopy
    
    If maskID Then
        ' using the mask; significantly less memory objects used &
        ' significantly less time...
        
        ' move the region to where we want to draw
        OffsetRgn rgnMask, dRect.Left, dRect.Top
        
        ' make it the clipping region
        SelectClipRgn Picture1.hdc, rgnMask
        
        ' draw the image however you want; using Render, PaintPicture, BitBlt, etc
        Image1.Picture.Render Picture1.hdc + 0, dRect.Left + 0, dRect.Top + 0, _
            dRect.Right - dRect.Left + 1, dRect.Bottom - dRect.Top + 1, _
            0, Image1.Picture.Height, Image1.Picture.Width, -Image1.Picture.Height, ByVal 0&
        
        ' for testing only, want to distinguish it from masked drawing
        InvertRgn Picture1.hdc, rgnMask
        
        ' done; reset the mask region offsets & refresh the DC
        OffsetRgn rgnMask, -dRect.Left, -dRect.Top
        RedrawWindow Picture1.hwnd, dRect, ByVal 0&, 1
        
        ' remove clipping region
        SelectClipRgn Picture1.hdc, ByVal 0&
        
    Else
        ' call complicated masking routines
        ' Granted that the one I'm calling here is more complex than some
        ' but the results are the same for a single frame:
        '   masking takes longer than clipping & uses more resources
        DrawTransparentBitmap Picture1.hdc, dRect, Image1.Picture, sRect, , , , tmpDC, , , , , Picture1.hwnd
    
    End If

End Sub

Private Sub ErasePrevImage(bResetAll As Boolean)
' replace original frame

    BitBlt Picture1.hdc, dRect.Left, dRect.Top, dRect.Right - dRect.Left + 1, dRect.Bottom - dRect.Top + 1, tmpDC, 0, 0, vbSrcCopy
    RedrawWindow Picture1.hwnd, dRect, ByVal 0&, 1
    
    If bResetAll Then   ' reset rectangle structure
        dRect.Left = 0
        dRect.Right = 0
        dRect.Bottom = 0
        dRect.Top = 0
    End If

End Sub

Private Sub timerMasked_Timer()
' timer used for mask drawing; 10 iterations & setup region drawing
If Val(timerMasked.Tag) < 10 Then
    
    ShowNextFrame 0
    timerMasked.Tag = Val(timerMasked.Tag) + 1
    
Else
    
    ErasePrevImage True
    timerMasked.Interval = 0
    timerMasked.Tag = 0
    timerRegions.Interval = 150

End If

End Sub


Private Sub timerRegions_Timer()
' timer used for region drawing; 10 iterations and then stop
If Val(timerRegions.Tag) < 10 Then
    
    ShowNextFrame 1
    timerRegions.Tag = Val(timerRegions.Tag) + 1
    
Else
    
    timerRegions.Interval = 0
    timerRegions.Tag = 0
    dRect.Left = 0
    dRect.Right = 0
    dRect.Bottom = 0
    dRect.Top = 0

End If
    
    
End Sub
