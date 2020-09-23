Attribute VB_Name = "modTransBitmap"
Option Explicit

' NOTE: USE AT YOUR OWN RISK...
' THIS WAS CUT & PASTE OUT OF OTHER PROJECTS I HAVE
' JUST SO I COULD PROVIDE THE ANIMATED BITMAP EXAMPLE.

Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

' GDI32 Function Calls
' =====================================================================
' DC manipulation
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
' Other drawing functions
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' USER32 Function Calls
' =====================================================================
' General Windows related functions
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' Standard TYPE Declarations used
' =====================================================================
'Public Type RECT                    ' used to set/ref boundaries of a rectangle
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type
Private Type BITMAP                  ' used to determine if an image is a bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const DSna = &H220326 '0x00220326


Public Sub DrawTransparentBitmap(lHDCdest As Long, DestRect As RECT, _
                                                    lBMPsource As Long, bmpRect As RECT, _
                                                    Optional lMaskColor As Long = -1, _
                                                    Optional lNewBmpCx As Long, _
                                                    Optional lNewBmpCy As Long, _
                                                    Optional lBkgHDC As Long, _
                                                    Optional bkgX As Long, _
                                                    Optional bkgY As Long, _
                                                    Optional FlipHorz As Boolean = False, _
                                                    Optional FlipVert As Boolean = False, _
                                                    Optional hWndToRefresh As Long)
' Above parameters are described...
' lHDCdest is the DC where the drawing will take place
' destRect is a RECT type indicating the left, top, right & bottom coords where drawing will be done
' lBMPsource is the handle to the bitmap to be made transparent and be re-drawn on lHDCdest
' bmpRect is a Rect type indicating the source bitmap's coords to use for drawing
'   -- Note: If null Rect, the entire bitmap is used.
' lMaskColor is the bitmap color to be made transparent. The value of -1 picks the top left corner pixel
' lNewBmpCx is the destination width of the source bitmap
'  -- Note: If not provided, the bitmap width is drawn with a 1:1 ratio
' lNewBmpCy is the destination height of the source bitmap
' -- Note: If not provided, the bitmap height is drawn with a 1:1 ratio
' ************ Following parameters are used if a separate HDC is used as a background or mask
'                 to be used for drawing. This option is used primarily as a background for animation
' lBkgHDC is the DC of the background image container
' bkgX, bkgYare the upper left/top coords to use on the background/mask DC for drawing on the
'   the destination DC. The width and height are determined by destRect's overall width/height

'-----------------------------------------------------------------
    Dim udtBitMap As BITMAP
    Dim lMask2Use As Long 'COLORREF
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long, lBMsrcOld As Long
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long
    Dim lHDCMem As Long, lHDCscreen As Long, lHDCsrc As Long, lHDCMask As Long, lHDCcolor As Long
    Dim OrientX As Long, OrientY As Long
    Dim x As Long, y As Long, srcX As Long, srcY As Long
    Dim lRatio(0 To 1) As Single
'-----------------------------------------------------------------
    Dim hPalOld As Long, hPalMem As Long
    lHDCscreen = GetDC(0&)
    lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
    
    lBMsrcOld = SelectObject(lHDCsrc, lBMPsource) 'Select the bitmap
    GetGDIObject lBMPsource, Len(udtBitMap), udtBitMap
    lMask2Use = lMaskColor
    If lMask2Use < 0 Then lMask2Use = GetPixel(lHDCsrc, 0, 0)
    'OleTranslateColor lMask2Use, 0, lMask2Use
    
    ' Bmp size needed for original source
        srcX = udtBitMap.bmWidth                  'Get width of bitmap
        srcY = udtBitMap.bmHeight                 'Get height of bitmap
        If lNewBmpCx = 0 Then
            If bmpRect.Right > 0 Then lNewBmpCx = bmpRect.Right - bmpRect.Left Else lNewBmpCx = srcX
        End If
        'Use passed width and height parameters if provided
        If lNewBmpCy = 0 Then
            If bmpRect.Bottom > 0 Then lNewBmpCy = bmpRect.Bottom - bmpRect.Top Else lNewBmpCy = srcY
        End If
        
        If bmpRect.Right = 0 Then bmpRect.Right = srcX Else srcX = bmpRect.Right - bmpRect.Left
        If bmpRect.Bottom = 0 Then bmpRect.Bottom = srcY Else srcY = bmpRect.Bottom - bmpRect.Top
    ' Calculate size needed for drawing
        If (DestRect.Right) = 0 Then x = lNewBmpCx Else x = (DestRect.Right - DestRect.Left)
        If (DestRect.Bottom) = 0 Then y = lNewBmpCy Else y = (DestRect.Bottom - DestRect.Top)
'=========================================================================
' This routine will fail to draw properly if you try to draw a  larger image (lNewBmpCX or lNewBmpCy
' than is larger than the destination dimensions. Therefore, if the source dimensions are larger, then
' the routine will attempt to automatically scale the source image as needed.
'=========================================================================
        If lNewBmpCx > x Or lNewBmpCy > y Then
            lRatio(0) = (x / lNewBmpCx)
            lRatio(1) = (y / lNewBmpCy)
            If lRatio(1) < lRatio(0) Then lRatio(0) = lRatio(1)
            lNewBmpCx = lRatio(0) * lNewBmpCx
            lNewBmpCy = lRatio(0) * lNewBmpCy
            Erase lRatio
        End If
            
    
    'Create some DCs to hold temporary data
    lHDCMask = CreateCompatibleDC(lHDCscreen)
    lHDCMem = CreateCompatibleDC(lHDCscreen)
    lHDCcolor = CreateCompatibleDC(lHDCscreen)
    'Create a bitmap for each DC.  DCs are required for a number of GDI functions
    'Compatible DC's
    lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
    lBmAndMem = CreateCompatibleBitmap(lHDCscreen, x, y)
    lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
    
    'Each DC must select a bitmap object to store pixel data.
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
    ReleaseDC 0&, lHDCscreen
    
' ====================== Start working here ======================
    
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    hPalMem = SelectPalette(lHDCMem, 0, True)
    RealizePalette lHDCMem
    'Copy the background of the main DC to the destination
    If (lBkgHDC <> 0) Then
            BitBlt lHDCMem, 0, 0, x, y, lBkgHDC, bkgX, bkgY, vbSrcCopy
    Else
            BitBlt lHDCMem, 0&, 0&, x, y, lHDCdest, DestRect.Left, DestRect.Top, vbSrcCopy
    End If
    
    'Set proper mapping mode.
    hPalOld = SelectPalette(lHDCcolor, 0, True)
    RealizePalette lHDCcolor
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    
    ' Get working copy of source bitmap
    'StretchBlt lHDCcolor, srcX, 0, -srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, srcX, srcY, vbSrcCopy
    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
'    If FlipHorz Then StretchBlt lHDCcolor, srcX, 0, -srcX, srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
'    If FlipVert Then StretchBlt lHDCcolor, 0, srcY, srcX, -srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
    ' set working color back/fore colors. These colors will help create the mask
    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite
    
    'Create the object mask for the bitmap by performaing a BitBlt
    BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    
    ' This will create a mask of the source color
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna

    'Mask out the places where the bitmap will be placed while resizing as needed
    StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
    
    'XOR the bitmap with the background on the destination DC while resizing as needed
    StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
    
    'Copy to the destination
    BitBlt lHDCdest, DestRect.Left, DestRect.Top, x, y, lHDCMem, 0&, 0&, vbSrcCopy
    'StretchBlt lHDCdest, destRect.Left, destRect.Top, X, Y, lHDCMem, 0, 0, X, Y, vbSrcCopy
    
    
    'Delete memory bitmaps
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    SelectObject lHDCsrc, lBMsrcOld
    
    'Delete memory DC's
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
    DeleteDC lHDCsrc
    
    If hWndToRefresh Then RedrawWindow hWndToRefresh, DestRect, ByVal 0&, 1
'-----------------------------------------------------------------
End Sub


