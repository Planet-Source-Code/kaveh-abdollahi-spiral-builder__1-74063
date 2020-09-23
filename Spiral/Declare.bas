Attribute VB_Name = "Declare"
Option Explicit

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal xW As Long, ByVal yW As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal xW As Long, ByVal yW As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dW As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public pot As POINTAPI
Public PadBytes As Long
Public BytesPerScanLine As Long
Public Type BITMAPINFOHEADER '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public m_hDIb As Long, m_hBmpOld As Long
Public m_hDC As Long, DIBPtr As Long


Public Primes() As Long, nPR() As Byte, nPR2() As Byte, gap() As Byte              ' About 10%  Byte Farster Than Boolian

Public Sub PrimeBase()
Dim Lp1 As Long, Lp2 As Long, sqR1 As Long, ST As Long, sqR2 As Long, PrCount As Long, T1 As Long, LPrime As Long
ReDim nPR(1 To 200000000)
ReDim nPR2(1 To 200000000)
ReDim gap(1 To 11078938)
ReDim Primes(1 To 11078938)
    
    ST = 3
    sqR2 = 200000000

Rx: ''''''''' Start ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    sqR1 = sqR2 ^ 0.5
        For Lp1 = 3 To sqR1 Step 2
            
            If nPR(Lp1) = False Then
               
               For Lp2 = Lp1 To sqR2 Step Lp1 * 2
                 nPR(Lp2) = True
               Next Lp2
            Else
            
            End If
            
            If Lp1 Mod 17 = 1 Then DoEvents
        Next Lp1

'''''''''''''''''''''''''
    
        For Lp1 = 3 To sqR1 Step 2
            nPR(Lp1) = False
        Next Lp1
        
    If sqR1 > 5 Then sqR2 = sqR1:  GoTo Rx
    
'''''''''' Finish ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Primes(1) = 2: Primes(2) = 3
    PrCount = 3: LPrime = 3
    For Lp2 = 3 To 200000000 Step 2
        If nPR(Lp2) <> True Then
            Primes(PrCount) = Lp2
            PrCount = PrCount + 1
            If PrCount Mod 123456 = 0 Then frmBase.lblPrCount.Caption = PrCount - 1: DoEvents
        End If
    Next Lp2
    gap(1) = 1: gap(2) = 2
    For Lp2 = 3 To PrCount - 2
        gap(Lp2) = Primes(Lp2 + 1) - Primes(Lp2)
    Next Lp2
    
    frmBase.lblPrCount.Caption = PrCount - 1

End Sub

Public Sub SETBMI()
Dim SBI As BITMAPINFOHEADER

   With SBI
      .biSize = 40
      .biWidth = frmBase.pic1.Width / Screen.TwipsPerPixelX
      .biHeight = frmBase.pic1.Height / Screen.TwipsPerPixelY
      .biPlanes = 1
      .biBitCount = 32 '24
      .biCompression = 0

      BytesPerScanLine = (((.biWidth * .biBitCount) + 31) / 32) * 4
      PadBytes = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) / 8)
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)

      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With

   m_hDC = CreateCompatibleDC(0)
   m_hDIb = CreateDIBSection(m_hDC, SBI, 0, DIBPtr, 0, 0)
   m_hBmpOld = SelectObject(m_hDC, m_hDIb)
End Sub

Public Sub SaveJpeg(FSpec$, ByVal TheQuality As Long, APIC As PictureBox)
Dim pvGDI As GDIPlusJPGConvertor
   SETBMI
   
   BitBlt m_hDC, 0, 0, frmBase.pic1.Width / Screen.TwipsPerPixelX, frmBase.pic1.Height / Screen.TwipsPerPixelY, APIC.hdc, 0, 0, vbSrcCopy
  
   Set pvGDI = New GDIPlusJPGConvertor
   
   pvGDI.SaveDIB frmBase.pic1.Width / Screen.TwipsPerPixelX, frmBase.pic1.Height / Screen.TwipsPerPixelY, DIBPtr, FSpec$, TheQuality
 
   Set pvGDI = Nothing
    
   SelectObject m_hDC, m_hBmpOld
   DeleteObject m_hDIb
   DeleteDC m_hDC
End Sub



Public Function GetDistance(ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long) As Long

    Dim sngDx As Single
    Dim sngDy As Single

    sngDx = lX2 - lX1
    sngDy = lY2 - lY1
    
    GetDistance = Sqr(sngDx * sngDx + sngDy * sngDy)

End Function


