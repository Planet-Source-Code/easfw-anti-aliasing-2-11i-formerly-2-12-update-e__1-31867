VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   Caption         =   "anti-aliasing 2.11i"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   36
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'anti-alias v2.11i by easfw - fluoats@hotmail.com
'Mar 5, 2002

'F.Y.I. This program writes a "colorsav1.txt" into the same directory
'as the project or .exe file of this program.

'For a brief intro to the purpose of this program, run it, hold down
'right arrow key, watch which controls change.

'You can save your own settings by holding Shift while clicking a square,
'or you can swap the settings around by holding Shift and using arrow keys.

'want to draw some lines in your own program .. copy the first
'commented-out section past the normal code

'want to see a simplified example of these form controls
'in action, Copy the 2nd commented-out section

'instructions for immediate operation are included
'in either set

'thanks to 'AedSeed' for helping me close the program properly _
 in the middle of an infinite loop

Dim vLeft(1 To 59) As Integer, vRight(1 To 59) As Integer
Dim vTop(1 To 59) As Integer, vBot(1 To 59) As Integer
Dim vMax(1 To 59) As Long, vMin(1 To 59) As Integer
Dim vVal(1 To 59) As Single, vval2(3 To 5) As Single
Dim spotn As Byte, spotn2 As Byte, swapping As Boolean
Dim mcolor As Long

'swap settings
Dim swopa As Byte, swopi As Byte, swds As Byte
Dim swred As Byte, swgrn As Byte, swblu As Byte
Dim swri As Single, swgi As Single, swbi As Single

'hold settings
Dim svr(21 To 58) As Byte, svg(21 To 58) As Byte, svb(21 To 58) As Byte
Dim svri(21 To 59) As Single, svgi(21 To 59) As Single
Dim svbi(21 To 59) As Single, svds(21 To 59) As Byte
Dim svo(21 To 57) As Byte, svoa(21 To 57) As Byte

'opacity / intensity
Dim oscopa As Single, s1a As Single

Dim styleSelectEnabled As Boolean, iclearX As Integer, iclearY As Integer
Dim selectv As Byte, twistenabled As Boolean, opaenabled As Boolean

Dim ClearMe As Boolean, qRecall As Boolean
Dim yInit As Integer, wavebreak As Boolean

Dim drawselect As Byte 'select case for different pixels() light/opacity _
formulas .. picselect_mousedown() changes this value, and pixels() accesses it

'Sub colors()
Dim fading As Boolean
Dim ir1 As Single, ig1 As Single, ib1 As Single
Dim itr As Single, gra As Byte, nli As Byte
Dim imode As Boolean

Dim gs(0 To 255) As Single, mstep As Integer, int1 As Single
Dim breakloop As Boolean, movinglines As Boolean
Dim w As Integer, ribbon As Boolean, nottesting As Boolean
Dim r As Single, g As Single, b As Single, shape As Byte, pow As Single
Dim n As Integer, n2 As Long 'all-purpose
Dim incr As Single, incg As Single, incb As Single
Dim aliasing As Boolean  'accessed by waveform()
Dim randfirst As Boolean 'accessed by waveform()
Dim bwi As Byte       'accessed by Dither()
Dim bg(1 To 6) As Byte, lightbgchanged As Boolean
Dim sr(1 To 3000) As Long  'used by BGcopy() and BGpaste()
Dim savX(0 To 200000) As Integer
Dim savy(0 To 200000) As Integer
Dim savcL(0 To 200000) As Long
Dim savDS(0 To 200000) As Boolean
Dim ns1 As Long, s1 As Single
Dim bar As Long, barcol As Long, cL As Long, outline As Long, outline2 As Long

Dim stuffed As Boolean, keyns As Integer
Dim FreeFileN As Integer

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Sub Form_Load()
Form1.ScaleMode = 3: Form1.AutoRedraw = False
Form1.Width = 10420: Form1.Height = 8192
FreeFileN = FreeFile

drawselect = 0     '0, 1, 2, 3, 4 or 5 in this version (2.11i)
vVal(2) = 175       '0 = totally transparent - only works drawselect = 0,1,2,3 or 5
vVal(1) = 30       '255 = max "easfw-enhanced"  anti-aliasing
vVal(12) = 5       'line spacing
'vVal(17) = 200000  'how many pixels before tail begins to unwind
 
bwi = 1            '0 or 1 .. accessed by Dither()
 
bg(1) = 149: bg(2) = 172: bg(3) = 203
bg(1) = 198: bg(2) = 212: bg(3) = 205
bg(4) = 144: bg(5) = 70: bg(6) = 188
bg(4) = 50: bg(5) = 65: bg(6) = 110

'initial line color
vVal(3) = 255
vVal(4) = 169
vVal(5) = 0

'colorchange amount (try to keep it between 0 and 10)
vVal(18) = 0
vVal(19) = 1.11
vVal(20) = 3.01
 
'========
'Below here we have what could be thought of as control array properties
vLeft(1) = 10: vRight(1) = vLeft(1) + 9: vTop(1) = Form1.ScaleHeight - 228: vBot(1) = vTop(1) + 96
vLeft(2) = 23: vRight(2) = vLeft(2) + 9: vTop(2) = vTop(1): vBot(2) = vBot(1)
vLeft(3) = 10: vRight(3) = vLeft(3) + 10: vTop(3) = vBot(2) + 3: vBot(3) = vTop(3) + 70
vLeft(4) = 34: vRight(4) = vLeft(4) + 10: vTop(4) = vTop(3): vBot(4) = vBot(3)
vLeft(5) = 58: vRight(5) = vLeft(5) + 10: vTop(5) = vTop(4): vBot(5) = vBot(4)
vLeft(6) = 81: vRight(6) = vLeft(6) + 10: vBot(6) = vBot(5): vTop(6) = vTop(5) + 18
 
'These are tied to each other
vRight(10) = Form1.ScaleWidth - 11: vLeft(10) = vRight(10) - 10 'freak with the - n
vBot(10) = Form1.ScaleHeight - 9: vTop(10) = vBot(10) - 69 'freak with the - n
vRight(9) = vLeft(10) - 2: vLeft(9) = vRight(9) - 10: vTop(9) = vTop(10): vBot(9) = vBot(10)
vRight(8) = vLeft(9) - 2: vLeft(8) = vRight(8) - 10: vTop(8) = vTop(10): vBot(8) = vBot(10)
vRight(7) = vRight(10): vLeft(7) = vRight(7) - 13: vTop(7) = vTop(10) - 19: vBot(7) = vTop(7) + 13
 
'line spacing, number of lines ..
vLeft(12) = vLeft(8) - 16: vRight(12) = vLeft(12) + 8: vBot(12) = vBot(8): vTop(12) = vTop(8) + 12
vLeft(13) = vLeft(12) - 11: vRight(13) = vLeft(13) + 8: vBot(13) = vBot(12): vTop(13) = vTop(12) + 1
 
'clear
vLeft(11) = 10: vRight(11) = vLeft(11) + 46: vTop(11) = vBot(3) + 5: vBot(11) = vTop(11) + 22
 
'randfirst
vLeft(14) = vRight(6) + 6: vRight(14) = vLeft(14) + 11: vBot(14) = vBot(6): vTop(14) = vBot(14) - 11
 
'anti-aliasing
vLeft(15) = 10: vRight(15) = vLeft(15) + 87: vTop(15) = vBot(11) + 6: vBot(15) = vTop(15) + 16
 
'drawstyle selector
vLeft(16) = 105: vRight(16) = vLeft(16) + 60: vBot(16) = vBot(15): vTop(16) = vBot(16) - 8

'ribbon length
vRight(17) = vRight(13): vLeft(17) = vRight(17) - 8: vBot(17) = vTop(13) - 4: vTop(17) = vBot(17) - 46
vMax(17) = 200000
 
'rgb oscillation sliders
vLeft(18) = vRight(3) + 2: vRight(18) = vLeft(18) + 6: vTop(18) = vTop(3): vBot(18) = vBot(3)
vLeft(19) = vRight(4) + 2: vRight(19) = vLeft(19) + 6: vTop(19) = vTop(3): vBot(19) = vBot(3)
vLeft(20) = vRight(5) + 2: vRight(20) = vLeft(20) + 6: vTop(20) = vTop(3): vBot(20) = vBot(3)
 
'Save squares
For n = 21 To 30
vLeft(n) = vRight(16) - 200 + 11 * n: vRight(n) = vLeft(n) + 9: vTop(n) = vTop(15) - 4: vBot(n) = vTop(n) + 9
w = n + 10
vLeft(w) = vLeft(n): vRight(w) = vRight(n): vTop(w) = vBot(n) + 2: vBot(w) = vTop(w) + 9
 Next n
  
'Randomizer
vLeft(58) = vRight(40) + 25: vRight(58) = vLeft(58) + 70: vTop(58) = vTop(21) - 2: vBot(58) = vTop(58) + 22
 
'Temporary Store/Recall
vLeft(57) = vRight(58) + 8: vRight(57) = vLeft(57) + 43: vTop(57) = vTop(58): vBot(57) = vTop(57) + 22

'transparency oscillation slider
vLeft(59) = vRight(2) + 2: vRight(59) = vLeft(59) + 6: vTop(59) = vTop(2): vBot(59) = vBot(2)
vMax(59) = 255: vVal(59) = 30
oscopa = vVal(2) * vVal(59) / 65536
 
'top 2 sliders control line intensity and the 'no-twist', or anti-anti-aliasing effect
vMax(1) = 255
vMax(2) = 255

'rgb sliders
vMax(3) = 255
vMax(4) = 255
vMax(5) = 255
 
'colorshift amount (next to rgb)
vMax(6) = 16: vVal(6) = 6

'background rgb
vMax(8) = 255
vMax(9) = 255
vMax(10) = 255
 
'path characteristic
vMax(12) = 85
vMax(13) = 20000: vMin(13) = 10: vVal(13) = 9950
 
'colori sliders
vMax(18) = 10: vMax(19) = 10: vMax(20) = 10

'initial colorset data, in case no colorsav2.txt
For n = 21 To 40
svo(n) = 97: svoa(n) = 30
Next n
svr(21) = 166: svg(21) = 123: svb(21) = 158
svri(21) = 2.33: svgi(21) = 0.157: svbi(21) = 2.51
svds(21) = 1: svo(21) = 175: svoa(21) = 58
svr(22) = 179: svg(22) = 92: svb(22) = 35
svri(22) = 4.89: svgi(22) = 0.943: svbi(22) = -2.75
svo(22) = 175: svoa(22) = 40
svr(23) = 157: svg(23) = 64: svb(23) = 20
svri(23) = 3.31: svgi(23) = 0.49: svbi(23) = -0.81
svds(23) = 1: svo(23) = 210: svoa(23) = 82
svr(24) = 0: svg(24) = 0: svb(24) = 255
svri(24) = 10: svgi(24) = 10: svbi(24) = 0
svo(24) = 170: svoa(24) = 58
svr(25) = 144: svg(25) = 185: svb(25) = 241
svri(25) = 1.253: svgi(25) = -2.392: svbi(25) = -2.499
svo(25) = 88
svr(26) = 78: svg(26) = 177: svb(26) = 245
svri(26) = -1.934: svgi(26) = 0.093: svbi(26) = -3.73
svo(26) = 167
svr(27) = 9: svg(27) = 194: svb(27) = 29
svri(27) = -0.75: svgi(27) = 0.65: svbi(27) = -1.48
svds(27) = 2
svr(28) = 255: svg(28) = 76: svb(28) = 56
svri(28) = 4.24: svgi(28) = 0.967: svbi(28) = 1.672
svds(28) = 1: svo(28) = 205: svoa(28) = 69
svr(29) = 20: svg(29) = 63: svb(29) = 207
svri(29) = 3.656: svgi(29) = 1.318: svbi(29) = 4.861
svds(29) = 2
svr(30) = 255: svg(30) = 255: svb(30) = 255
svri(30) = 10: svgi(30) = 0: svbi(30) = 10
svo(30) = 210: svoa(30) = 61
svr(31) = 110: svg(31) = 201: svb(31) = 226
svri(31) = -4.766: svgi(31) = 0.085: svbi(31) = 3.0578
svds(31) = 1: svo(31) = 199
svr(32) = 14: svg(32) = 174: svb(32) = 11
svri(32) = 1.536: svgi(32) = -0.505: svbi(32) = -3.8708
svds(32) = 2
svr(33) = 58: svg(33) = 236: svb(33) = 10
svri(33) = 0.0088: svgi(33) = -1.05: svbi(33) = 3.201
svds(33) = 3: svo(33) = 72: svoa(33) = 0
svr(34) = 55: svg(34) = 137: svb(34) = 124
svri(34) = -1.84: svgi(34) = -0.275: svbi(34) = -4.825
svo(34) = 172: svoa(34) = 96
svr(35) = 250: svg(35) = 159: svb(35) = 120
svri(35) = 0.36: svgi(35) = 1.956: svbi(35) = 0.6917
svds(35) = 1: svo(35) = 210: svoa(35) = 98
svr(36) = 112: svg(36) = 18: svb(36) = 123
svri(36) = 2.8: svgi(36) = 4.93: svbi(36) = 4.56
svds(36) = 0: svo(36) = 175: svoa(36) = 30
svr(37) = 0: svg(37) = 255: svb(37) = 0
svri(37) = 10: svgi(37) = 0: svbi(37) = 0
svds(37) = 1: svo(37) = 135: svoa(37) = 53
svr(38) = 161: svg(38) = 57: svb(38) = 93
svri(38) = 0.998: svgi(38) = 1.16: svbi(38) = -1.5477


svr(40) = 149: svg(40) = 103: svb(40) = 238
svri(40) = 3.07: svgi(40) = 0.25: svbi(40) = -2.22
svds(40) = 1: svo(40) = 210: svoa(40) = 82

'This creates file if none exists
Open "colorsav1.txt" For Append As #FreeFileN
Close #FreeFileN
 
'If just created colorsav1.txt, will be empty.
'It will get filled when anti-aliasing 212 closes
Open "colorsav1.txt" For Input As #FreeFileN
For n = 21 To 40
If EOF(FreeFileN) Then Exit For
Input #FreeFileN, svr(n), svg(n), svb(n), svri(n), svgi(n), svbi(n), _
svds(n), svo(n), svoa(n)
Next n
If Not EOF(FreeFileN) Then Input #FreeFileN, _
vVal(3), vVal(4), vVal(5), vVal(18), vVal(19), vVal(20), drawselect, _
vVal(2), vVal(59)
If Not EOF(FreeFileN) Then Input #FreeFileN, _
bg(1), bg(2), bg(3), bg(4), bg(5), bg(6), bwi
Close #FreeFileN

Form1.ForeColor = vbWhite
ribbon = False
opaenabled = True
twistenabled = True
styleSelectEnabled = True
nottesting = True: qRecall = True
movinglines = 1: s1 = vVal(2) / 255
r = vVal(3): g = vVal(4): b = vVal(5)
incb = vVal(6) / 6
incr = vVal(18) * incb
incg = vVal(19) * incb
incb = vVal(20) * incb
s1a = vVal(2) / 255
If opaenabled Then oscopa = vVal(2) * vVal(59) / 65536

Select Case bwi 'light or dark background, what color settings
Case 0
 vVal(8) = bg(4)
 vVal(9) = bg(5)
 vVal(10) = bg(6)
 barcol = RGB(99, 99, 99)
Case 1
 vVal(8) = bg(1)
 vVal(9) = bg(2)
 vVal(10) = bg(3)
 barcol = RGB(110, 106, 156)
End Select
End Sub

Private Sub Form_Paint()
Dim rr As Integer, gg As Integer, bb As Integer
Static initclick As Boolean
 
 If Not initclick Then
  Dither Me
  Randomize: initclick = 1
  Form1.CurrentX = 70 + 320 * Rnd: Form1.CurrentY = 80 + 320 * Rnd
  iclearX = Form1.CurrentX: iclearY = Form1.CurrentY
  Form1.Font = "New Times Roman": Form1.FontSize = 27
  Form1.Print "cLicK aNYwheRe"
  Form1.Font = "Times New Roman"
 End If
 
 If drawselect = 1 Or aliasing = True Then
  twistenabled = False
 End If

 If drawselect = 4 Or aliasing = True Then
  opaenabled = False
 End If
 
 If bwi Then
  Form1.ForeColor = RGB(100, 75, 190)
  outline = RGB(70, 86, 90)
  outline2 = RGB(54, 78, 81)
 Else
  Form1.ForeColor = RGB(185, 140, 255)
  outline = RGB(140, 70, 216)
  outline2 = RGB(131, 10, 216)
  Form1.ForeColor = RGB(0, 195, 198)
  outline = RGB(50, 0, 196)
 End If

 'outlines of buttons and sliders
 For n = 1 To 59 Step 1
  If n <> 14 And n <> 16 And n <> 17 Then
   If n > 20 And n < 41 Or n = 57 Then
   Form1.Line (vLeft(n), vTop(n))-(vRight(n), vBot(n)), outline2, B 'RGB(svr(n), svg(n), svb(n)), B
   Else
   Form1.Line (vLeft(n), vTop(n))-(vRight(n), vBot(n)), outline, B
    End If
    
   'slider bars
   Select Case n
   Case 1 To 6, 8 To 10, 12, 13, 59
    bar = vBot(n) - (vBot(n) - vTop(n) - 4) * (vVal(n) - vMin(n)) / (vMax(n) - vMin(n)) - 2
    If n = 1 And twistenabled Or n = 2 And opaenabled Or n > 2 Then
     If n = 59 And opaenabled Or n <> 59 Then
     Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), barcol
     End If
    End If
   End Select
  End If
 Next n
 
 For n = 18 To 20
 updateCLibar
 Next n
  
 Form1.FontSize = 10
 Form1.CurrentX = vLeft(11) + 10: Form1.CurrentY = vTop(11) + 4
 Form1.Print "Clear"
 
 For n = vTop(15) + 1 To vBot(15) - 1 Step 1
  cL = GetPixel(Form1.hdc, 1, n)
  Form1.Line (vLeft(15) + 1, n)-(vRight(15), n), cL
 Next n
 Form1.FontSize = 9
 Form1.CurrentX = vLeft(15) + 3: Form1.CurrentY = vTop(15) + 1
 Select Case aliasing
 Case False
  Form1.Print "Anti-Aliasing on"
 Case Else
  Form1.Print "Anti-Aliasing off"
 End Select
 yellowsquarepaint
 
 drawStoreRecallbutton
 
 Form1.CurrentX = vLeft(58) + 6: Form1.CurrentY = vTop(58) + 4
 Form1.Print " Randomize"
 
 Select Case aliasing
 Case False
  Dim st As Integer
  
  Select Case bwi
  Case 1
   outline = RGB(195, 195, 195)
   n = 180: n2 = 230: st = 7: rr = 145: gg = 135: bb = 95
  Case Else
   outline = RGB(135, 135, 135)
   n = 80: n2 = 30: st = -7: rr = 45: gg = 40: bb = 240
  End Select
  
  w = vTop(16) 'drawmode selector
  For n = n To n2 Step st
   Form1.Line (vLeft(16), w)-(vRight(16), w), RGB(n, n, n)
   w = w + 1: Next n
   Form1.Line (vLeft(16), w)-(vRight(16), w), outline
  For n = 1 To 5
   Form1.Line (vLeft(16) + 10 * n, vTop(16))-(vLeft(16) + 10 * n, vBot(16) + 1), RGB(rr, gg, bb)
  Next n
  n = vLeft(16) + 10 * drawselect + 1
  Form1.Line (n, vBot(16) - 1)-(n + 9, vBot(16) - 1), barcol - 612550
 End Select
 DrawSpotN
End Sub
Private Sub DrawSpotN()
If spotn > 20 And spotn < 42 Then Form1.Line (vLeft(spotn), vTop(spotn))-(vRight(spotn), vBot(spotn)), mcolor, B
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'
End Sub
Private Sub dblsngclickcommon()
  If selectv > 0 Then nli = 0
  Select Case selectv
  Case 0
  If wavebreak Then
  Select Case movinglines
  Case 0
  movinglines = 1
  Case 1
  movinglines = 0
  End Select
   End If
  
  Case 7
  clearsav
  If bwi Then
  bwi = 0
  If Not lightbgchanged Then
  vVal(8) = bg(4)
  vVal(9) = bg(5)
  vVal(10) = bg(6)
  Else
  vVal(8) = bg(4)
  vVal(9) = bg(5)
  vVal(10) = bg(6)
   End If
  barcol = RGB(99, 99, 99)
  Else
  bwi = 1
  vVal(8) = bg(1)
  vVal(9) = bg(2)
  vVal(10) = bg(3)
  barcol = RGB(110, 106, 156)
   End If
  Dither Me: Form_Paint
   
  Case 11
  clearsav
  If Not wavebreak Then
  ClearMe = True
  Call waveform
   End If
  Dither Me: Form_Paint
  
  Case 14
  If randfirst Then
  'andfirst = 0
  Else
  'andfirst = 1
   End If
  yellowsquarepaint
  
  Case 15
  If aliasing Then
  aliasing = 0
  styleSelectEnabled = True
  If drawselect <> 2 Then opaenabled = True
  If drawselect <> 1 Then twistenabled = True
  Else
  aliasing = 1
  styleSelectEnabled = False
  opaenabled = False
  twistenabled = False
  updateThoseTwoBars
  For n = vTop(16) To vBot(16) Step 1
  cL = GetPixel(Form1.hdc, 1, n)
  Form1.Line (vLeft(16), n)-(vRight(16), n), cL
   Next n
   End If
  Form_Paint
   
  Case 21 To 57
  If qRecall Then
  r = svr(selectv): vVal(3) = r
  g = svg(selectv): vVal(4) = g
  b = svb(selectv): vVal(5) = b
  incr = svri(selectv): vVal(18) = incr
  incg = svgi(selectv): vVal(19) = incg
  incb = svbi(selectv): vVal(20) = incb
  drawselect = svds(selectv)
  vVal(2) = svo(selectv): vVal(59) = svoa(selectv)
  s1 = vVal(2) / 255: s1a = s1
  oscopa = s1 * vVal(59) / 255
  clearsav
  
  If selectv <> 57 Then
  spotn = selectv: spotn2 = spotn: mcolor = RGB(r, g, b)
  ElseIf spotn > 20 Then
  mcolor = RGB(svr(spotn), svg(spotn), svb(spotn))
   End If
 
  If styleSelectEnabled Then updateThoseTwoBars
  
  Dither Me: Form_Paint
  
  Else
  If swapping And selectv <> 57 Then
  spotn = selectv
  swopa = svo(spotn): swopi = svoa(spotn): swds = svds(spotn)
  swred = svr(spotn): swgrn = svg(spotn): swblu = svb(spotn)
  swri = svri(spotn): swgi = svgi(spotn): swbi = svbi(spotn)
  svo(spotn) = svo(spotn2): svoa(spotn) = svoa(spotn2): svds(spotn) = svds(spotn2)
  svr(spotn) = svr(spotn2): svg(spotn) = svg(spotn2): svb(spotn) = svb(spotn2)
  svri(spotn) = svri(spotn2): svgi(spotn) = svgi(spotn2): svbi(spotn) = svbi(spotn2)
  svo(spotn2) = swopa: svoa(spotn2) = swopi: svds(spotn2) = swds
  svr(spotn2) = swred: svg(spotn2) = swgrn: svb(spotn2) = swblu
  svri(spotn2) = swri: svgi(spotn2) = swgi: svbi(spotn2) = swbi
  'r = svr(spotn): ..
  Form_Paint
  spotn2 = spotn
  Else
  svr(selectv) = vVal(3): svg(selectv) = vVal(4): svb(selectv) = vVal(5)
  svri(selectv) = vVal(18): svgi(selectv) = vVal(19): svbi(selectv) = vVal(20)
  svds(selectv) = drawselect: svo(selectv) = vVal(2): svoa(selectv) = vVal(59)
  spotn2 = selectv: cL = RGB(svr(spotn2), svg(spotn2), svb(spotn2)): mcolor = cL
  Form1.Line (vLeft(spotn2), vTop(spotn2))-(vRight(spotn2), vBot(spotn2)), cL, B
    End If
   End If
  
  Case 58
  vVal(3) = Rnd * 255: vVal(4) = Rnd * 255: vVal(5) = Rnd * 255
  r = vVal(3): g = vVal(4): b = vVal(5)
  incr = vMax(18) * (Rnd - 0.5): incg = vMax(19) * (Rnd - 0.5): incb = vMax(20) * (Rnd - 0.5)
  vVal(18) = incr: vVal(19) = incg: vVal(20) = incb
  Call clearsav: swapping = False
  Dither Me: Form_Paint
  nli = 0
   End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 81 Or KeyCode = 65 Then n = 3
If KeyCode = 87 Or KeyCode = 83 Then n = 4
If KeyCode = 69 Or KeyCode = 68 Then n = 5

Select Case KeyCode
Case 27 'Esc
WriteToFile
Unload Me

Case 32 'spacebar

If wavebreak Then
selectv = spotn: swapping = False: dblsngclickcommon
Else
dblsngclickcommon: clearcLicKaNYwheRe: waveform
 End If

Case 49 To 54 'number keys
 drawselect = KeyCode - 49
 swapping = False
 If styleSelectEnabled Then
 updateThoseTwoBars
 Form_Paint
  End If
  
'q or w or e is being pressed, increase color and slider
Case 81, 87, 69
pow = vVal(n)
If vVal(n) < 238 Then
vVal(n) = vVal(n) + 18
Else
vVal(n) = 255
 End If
updateCLibar

'a or s or d is being pressed, decrease color and slider
Case 65, 83, 68
pow = vVal(n)
If vVal(n) > 17 Then
vVal(n) = vVal(n) - 18
Else
vVal(n) = 0
 End If
updateCLibar

Case 67 'C key - clear
Dither Me: Form_Paint

Case 82 'R key - Randomize colors
selectv = 58: dblsngclickcommon: selectv = spotn: swapping = True
If Not wavebreak Then Call waveform

Case 84 'T key - Access temporary save
selectv = 57: dblsngclickcommon

Case 90 'Z key - anti-aliasing on/off
selectv = 15: dblsngclickcommon

Case vbKeyLeft 'load settings from previous bank
If spotn < 22 Then
spotn = 40
Else
spotn = spotn - 1
 End If: selectv = spotn
swapping = True: Call dblsngclickcommon
If Not wavebreak Then Dither Me: Form_Paint: Call waveform

Case vbKeyRight 'load settings from next bank
If spotn < 21 Or spotn = 40 Then
spotn = 21
Else
spotn = spotn + 1: End If: selectv = spotn
swapping = True: Call dblsngclickcommon
If Not wavebreak Then Dither Me: Form_Paint: Call waveform

Case vbKeyUp, vbKeyDown
If spotn > 20 And spotn < 31 Then
spotn = spotn + 10
ElseIf spotn > 20 Then
spotn = spotn - 10
 End If: selectv = spotn
swapping = True: Call dblsngclickcommon

'Shift key
Case 16
qRecall = False: drawStoreRecallbutton
 End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 16
qRecall = True
drawStoreRecallbutton
End Select
End Sub
Private Sub drawStoreRecallbutton()
 For n = vTop(57) + 1 To vBot(57) - 1 Step 1
 cL = GetPixel(Form1.hdc, 1, n)
 Form1.Line (vLeft(57) + 1, n)-(vRight(57), n), cL
  Next n
 If qRecall Then
 Form1.CurrentX = vLeft(57) + 7: Form1.CurrentY = vTop(57) + 4
 Form1.Print "Recall"
 Else
 Form1.CurrentX = vLeft(57) + 7: Form1.CurrentY = vTop(57) + 4
 Form1.Print " Store"
  End If
End Sub
Private Sub clearcLicKaNYwheRe()
 For n = iclearY To iclearY + 34 Step 1
  cL = GetPixel(Form1.hdc, 1, n)
  Form1.Line (iclearX, n)-(iclearX + 301, n), cL
 Next n
End Sub
Private Sub Dither(vForm As Form)
Dim h As Integer, q As Integer, ditr As Byte, ditg As Byte, ditb As Byte
Dim lcolor As Long

 Select Case bwi
 Case True: Form1.BackColor = vbWhite
 Case Else: Form1.BackColor = vbBlack
 End Select

 For h = 255 To 0 Step -1
  q = 741 - h * 3

  ditr = vVal(8) / 255 * h: ditg = vVal(9) / 255 * h: ditb = vVal(10) / 255 * h

  Select Case bwi
  Case 0
   lcolor = RGB(ditr, ditg, ditb): q = h * 3
  Case Else
   lcolor = RGB(255 - ditr, 255 - ditg, 255 - ditb): q = 741 - h * 3
  End Select

  vForm.Line (0, q)-(Form1.ScaleWidth, q), lcolor, B
  vForm.Line (0, q - 1)-(Form1.ScaleWidth, q - 1), lcolor, B
  vForm.Line (0, q - 2)-(Form1.ScaleWidth, q - 2), lcolor, B
 Next h
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 selectv = 0
 For n = 1 To 59
 Select Case X
 Case vLeft(n) To vRight(n)
 Select Case Y
 Case vTop(n) To vBot(n)
 selectv = n 'We've landed inside a control's dimensions
 yInit = 0
  End Select
   End Select
    Next n
 Select Case selectv
 Case 0, 7, 11, 14, 15, 21 To 58
  swapping = False: dblsngclickcommon
 Case 16
 If styleSelectEnabled Then
 For n = 0 To 5
 Select Case X
 Case vLeft(16) + 10 * n To vLeft(16) + 10 * (n + 1)
 drawselect = n
  End Select
   Next n
 updateThoseTwoBars
 Form_Paint
  End If
 Case Else
 Call Form_MouseMove(Button, Shift, X, Y)
  End Select
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rr As Byte, gg As Byte, bb As Byte

Select Case Y
Case Is <> yInit
yInit = Y: n = selectv
If twistenabled And n = 1 Or opaenabled And n = 2 _
Or n > 2 And n <> 59 Or n = 59 And opaenabled Then
Select Case n
Case 1 To 6, 12, 13, 18 To 20, 59
swapping = False: erasebar
nli = 0 'resets r,g,b and transparency
 End Select
    
Select Case Y
Case vTop(n) To vBot(n)
vVal(n) = vMax(n) * (vBot(n) - Y) / (vBot(n) - vTop(n))
If n = 17 Then ribbon = True
Case Is < vTop(n)
vVal(n) = vMax(n)
If n = 17 Then ribbon = False
Case Is > vBot(n)
vVal(n) = vMin(n)
If n = 17 Then ribbon = True
If n = 59 Then s1 = vVal(2) / 255
End Select
Select Case n
Case 1 To 6, 12, 13, 18 To 20, 59
updateCLibar
End Select
End If
  
  Select Case n
  Case 1
   If twistenabled Then
   Call noTwist
   rr = r: gg = g: bb = b: w = vVal(2)
   r = vVal(3): g = vVal(4): b = vVal(5)
   vVal(2) = 255
   Call linestest
   r = rr: g = gg: b = bb: vVal(2) = w
   End If
  Case 2
   s1 = vVal(2) / 255: s1a = s1
   oscopa = s1 * vVal(59) / 255
  Case 3, 4, 5
   r = vVal(3)
   g = vVal(4)
   b = vVal(5)
   yellowsquarepaint
  Case 6
   incb = vVal(6) / 6
   incr = vVal(18) * incb
   incg = vVal(19) * incb
   incb = vVal(20) * incb
  Case 12
   mstep = vVal(12) + 4
  Case 13
   int1 = vMax(13) * (vVal(13) / vMax(13)) ^ 2.5
  Case 17
   If vVal(17) < 1300 Then vVal(17) = 1300
  Case 18
   incr = vVal(18) * vVal(6) / 6
  Case 19
   incg = vVal(19) * vVal(6) / 6
  Case 20
   incb = vVal(20) * vVal(6) / 6
  Case 59
   If opaenabled Then oscopa = vVal(2) * vVal(59) / 65536
  End Select
 End Select
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case selectv
 Case 1
 If twistenabled Then BGpaste
 selectv = 0
 Case 8, 9, 10
 If bwi Then
 bg(selectv - 7) = vVal(selectv)
 Else
 lightbgchanged = 1
 bg(4) = vVal(8)
 bg(5) = vVal(9)
 bg(6) = vVal(10)
  End If
 selectv = 0
 clearsav
 Dither Me: Form_Paint
 Case 2 To 6, 12, 13, 17 To 20, 59
 selectv = 0
  End Select
 If Not wavebreak Then
 If Not selectv = 7 Then
 clearcLicKaNYwheRe
  End If
 Call waveform
   End If
End Sub
Private Sub Form_DblClick()
 dblsngclickcommon
End Sub

Private Sub clearsav()
If ribbon Then
For cL = 0 To vVal(17) Step 1
savcL(cL) = 0
savDS(cL) = 0
 Next cL
End If
ns1 = 0
End Sub
Private Sub erasebar()
 Select Case n
 Case 18, 19, 20
 bar = vBot(n) - (vBot(n) - vTop(n) - 4) * (Abs(vVal(n)) - vMin(n)) / (vMax(n) - vMin(n)) - 2
 Case Else
 bar = vBot(n) - (vBot(n) - vTop(n) - 4) * (vVal(n) - vMin(n)) / (vMax(n) - vMin(n)) - 2
 End Select
 cL = GetPixel(Form1.hdc, 1, bar)
 Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), cL
End Sub
Private Sub updateCLibar()
 bar = vBot(n) - (vBot(n) - vTop(n) - 4) * (pow - vMin(n)) / (vMax(n) - vMin(n)) - 2
 cL = GetPixel(Form1.hdc, 1, bar)
 Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), cL
 Select Case n
 Case 1, 2, 6, 18, 19, 20
 bar = vBot(n) - (vBot(n) - vTop(n) - 4) * (Abs(vVal(n)) - vMin(n)) / (vMax(n) - vMin(n)) - 2
 Case Else
 bar = vBot(n) - (vBot(n) - vTop(n) - 4) * (vVal(n) - vMin(n)) / (vMax(n) - vMin(n)) - 2
 r = vVal(3): g = vVal(4): b = vVal(5)
 If n = 3 Or n = 4 Or n = 5 Then yellowsquarepaint
 End Select
 Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), barcol
End Sub
Private Sub updateThoseTwoBars()
 Select Case drawselect
 Case 0
 opaenabled = True: twistenabled = True
 Case 1
 opaenabled = True: twistenabled = False
 shape = 255
 Case 2
 opaenabled = False: twistenabled = True
 Case 3
 opaenabled = True: twistenabled = True
 Case 4
 opaenabled = True: twistenabled = True
 Case 5
 opaenabled = True: twistenabled = True
  End Select
 n = 1: erasebar
 If twistenabled Then
  Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), barcol
 End If
 n = 2: erasebar
 If opaenabled Then
  Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), barcol
 End If
 n = 59: erasebar
 If opaenabled Then
  Form1.Line (vLeft(n) + 1, bar)-(vRight(n), bar), barcol
 End If

 Call noTwist
End Sub

Private Sub antialias(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
Dim spx As Single, epx As Single
Dim spy As Single, epy As Single
Dim ax As Double, bx As Double, cx As Double, dx As Double
Dim ay As Double, by As Double, cy As Double, dy As Double
Dim ex As Double, ey As Double
Dim mp5 As Single, pp5 As Single
Dim rex As Integer, rey As Integer

Dim trz As Single, tri As Double
Dim lwris As Double, lwrun As Double
Dim zsl As Single
Dim slope As Double, lsope As Double
Dim midx As Double, midy As Double
Dim sl2 As Single, ris As Double, run As Double
Dim distanc1 As Double, distanc2 As Double
Dim diagonal As Boolean
Dim a As Single
Dim st As Integer
Dim one As Single
Static sr(0 To 128) As Single
Static firstrun As Boolean

If Not firstrun Then 'initialize some things
 sr(0) = 0: sr(128) = 0.5
 For w = 1 To 127
 sr(w) = (w / 128) ^ 2 / 2
 Next w

'Initialize An array which maps a curve of 256 Single values
pow = 1 / (1 + shape / 255)
For w = 1 To 254
gs(w) = (w / 255) ^ pow
Next w: gs(0) = 0: gs(255) = 1
firstrun = 1: End If

If x1 < x2 Then
 epy = y2: epx = x2
 spy = y1: spx = x1
Else
spy = y2: spx = x2
epy = y1: epx = x1: End If

ris = epy - spy: run = epx - spx

If epx = spx Or epy = spy Then
 diagonal = 0
 If epy > spy Then
  st = 1
 Else: st = -1: End If

Else: slope = ris / run: lsope = -run / ris
diagonal = 1: End If

midx = 0.5 * lsope: midy = 0.5 * slope
sl2 = slope ^ 2: one = 1 / Sqr(ris ^ 2 + run ^ 2)
lwris = 1 - (one * run - Abs(slope) + slope * one * ris)
lwrun = lwris * Abs(lsope)
distanc1 = 0.5 / Sqr(1 + sl2)
distanc2 = slope * distanc1
ax = spx - distanc1 - distanc2
ay = spy + distanc1 - distanc2
bx = epx + distanc1 - distanc2
by = epy + distanc1 + distanc2
cx = epx + distanc1 + distanc2
cy = epy - distanc1 + distanc2
dx = spx - distanc1 + distanc2
dy = spy - distanc1 - distanc2

one = 255 * (1 - 0.5 * lwris * lwrun)

If diagonal Then
If slope > 0 Then
If slope <= 1 Then
ey# = slope# * (Round(ax#) + 1.5 - ax#) + ay#
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1
tri# = pp5! - lwris#: trz! = pp5! - slope#: zsl! = mp5! - midy
For ex# = Round(ax#) + 1.5 To Round(bx#) - 1.5 Step 1
 If ey# > tri# Then
  Call pixels(Int(ex# + 1), rey%, gs(one!))
 Else
  If ey# > trz! Then
  Call pixels(Int(ex# + 1), rey%, gs(255 * (1 + lsope# * sr(Int((pp5! - ey#) * 128)))))
  Else: Call pixels(Int(ex# + 1), rey%, gs(255 * (ey# - zsl!))): End If
 End If
If ey# > trz! Then
 pp5! = pp5! + 1: tri# = tri# + 1: trz! = trz! + 1
 zsl! = zsl! + 1: rey% = rey% + 1: End If
ey = ey# + slope: Next ex#

ey# = cy# - slope# * (cx# - Round(cx#) + 1.5)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! + midy
tri# = mp5! + lwris: trz! = mp5! + slope
For ex# = Round(cx#) - 1.5 To Round(dx#) + 1.5 Step -1
If ey# > tri# Then
 If ey# > trz! Then
 Call pixels(Int(ex#), rey%, gs(255 * (zsl! - ey#)))
 Else: Call pixels(Int(ex#), rey%, gs(255 * (1 + lsope# * sr(Int((ey# - mp5!) * 128)))))
 End If
End If
If ey# < trz! Then
 mp5! = mp5! - 1: tri# = tri# - 1: trz! = trz! - 1
 zsl! = zsl! - 1: rey% = rey% - 1: End If
ey# = ey# - slope: Next ex#

ex# = cx# + lsope * (cy# - Round(cy#) + 0.5)
For ey# = Round(cy#) - 0.5 To Round(dy#) + 1.5 Step -1
 rex% = Round(ex#)
 Call pixels(rex%, Int(ey#), gs(255 * (slope# * sr(Int((ex# - rex% + 0.5) * 128)))))
ex# = ex# + lsope#: Next ey#

ex# = ax# - lsope# * (Round(ay#) + 0.5 - ay#)
For ey# = Round(ay) + 0.5 To Round(by) - 0.5 Step 1
 rex% = Round(ex#)
 Call pixels(rex%, Int(ey# + 1), gs(255 * (slope# * sr(Int((rex% + 0.5 - ex#) * 128)))))
ex# = ex# - lsope#: Next ey#

Else
ex# = dx# - lsope# * (Round(dy#) + 1.5 - dy#): rex% = Round(ex#): mp5! = rex% - 0.5
pp5! = mp5! + 1: tri# = pp5! - lwrun: trz! = pp5! + lsope: zsl! = mp5! + midx
For ey# = Round(dy#) + 1.5 To Round(cy#) - 0.5 Step 1
 If ex# > tri# Then
  Call pixels(rex%, Int(ey# + 1), gs(one!))
 Else
  If ex# > trz! Then
   Call pixels(rex%, Int(ey# + 1), gs(255 * (1 - slope * sr(Int((pp5! - ex#) * 128)))))
  Else: Call pixels(rex%, Int(ey# + 1), gs(255 * (ex# - zsl!))): End If
 End If
If ex# > trz! Then
 tri# = tri# + 1: trz! = trz! + 1: pp5! = pp5! + 1
 zsl! = zsl! + 1: rex% = rex% + 1: End If
ex# = ex# - lsope#: Next ey#

ex# = bx# + lsope# * (by# - Round(by#) + 0.5): rex% = Round(ex#): mp5! = rex% - 0.5
pp5! = mp5! + 1: tri# = mp5! + lwrun: trz! = mp5! - lsope: zsl! = pp5! - midx#
For ey# = Round(by#) - 0.5 To Round(ay#) + 1.5 Step -1
If ex# > tri# Then
 If ex# < trz! Then
  Call pixels(rex%, Int(ey#), gs(255 * (1 - slope# * sr(Int((ex# - mp5!) * 128)))))
 Else: Call pixels(rex%, Int(ey#), gs(255 * (zsl! - ex#))): End If
End If
If ex# < trz! Then
 tri# = tri# - 1: trz! = trz! - 1: mp5! = mp5! - 1
 rex% = rex% - 1: zsl! = zsl! - 1: End If
ex# = ex# + lsope#: Next ey#

ey# = dy# + slope# * (Round(dx#) + 0.5 - dx#)
For ex# = Round(dx#) + 0.5 To Round(cx#) - 1.5 Step 1
 rey% = Round(ey#)
 Call pixels(Int(ex# + 1), rey%, gs(255 * (-lsope# * sr(Int((rey% + 0.5 - ey#) * 128)))))
ey# = ey# + slope#: Next ex#

ey# = ay# + slope# * (Round(ax#) + 1.5 - ax#)
For ex# = Round(ax#) + 1.5 To Round(bx#) - 0.5 Step 1
 rey% = Round(ey#)
 Call pixels(Int(ex#), rey%, gs(255 * (-lsope# * sr(Int((ey# - rey% + 0.5) * 128)))))
ey# = ey# + slope: Next ex#
End If

Else
If slope > -1 Then
ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! - midy#
tri# = mp5! + lwris: trz! = mp5! - slope#
For ex# = Round(dx#) + 1.5 To Round(cx#) - 1.5 Step 1
 If ey# < tri# Then
  Call pixels(Int(ex# + 1), rey%, gs(one!))
 Else
  If ey# < trz! Then
   Call pixels(Int(ex# + 1), rey%, gs(255 * (1 - lsope# * sr(Int((ey# - mp5!) * 128)))))
  Else: Call pixels(Int(ex# + 1), rey%, gs(255 * (zsl! - ey#))): End If
 End If
If ey# < trz! Then
  mp5! = mp5! - 1: zsl! = zsl! - 1: tri# = tri# - 1
  rey% = rey% - 1: trz! = trz! - 1: End If
ey# = ey# + slope#: Next ex#

ey# = by# - slope# * (bx# - Round(bx#) + 1.5)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = mp5! + midy#
tri# = pp5! - lwris: trz! = pp5! + slope#
For ex# = Round(bx#) - 1.5 To Round(ax#) + 1.5 Step -1
If ey# < tri# Then
 If ey# < trz! Then
  Call pixels(Int(ex#), rey%, gs(255 * (ey# - zsl!)))
 Else
  Call pixels(Int(ex#), rey%, gs(255 * (1 - lsope# * sr(Int((pp5! - ey#) * 128)))))
 End If
End If
If ey# > trz! Then
  rey% = rey% + 1: pp5! = pp5! + 1: zsl! = zsl! + 1
  tri# = tri# + 1: trz! = trz! + 1: End If
ey# = ey# - slope#: Next ex#

ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5)
For ey# = Round(ay#) - 1.5 To Round(by#) + 0.5 Step -1
 rex% = Round(ex#)
 Call pixels(rex%, Int(ey# + 1), gs(255 * (-slope * sr(Int((ex# - rex% + 0.5) * 128)))))
ex# = ex# + 2 * midx#: Next ey#

ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#)
For ey# = Round(cy#) + 1.5 To Round(dy#) - 0.5
 rex% = Round(ex#)
 Call pixels(rex%, Int(ey#), gs(255 * (-slope# * sr(Int((rex% + 0.5 - ex#) * 128)))))
ex# = ex# - 2 * midx#: Next ey#

Else
ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5): rex% = Round(ex#): mp5! = rex% - 0.5
zsl! = mp5! - midx#: pp5! = mp5! + 1: tri# = pp5! - lwrun#: trz! = pp5! - lsope#
For ey# = Round(ay#) - 1.5 To Round(by#) + 1.5 Step -1
 If ex# > tri# Then
  Call pixels(rex%, Int(ey#), gs(one!))
 Else
  If ex# > trz! Then
   Call pixels(rex%, Int(ey#), gs(255 * (1 + slope# * sr(Int((pp5! - ex#) * 128)))))
  Else: Call pixels(rex%, Int(ey#), gs(255 * (ex# - zsl!))): End If
 End If
If ex# > trz! Then
  tri# = tri# + 1: trz! = trz! + 1: pp5! = pp5! + 1
  rex% = rex% + 1: zsl! = zsl! + 1: End If
ex# = ex# + lsope#: Next ey#

ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#): rex% = Round(ex#): mp5! = rex% - 0.5
tri# = mp5! + lwrun#: trz! = mp5! + lsope#: zsl! = mp5! + 1 + midx#
For ey# = Round(cy#) + 1.5 To Round(dy#) - 1.5 Step 1
 If ex# > tri# Then
  If ex# < trz! Then
   Call pixels(rex%, Int(ey# + 1), gs(255 * (1 + slope * sr(Int((ex# - mp5!) * 128)))))
  Else: Call pixels(rex%, Int(ey# + 1), gs(255 * (zsl! - ex#))): End If
 End If
If ex# < trz! Then
 tri# = tri# - 1: trz! = trz! - 1: zsl! = zsl! - 1
 rex% = rex% - 1: mp5! = mp5! - 1: End If
ex# = ex# - lsope#: Next ey#

ey# = by# - slope# * (bx# - Round(bx#) + 2.5)
For ex# = Round(bx#) - 2.5 To Round(ax#) + 0.5 Step -1
 rey% = Round(ey#)
 Call pixels(Int(ex# + 1), rey%, gs(255 * (lsope# * sr(Int((ey# - rey% + 0.5) * 128)))))
ey# = ey# - slope#: Next ex#

ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
For ex# = Round(dx#) + 1.5 To Round(cx#) - 0.5 Step 1
 rey% = Round(ey#)
 Call pixels(Int(ex#), rey%, gs(255 * (lsope# * sr(Int((rey% + 0.5 - ey#) * 128)))))
ey# = ey# + slope#: Next ex#

End If
End If

Else
If epy! = spy! Then
 rey% = Round(ay#): a! = ay# - rey% + 0.5
For ex# = Round(ax#) + 1.5 To Round(bx#) - 0.5
 Call pixels(Int(ex#), rey%, gs(255 * a!))
 Call pixels(Int(ex#), rey% - 1, 1 - gs(255 * a!))
Next ex#
Else
If epx! = spx! Then
 rex% = Round(ax#): a! = ax# - rex% + 0.5
For ey# = Round(ay#) - 0.5 To Round(by) + 1.5 Step st
 Call pixels(rex%, Int(ey#), gs(255 * a!))
 Call pixels(rex% - 1, Int(ey#), 1 - gs(255 * a!))
Next ey#: End If
End If
End If
End Sub
Private Sub pixels(X As Integer, Y As Integer, a As Single)
Dim r2 As Long, g2 As Long, b2 As Long, s2 As Single
Dim BGR As Long

Select Case drawselect
Case Is <> 2
 BGR& = GetPixel(Form1.hdc, X, Y)
 b2& = ((BGR& And &HFF0000) / &H10000) And &HFF
 g2& = ((BGR& And &HFF00) / &H100) And &HFF
 r2& = BGR& And &HFF
 s2! = a! * s1
End Select

Select Case drawselect

Case 0 ' "Normal"
 cL = RGB(r2& - s2! * (r2& - r!), g2& - s2! * (g2& - g!), b2& - s2! * (b2& - b!))

Case 1 ' "Paint"
 cL = RGB(r2& - s2! * (r2& - a! * r!), g2& - s2! * (g2& - a! * g!), b2& - s2! * (b2& - a! * b!))

Case 2 ' "Pixelsticks"
 cL = RGB(r! * a!, g! * a!, b! * a!)
 'BGR = GetPixel(Form1.hdc, X, Y)

Case 3 ' "Laser Smoke"
 cL = RGB(r2& + s2! * Abs(r2& - r!) ^ 0.85, g2& + s2! * Abs(g2& - g!) ^ 0.85, b2& + s2! * Abs(b2& - b!) ^ 0.85)
 
Case 4 ' "Lighten"
 If r2& < r! Then
 r2& = r2& + s2! * (r! - r2&): End If
 If g2& < g! Then
 g2& = g2& + s2! * (g! - g2&): End If
 If b2& < b! Then
 b2& = b2& + s2! * (b! - b2&): End If
 cL = RGB(r2&, g2&, b2&)

Case 5 ' "Phasine"
 cL = RGB(r2& - gs(r! * s2!) * (r2& - r!) * (vVal(2) / 255), g2& - gs(g! * s2!) * (g2& - g!) * (vVal(2) / 255), b2& - gs(b! * s2!) * (b2& - b!) * (vVal(2) / 255))
End Select

SetPixelV Form1.hdc, X, Y, cL

Select Case ribbon
Case True
Select Case nottesting
Case True
savX(ns1) = X: savy(ns1) = Y

Select Case drawselect
Case Is <> 2: savcL(ns1) = BGR - cL: savDS(ns1) = 0
Case 2: savDS(ns1) = 1: savcL(ns1) = BGR: End Select

Select Case ns1
Case 0: ns1 = 1: n2 = 1
Case Is >= vVal(17): ns1 = 0: n2 = 0
Case Else: ns1 = ns1 + 1: n2 = ns1: End Select

Select Case savDS(n2)
Case 0
BGR = GetPixel(Form1.hdc, savX(n2), savy(n2))
SetPixelV Form1.hdc, savX(n2), savy(n2), BGR + savcL(n2)
Case 1
SetPixelV Form1.hdc, savX(n2), savy(n2), savcL(n2)
 End Select
  End Select
   End Select
End Sub
Public Sub waveform()
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim z As Single: Static firstrun As Boolean
Dim mult1 As Single, mult2 As Single
Static sw As Single, sh As Single

 sw = 9 + Form1.ScaleWidth / 2: sh = Form1.ScaleHeight / 2
 mult1 = 5 * (Rnd - 0.5)
 mult2 = 5 * (Rnd - 0.5)

 If Not wavebreak Then
  breakloop = False
  mstep = vVal(12) + 4: int1 = vVal(13): wavebreak = True
 End If

 If ClearMe Then
  Dither Me: Form_Paint
  ClearMe = False
 End If

 Do While Not breakloop
  DoEvents

  If movinglines Then
   x1 = sw + 115 * Sin(mult2 + z / 150) + 125 * Cos(mult2 + z / 190)
   x2 = sw + 115 * Sin(z * mult2 / 180) + 120 * Sin(mult2 + z / 230)
   y1 = sh + 110 * Sin(mult1 + z / 315)
   y2 = sh + 200 * Sin(4 * mult1 + z / 249)

   If Not aliasing Then
    Call antialias(x1, y1, x2, y2)
   Else
    Form1.Line (x1, y1)-(x2, y2), RGB(r, g, b)
    For x1 = 1 To 4450         'a pause loop
    y1 = x2 * Sin(50): Next x1 'heavy math should slow it down
   End If
   
   z = z + mstep
   Select Case z
   Case Is > int1
    z = 1
    mult1 = 5 * (Rnd - 0.5)
    mult2 = 5 * (Rnd - 0.5)
    If randfirst Then r = Rnd * 255: g = Rnd * 255: b = Rnd * 255: nli = 0
   End Select
   Call colors
  End If
 Loop

 Unload Me
End Sub
Private Sub colors()
 s1 = s1 + oscopa
 If s1 > s1a Then
 s1 = s1a: oscopa = -oscopa
 ElseIf s1 < 0 Then
 s1 = 0: oscopa = -oscopa
 End If
 
 r = r + incr
 If r < vval2(3) Then
 incr = -incr: r = vval2(3)
 ElseIf r > 255 Then
  incr = -incr: r = 255: End If
 g = g + incg
 If g < vval2(4) Then
 incg = -incg: g = vval2(4)
 ElseIf g > 255 Then
  incg = -incg: g = 255: End If
 b = b + incb
 If b < vval2(5) Then
 incb = -incb: b = vval2(5)
 ElseIf b > 255 Then
  incb = -incb: b = 255: End If
End Sub
Private Sub Form_LostFocus()
 clearsav
End Sub
Private Sub Form_Unload(Cancel As Integer)
 WriteToFile
 breakloop = True
End Sub

Private Sub yellowsquarepaint()
 If randfirst Then
  For n = vTop(14) + 2 To vBot(14) - 1 Step 1
   Form1.Line (vLeft(14) + 1, n)-(vRight(14) - 1, n), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
  Next n
  Form1.Line (vLeft(14), vTop(14) + 1)-(vRight(14) - 1, vBot(14)), RGB(vVal(3), vVal(4), vVal(5)), B
 Else
  Form1.Line (vLeft(14), vTop(14) + 1)-(vRight(14) - 1, vBot(14)), RGB(vVal(3), vVal(4), vVal(5)), BF
 End If
End Sub

Private Sub BGcopy()
Dim x1 As Integer, y1 As Integer
Dim x2 As Integer

 Select Case stuffed
 Case False
  n = 1
  For y1 = 99 To 301 Step 1 'Save background data in this specific area
   x2 = Int((y1 - 99) / 20)
   For x1 = 77 - x2 To 79 - x2 Step 1
    sr(n) = GetPixel(Form1.hdc, x1, y1)
    n = n + 1
   Next x1
  Next y1
  stuffed = True
 End Select
 
End Sub
Private Sub BGpaste()
Dim x1 As Integer, y1 As Integer
Dim x2 As Integer, y2 As Integer

 n = 1
 For y1 = 99 To 301 Step 1
  x2 = Int((y1 - 99) / 20)
  For x1 = 77 - x2 To 79 - x2 Step 1
   SetPixelV Form1.hdc, x1, y1, sr(n)
   n = n + 1
  Next x1
 Next y1
 stuffed = False
 
End Sub
Private Sub linestest()
Dim rr As Byte, gg As Byte, bb As Byte

 If Not stuffed Then
  Call BGcopy
 End If

 Select Case bwi
 Case 0: rr = 255 - vVal(8): gg = 255 - vVal(9): bb = 255 - vVal(10)
 Case Else: rr = vVal(8): gg = vVal(9): bb = vVal(10): End Select

 BGpaste
 stuffed = 1
 
 rr = r: gg = g: bb = b: r = vVal(3): g = vVal(4): b = vVal(5)
 nottesting = False
 pow = s1 'store line intensity from last drawn
 s1 = 1 'line at full intensity drawn off to side
 Call antialias(78, 100, 68, 300)
 s1 = pow 'resume
 nottesting = True
 r = rr: g = gg: b = bb
End Sub
Private Sub noTwist()
 Select Case twistenabled
 Case True
  shape = vVal(1)
 End Select

 pow = 1 / (1 + shape / 255)
 For w = 1 To 254
  gs(w) = (w / 255) ^ pow
 Next w: gs(0) = 0: gs(255) = 1
End Sub
Private Sub WriteToFile()
 Dim FreeFileNum As Integer
 FreeFileNum = FreeFile
 Open "colorsav1.txt" For Output As #FreeFileNum
 For n = 21 To 40
 Print #FreeFileNum, svr(n), svg(n), svb(n), svri(n), svgi(n), svbi(n), _
 svds(n), svo(n), svoa(n)
 Next n
 Print #FreeFileNum, vVal(3), vVal(4), vVal(5), vVal(18), vVal(19), vVal(20), _
 drawselect, vVal(2), vVal(59)
 Print #FreeFileNum, bg(1), bg(2), bg(3), bg(4), bg(5), bg(6), bwi
Close #FreeFileNum
End Sub


' Option Explicit

'fluoats@hotmail.com
'TO GET THIS CODE TO WORK, Replace all "' " with ""
'That means replace all <comment><space> with <nothing>

' 'Form_Load will tell you most of what you need to know


' 'To implement in your own program without all the fancy flying around,
' 'Totally Necessary variables: (under Option Explicit)
' Dim gs(0 To 255) As Single
' Dim shape As Byte, pow As Single
' Dim drawselect As Byte
' Dim s1 as single 's1 must be between 0 and 1
' Dim r as single, g as single, b as single

' Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
' Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

' 'Totally necessary subs:
' 'antialias() and pixels()
' 'and then in Form_Load or elsewhere,
' 'make sure s1 is between 0 and 1
' 'drawselect can be 0,1,2,3,4, or 5

' 'Otherwise, here's all the non-essential variables ..
' 'things like fancy flight data, r,g,b shift intensity

' Dim bwi As Boolean
' Dim backdred As Byte, backdgrn As Byte, backdblu As Byte
' Dim incr As Single, incg As Single, incb As Single
' Dim mstep As Integer, int1 As Single
' Dim n As Integer, n2 As Integer 'all-purpose
' Dim gbye As Boolean 'used to exit in sub waveform()
'Rather not explain  :P

' Private Sub Form_Load()
'  Form1.ScaleMode = 3: Form1.AutoRedraw = False
'  Form1.Width = 10420: Form1.Height = 8192
'  Randomize

'  drawselect = 2  '0, 1, 2, 3, 4 or 5 in this version (2.12)
'  s1 = 100 / 255  '0 = totally transparent w/ drawselect 0,1,2,3,5
'  shape = 60      '255 = max "easfw-enhanced"  anti-aliasing

'  bwi = 1         'make this 0 for inverted background
'  backdred = 200  'background dither colors
'  backdgrn = 200
'  backdblu = 200
' End Sub
' Private Sub Form_Activate()
'  Call waveform
' End Sub


' Private Sub antialias(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
' Dim spx As Single, epx As Single
' Dim spy As Single, epy As Single
' Dim ax As Double, bx As Double, cx As Double, dx As Double
' Dim ay As Double, by As Double, cy As Double, dy As Double
' Dim ex As Double, ey As Double
' Dim mp5 As Single, pp5 As Single
' Dim rex As Integer, rey As Integer

' Dim trz As Double, tri As Double
' Dim lwris As Double, lwrun As Double
' Dim zsl As Single
' Dim slope As Double, lsope As Double
' Dim midx As Double, midy As Double
' Dim sl2 As Single
' Dim distanc1 As Double, distanc2 As Double
' Dim diagonal As Boolean
' Dim a As Single
' Dim st As Integer
' Dim one As Single
' Static sr(0 To 128) As Single
' Static firstrun As Boolean
' Dim w As Byte

' If Not firstrun Then 'initialize some things
'  sr(0) = 0: sr(128) = 0.5
'  For w = 1 To 127
'  sr(w) = (w / 128) ^ 2 / 2
'  Next w

'Initialize An array which maps a curve of 256 Single values
' pow = 1 / (1 + shape / 255)
' For w = 1 To 254
' gs(w) = (w / 255) ^ pow
' Next w: gs(0) = 0: gs(255) = 1
' firstrun = 1: End If

' If x1 < x2 Then
'  epy = y2: epx = x2
'  spy = y1: spx = x1
' Else
' spy = y2: spx = x2
' epy = y1: epx = x1: End If

' If epx = spx Or epy = spy Then
'  diagonal = 0
'  If epy > spy Then
'   st = 1
'  Else: st = -1: End If

' Else: slope = (epy - spy) / (epx - spx): lsope = -1 / slope
' diagonal = 1: End If

' midx = 0.5 * lsope: midy = 0.5 * slope
' sl2 = slope * slope: one = 1 / Sqr(1 + sl2)
' lwris = 1 - (one - Abs(slope) + sl2 * one)
' lwrun = lwris * Abs(lsope)
' distanc1 = 0.5 * one
' distanc2 = slope * distanc1
' ax = spx - distanc1 - distanc2
' ay = spy + distanc1 - distanc2
' bx = epx + distanc1 - distanc2
' by = epy + distanc1 + distanc2
' cx = epx + distanc1 + distanc2
' cy = epy - distanc1 + distanc2
' dx = spx - distanc1 + distanc2
' dy = spy - distanc1 - distanc2
' one = 255 * (1 - 0.5 * lwris * lwrun)

' If diagonal Then
' If slope > 0 Then
' If slope <= 1 Then
' ey# = slope# * (Round(ax#) + 1.5 - ax#) + ay#
' rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1
' tri# = pp5! - lwris#: trz# = pp5! - slope#: zsl! = mp5! - midy
' For ex# = Round(ax#) + 1.5 To Round(bx#) - 1.5 Step 1
'  If ey# > tri# Then
'   Call pixels(Int(ex# + 1), rey%, gs(one!))
'  Else
'   If ey# > trz# Then
'   Call pixels(Int(ex# + 1), rey%, gs(255 * (1 + lsope# * sr(Int((pp5! - ey#) * 128)))))
'   Else: Call pixels(Int(ex# + 1), rey%, gs(255 * (ey# - zsl!))): End If
'  End If
' If ey# > trz# Then
'  pp5! = pp5! + 1: tri# = tri# + 1: trz# = trz# + 1
'  zsl! = zsl! + 1: rey% = rey% + 1: End If
' ey = ey# + slope: Next ex#
' ey# = cy# - slope# * (cx# - Round(cx#) + 1.5)
' rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! + midy
' tri# = mp5! + lwris: trz# = mp5! + slope
' For ex# = Round(cx#) - 1.5 To Round(dx#) + 1.5 Step -1
' If ey# > tri# Then
'  If ey# > trz# Then
'  Call pixels(Int(ex#), rey%, gs(255 * (zsl! - ey#)))
'  Else: Call pixels(Int(ex#), rey%, gs(255 * (1 + lsope# * sr(Int((ey# - mp5!) * 128)))))
'  End If
' End If
' If ey# < trz# Then
'  mp5! = mp5! - 1: tri# = tri# - 1: trz# = trz# - 1
'  zsl! = zsl! - 1: rey% = rey% - 1: End If
' ey# = ey# - slope: Next ex#
' ex# = cx# + lsope * (cy# - Round(cy#) + 0.5)
' For ey# = Round(cy#) - 0.5 To Round(dy#) + 1.5 Step -1
'  rex% = Round(ex#)
'  Call pixels(rex%, Int(ey#), gs(255 * (slope# * sr(Int((ex# - rex% + 0.5) * 128)))))
' ex# = ex# + lsope#: Next ey#
' ex# = ax# - lsope# * (Round(ay#) + 0.5 - ay#)
' For ey# = Round(ay) + 0.5 To Round(by) - 0.5 Step 1
'  rex% = Round(ex#)
'  Call pixels(rex%, Int(ey# + 1), gs(255 * (slope# * sr(Int((rex% + 0.5 - ex#) * 128)))))
' ex# = ex# - lsope#: Next ey#
' Else
' ex# = dx# - lsope# * (Round(dy#) + 1.5 - dy#): rex% = Round(ex#): mp5! = rex% - 0.5
' pp5! = mp5! + 1: tri# = pp5! - lwrun: trz# = pp5! + lsope: zsl! = mp5! + midx
' For ey# = Round(dy#) + 1.5 To Round(cy#) - 0.5 Step 1
'  If ex# > tri# Then
'   Call pixels(rex%, Int(ey# + 1), gs(one!))
'  Else
'   If ex# > trz# Then
'    Call pixels(rex%, Int(ey# + 1), gs(255 * (1 - slope * sr(Int((pp5! - ex#) * 128)))))
'   Else: Call pixels(rex%, Int(ey# + 1), gs(255 * (ex# - zsl!))): End If
'  End If
' If ex# > trz# Then
'  tri# = tri# + 1: trz# = trz# + 1: pp5! = pp5! + 1
'  zsl! = zsl! + 1: rex% = rex% + 1: End If
' ex# = ex# - lsope#: Next ey#
' ex# = bx# + lsope# * (by# - Round(by#) + 0.5): rex% = Round(ex#): mp5! = rex% - 0.5
' pp5! = mp5! + 1: tri# = mp5! + lwrun: trz# = mp5! - lsope: zsl! = pp5! - midx#
' For ey# = Round(by#) - 0.5 To Round(ay#) + 1.5 Step -1
' If ex# > tri# Then
'  If ex# < trz# Then
'   Call pixels(rex%, Int(ey#), gs(255 * (1 - slope# * sr(Int((ex# - mp5!) * 128)))))
'  Else: Call pixels(rex%, Int(ey#), gs(255 * (zsl! - ex#))): End If
' End If
' If ex# < trz# Then
'  tri# = tri# - 1: trz# = trz# - 1: mp5! = mp5! - 1
'  rex% = rex% - 1: zsl! = zsl! - 1: End If
' ex# = ex# + lsope#: Next ey#
' ey# = dy# + slope# * (Round(dx#) + 0.5 - dx#)
' For ex# = Round(dx#) + 0.5 To Round(cx#) - 1.5 Step 1
'  rey% = Round(ey#)
'  Call pixels(Int(ex# + 1), rey%, gs(255 * (-lsope# * sr(Int((rey% + 0.5 - ey#) * 128)))))
' ey# = ey# + slope#: Next ex#
' ey# = ay# + slope# * (Round(ax#) + 1.5 - ax#)
' For ex# = Round(ax#) + 1.5 To Round(bx#) - 0.5 Step 1
'  rey% = Round(ey#)
'  Call pixels(Int(ex#), rey%, gs(255 * (-lsope# * sr(Int((ey# - rey% + 0.5) * 128)))))
' ey# = ey# + slope: Next ex#
' End If
' Else
' If slope > -1 Then
' ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
' rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! - midy#
' tri# = mp5! + lwris: trz# = mp5! - slope#
' For ex# = Round(dx#) + 1.5 To Round(cx#) - 1.5 Step 1
'  If ey# < tri# Then
'   Call pixels(Int(ex# + 1), rey%, gs(one!))
'  Else
'   If ey# < trz# Then
'    Call pixels(Int(ex# + 1), rey%, gs(255 * (1 - lsope# * sr(Int((ey# - mp5!) * 128)))))
'   Else: Call pixels(Int(ex# + 1), rey%, gs(255 * (zsl! - ey#))): End If
'  End If
' If ey# < trz# Then
'   mp5! = mp5! - 1: zsl! = zsl! - 1: tri# = tri# - 1
'   rey% = rey% - 1: trz# = trz# - 1: End If
' ey# = ey# + slope#: Next ex#
' ey# = by# - slope# * (bx# - Round(bx#) + 1.5)
' rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = mp5! + midy#
' tri# = pp5! - lwris: trz# = pp5! + slope#
' For ex# = Round(bx#) - 1.5 To Round(ax#) + 1.5 Step -1
' If ey# < tri# Then
'  If ey# < trz# Then
'   Call pixels(Int(ex#), rey%, gs(255 * (ey# - zsl!)))
'  Else
'   Call pixels(Int(ex#), rey%, gs(255 * (1 - lsope# * sr(Int((pp5! - ey#) * 128)))))
'  End If
' End If
' If ey# > trz# Then
'   rey% = rey% + 1: pp5! = pp5! + 1: zsl! = zsl! + 1
'   tri# = tri# + 1: trz# = trz# + 1: End If
' ey# = ey# - slope#: Next ex#
' ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5)
' For ey# = Round(ay#) - 1.5 To Round(by#) + 0.5 Step -1
'  rex% = Round(ex#)
'  Call pixels(rex%, Int(ey# + 1), gs(255 * (-slope * sr(Int((ex# - rex% + 0.5) * 128)))))
' ex# = ex# + 2 * midx#: Next ey#
' ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#)
' For ey# = Round(cy#) + 1.5 To Round(dy#) - 0.5
'  rex% = Round(ex#)
'  Call pixels(rex%, Int(ey#), gs(255 * (-slope# * sr(Int((rex% + 0.5 - ex#) * 128)))))
' ex# = ex# - 2 * midx#: Next ey#
' Else
' ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5): rex% = Round(ex#): mp5! = rex% - 0.5
' zsl! = mp5! - midx#: pp5! = mp5! + 1: tri# = pp5! - lwrun#: trz# = pp5! - lsope#
' For ey# = Round(ay#) - 1.5 To Round(by#) + 1.5 Step -1
'  If ex# > tri# Then
'   Call pixels(rex%, Int(ey#), gs(one!))
'  Else
'   If ex# > trz# Then
'    Call pixels(rex%, Int(ey#), gs(255 * (1 + slope# * sr(Int((pp5! - ex#) * 128)))))
'   Else: Call pixels(rex%, Int(ey#), gs(255 * (ex# - zsl!))): End If
'  End If
' If ex# > trz# Then
'   tri# = tri# + 1: trz# = trz# + 1: pp5! = pp5! + 1
'   rex% = rex% + 1: zsl! = zsl! + 1: End If
' ex# = ex# + lsope#: Next ey#
' ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#): rex% = Round(ex#): mp5! = rex% - 0.5
' tri# = mp5! + lwrun#: trz# = mp5! + lsope#: zsl! = mp5! + 1 + midx#
' For ey# = Round(cy#) + 1.5 To Round(dy#) - 1.5 Step 1
'  If ex# > tri# Then
'   If ex# < trz# Then
'    Call pixels(rex%, Int(ey# + 1), gs(255 * (1 + slope * sr(Int((ex# - mp5!) * 128)))))
'   Else: Call pixels(rex%, Int(ey# + 1), gs(255 * (zsl! - ex#))): End If
'  End If
' If ex# < trz# Then
'  tri# = tri# - 1: trz# = trz# - 1: zsl! = zsl! - 1
'  rex% = rex% - 1: mp5! = mp5! - 1: End If
' ex# = ex# - lsope#: Next ey#
' ey# = by# - slope# * (bx# - Round(bx#) + 2.5)
' For ex# = Round(bx#) - 2.5 To Round(ax#) + 0.5 Step -1
'  rey% = Round(ey#)
'  Call pixels(Int(ex# + 1), rey%, gs(255 * (lsope# * sr(Int((ey# - rey% + 0.5) * 128)))))
' ey# = ey# - slope#: Next ex#
' ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
' For ex# = Round(dx#) + 1.5 To Round(cx#) - 0.5 Step 1
'  rey% = Round(ey#)
'  Call pixels(Int(ex#), rey%, gs(255 * (lsope# * sr(Int((rey% + 0.5 - ey#) * 128)))))
' ey# = ey# + slope#: Next ex#
' End If
' End If
' Else
' If epy! = spy! Then
'  rey% = Round(ay#): a! = ay# - rey% + 0.5
' For ex# = Round(ax#) + 1.5 To Round(bx#) - 0.5
'  Call pixels(Int(ex#), rey%, gs(255 * a!))
'  Call pixels(Int(ex#), rey% - 1, 1 - gs(255 * a!))
' Next ex#
' Else
' If epx! = spx! Then
'  rex% = Round(ax#): a! = ax# - rex% + 0.5
' For ey# = Round(ay#) - 0.5 To Round(by) + 1.5 Step st
'  Call pixels(rex%, Int(ey#), gs(255 * a!))
'  Call pixels(rex% - 1, Int(ey#), 1 - gs(255 * a!))
' Next ey#: End If
' End If
' End If
' End Sub

' Private Sub pixels(X As Integer, Y As Integer, a As Single)
' Dim r2 As Long, g2 As Long, b2 As Long, s2 As Single
' Dim BGR As Long
' Select Case drawselect
' Case Is <> 4
'  BGR& = GetPixel(Form1.hdc, X, Y)
'  b2& = ((BGR& And &HFF0000) / &H10000) And &HFF
'  g2& = ((BGR& And &HFF00) / &H100) And &HFF
'  r2& = BGR& And &HFF
'  s2! = a! * s1
' End Select

' Select Case drawselect
' Case 0 '"Phasine"
'  SetPixelV Form1.hdc, X, Y, RGB(r2& - gs(r! * a!) * (r2& - r!) * s1, g2& - gs(g! * a!) * (g2& - g!) * s1, b2& - gs(b! * a!) * (b2& - b!) * s1)
' Case 1 '"Paint"
'  SetPixelV Form1.hdc, X, Y, _
'  RGB(r2& - s2! * (r2& - a! * r!), g2& - s2! * (g2& - a! * g!), b2& - s2! * (b2& - a! * b!))
' Case 2 '"Normal"
'  SetPixelV Form1.hdc, X%, Y%, RGB(r2& - s2! * (r2& - r!), g2& - s2! * (g2& - g!), b2& - s2! * (b2& - b!))
' Case 3 '"Laser Smoke"
'  SetPixelV Form1.hdc, X%, Y%, RGB(r2& + s2! * Abs(r2& - r!) ^ 0.85, g2& + s2! * Abs(g2& - g!) ^ 0.85, b2& + s2! * Abs(b2& - b!) ^ 0.85)
' Case 4 '"Pixelsticks"
'  SetPixelV Form1.hdc, X%, Y%, RGB(r! * a!, g! * a!, b! * a!)
' Case 5 '"Lighten"
'  If r2& < r! Then
'  r2& = r2& + s2! * (r! - r2&): End If
'  If g2& < g! Then
'  g2& = g2& + s2! * (g! - g2&): End If
'  If b2& < b! Then
'  b2& = b2& + s2! * (b! - b2&): End If
'  SetPixelV Form1.hdc, X%, Y%, RGB(r2&, g2&, b2&)
' End Select
' End Sub


' Private Sub waveform()
' Dim x1 As Single, y1 As Single
' Dim x2 As Single, y2 As Single
' Dim z As Single: Static firstrun As Boolean
' Dim mult1 As Single, mult2 As Single
' Static sw As Single, sh As Single, wavebreak As Boolean

'  sw = Form1.ScaleWidth / 2: sh = Form1.ScaleHeight / 2
'  mult1 = 5 * (Rnd - 0.5)
'  mult2 = 5 * (Rnd - 0.5)

'  If Not wavebreak Then
  'incr,incg,incb produce color shift per line
'   incr = 0.8: incg = -1.26: incb = 1.88
'   mstep = 11: int1 = 11000: wavebreak = True: Dither Me
'  End If

'  Do While Not gbye
'   DoEvents
'   If gbye Then Exit Do

'    x1 = sw + 120 * Sin(mult2 + z / 150) + 125 * Cos(mult2 + z / 190)
'    x2 = sw + 120 * Sin(z * mult2 / 180) + 120 * Sin(mult2 + z / 230)
'    y1 = sh + 110 * Sin(mult1 + z / 315)
'    y2 = sh + 200 * Sin(4 * mult1 + z / 249)
'    Call antialias(x1, y1, x2, y2)

'    Call colors
   
'    z = z + mstep
'    Select Case z
'    Case Is > int1
'     z = 1
'     mult1 = 5 * (Rnd - 0.5)
'     mult2 = 5 * (Rnd - 0.5)
'     y1 = 6
'     incr = y1 * (Rnd - 0.5)
'     incg = y1 * (Rnd - 0.5)
'     incb = y1 * (Rnd - 0.5)
'    End Select
'  Loop

' End Sub
' Private Sub colors()
'  r = r + incr
'  If r < 0 Then incr = -incr: r = 0
'  If r > 255 Then incr = -incr: r = 255
'  g = g + incg
'  If g < 0 Then incg = -incg: g = 0
'  If g > 255 Then incg = -incg: g = 255
'  b = b + incb
'  If b < 0 Then incb = -incb: b = 0
'  If b > 255 Then incb = -incb: b = 255
' End Sub
' Private Sub Dither(vForm As Form)
' Dim h As Integer, q As Integer, ditr As Byte, ditg As Byte, ditb As Byte
' Dim lcolor As Long

'  Select Case bwi
'  Case True: Form1.BackColor = vbWhite
'  Case Else: Form1.BackColor = vbBlack
'  End Select

'  For h = 255 To 0 Step -1
'   q = 741 - h * 3

'   ditr = backdred / 255 * h: ditg = backdgrn / 255 * h: ditb = backdblu / 255 * h

'  Select Case bwi
'   Case 0
'    lcolor = RGB(ditr, ditg, ditb): q = 601 - h * 3
'   Case Else
'    lcolor = RGB(255 - ditr, 255 - ditg, 255 - ditb): q = 741 - h * 3
'   End Select

'   vForm.Line (0, q)-(Form1.ScaleWidth, q), lcolor, B
'   vForm.Line (0, q - 1)-(Form1.ScaleWidth, q - 1), lcolor, B
'   vForm.Line (0, q - 2)-(Form1.ScaleWidth, q - 2), lcolor, B
'  Next h
' End Sub
' Private Sub Form_Unload(Cancel As Integer)
'  gbye = True
' End Sub


























































' Option Explicit
' 'fluoats@hotmail.com

' 'HEY YOU!  Replace All "' " with "", and you'll be set.
' 'Again, replace all <comment><space> with <nothing>

' 'Besides the unique control style, this code has 2 other
' 'immediately useful bits:
' '1. How to close the program safely while running an infinite loop
' '2. How to code around the double-click (so you can click rapidly)


' 'Controls
' Dim vLeft(1 To 2) As Integer, vRight(1 To 2) As Integer
' Dim vTop(1 To 2) As Integer, vBot(1 To 2) As Integer
' Dim vMax(1 To 2) As Integer, vMin(1 To 2) As Integer
' Dim vVal(1 To 2) As Single

' 'used in Form_MouseDown, MouseMove, MouseUp
' Dim yInit As Integer, mousedwn as Boolean

' Dim insidewhichbutton As Byte
' Dim barheight As Long, barcol As Long, cL As Long

'color manipulation
' Dim r As Single, g As Single, b As Single
' Dim ir As Single, ig As Single, ib As Single

' Dim n As Integer 'all-purpose
' Dim gbye As Boolean, moving As Boolean

' Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
' Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

' Private Sub Form_Load()
'  Form1.ForeColor = vbWhite: Form1.BackColor = vbBlack
'  Form1.ScaleMode = 3: Form1.AutoRedraw = False
'  Form1.Width = 7890: Form1.Height = 6192
              
 'Here we could think of things as "control array" properties

 'Slider
'  vLeft(1) = 9
'  vTop(1) = 191
'  vRight(1) = vLeft(1) + 10
'  vBot(1) = vTop(1) + 136
'  vMax(1) = 255
'  vVal(1) = 100

 '"Clear" Button
'  vLeft(2) = 10: vRight(2) = 56: vTop(2) = 356: vBot(2) = vTop(2) + 22

 'Just a different color I will use for slider value bar & button text
'  barcol = RGB(255, 177, 0)

 'Color increment
'  ir = -0.013: ig = 0.0015: ib = -0.011
' End Sub


' 'This is how you close the program during an infinite loop
' Private Sub Form_Activate()
'  Do While gbye = False
'   DoEvents             'gotta have this line
'   If gbye Then Exit Do 'gotta have this line
'   If moving Then Call rainbowcharacters
'  Loop
' End Sub
' Private Sub Form_Unload(Cancel As Integer)
'  gbye = True           'gotta have this line
' End Sub
' 'Pretty easy! - Major thanks to 'Aedseed' for helping me work this out

' Private Sub Form_Paint()
' Dim outline As Long, rr As Integer, gg As Integer, bb As Integer
' Static initclick As Boolean

'  If Not initclick Then
'   Dither Me
'   Randomize: initclick = 1
'   Form1.Font = "1"
'  End If

'  outline = vbBlue
 
 'Draw outlines of our controls
'  For n = 1 To 2
'   Form1.Line (vLeft(n), vTop(n))-(vRight(n), vTop(n)), outline
'   Form1.Line (vLeft(n), vTop(n))-(vLeft(n), vBot(n)), outline
'   Form1.Line (vRight(n), vBot(n))-(vLeft(n), vBot(n)), outline
'   Form1.Line (vRight(n), vBot(n))-(vRight(n), vTop(n)), outline
'  Next n

 'Draw slider level indicator
'  DrawNewBar (1)
 
 'Position, then print "Clear" in control number 2
'  Form1.FontSize = 10: Form1.ForeColor = barcol
'  Form1.CurrentX = vLeft(2) + 10: Form1.CurrentY = vTop(2) + 4
'  Form1.Print "Clear"
 
' End Sub
' Sub Dither(vForm As Form)
' Dim h As Integer, q As Integer, cRed As Byte, cGrn As Byte, cBlu As Byte
' Dim lcolor As Long

'  Form1.BackColor = vbBlack

'  For h = 255 To 0 Step -1
'   q = 741 - h * 3

'   cRed = 0 / 255 * h
'   cGrn = 0 / 255 * h
'   cBlu = 190 / 255 * h

'   lcolor = RGB(cRed, cGrn, cBlu): q = 601 - h * 3

'   vForm.Line (0, q)-(Form1.ScaleWidth, q), lcolor, B
'   vForm.Line (0, q - 1)-(Form1.ScaleWidth, q - 1), lcolor, B
'   vForm.Line (0, q - 2)-(Form1.ScaleWidth, q - 2), lcolor, B
'  Next h
' End Sub

' Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  mousedwn = 1 'CRITICAL for getting around double-click pause
'  insidewhichbutton = 0 'Also critical for smooth operation

 'This cycles through each control to determine which (if any) we clicked
'  For n = 1 To 2
'   Select Case X
'   Case vLeft(n) To vRight(n)
'    Select Case Y
'    Case vTop(n) To vBot(n)
'     insidewhichbutton = n 'We've landed inside a control's dimensions
'     yInit = 0  'initializes change-in-Y variable used by Form_MouseMove
'    End Select
'   End Select
'  Next n
 
'  Select Case insidewhichbutton
'  Case 0 'outside any buttons
'   Call singledoubleclick
'   Call Form_MouseMove(Button, Shift, X, Y)
'  Case 1 'Number 1 is the number of our slider
'   moving = 1
'   Call Form_MouseMove(Button, Shift, X, Y)
'  Case 2 'Clear button
'   Call singledoubleclick
'  End Select
  
' End Sub
' Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' If mousedwn then 'Critical
' Select Case insidewhichbutton
' Case 0
'  moving = 1
'  Form1.FontSize = 26
'  Form1.CurrentX = X: Form1.CurrentY = Y
'  Form1.Print X & ", " & Y
' Case 1
'  Select Case Y
'  Case Is <> yInit 'Dis-allows vertical scrollbar adjustments when there's no Y change
'   yInit = Y 'Also needed for the dis-allow

'   ErasePrevBar (1) 'Draws a line of background color over the current bar
    
  'This changes the "Val 'property'" of our slider
'   Select Case Y
'   Case vTop(1) To vBot(1)
'    vVal(1) = vMax(1) * (vBot(1) - Y) / (vBot(1) - vTop(1))
'   Case Is < vTop(1)
'    vVal(1) = vMax(1)
'   Case Is > vBot(1)
'    vVal(1) = vMin(1)
'   End Select
  
'   DrawNewBar (1)
     
'  End Select
' End Select
' End If 'Critical (duh)
' End Sub
' Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  moving = False
'  mousedwn = False 'Critical
' End Sub
' Private Sub Form_DblClick()
'  mousedwn = True
'  singledoubleclick 'Critical - shared single and double-click events go in here
' End Sub
' Private Sub singledoubleclick()
'  Select Case insidewhichbutton
'  Case 0 'outside any and all controls
'   moving = 1
'  Case 1 'slider
'   moving = 1
'  Case 2 'Clear button
'   Dither Me: Form_Paint
'  End Select
' End Sub

' Private Sub ErasePrevBar(n1 As Byte)
'  barheight = vBot(n1) - (vBot(n1) - vTop(n1) - 4) * (vVal(n1) - vMin(n1)) / (vMax(n1) - vMin(n1)) - 2
'  cL = GetPixel(Form1.hdc, 1, barheight)
'  Form1.Line (vLeft(n1) + 1, barheight)-(vRight(n1), barheight), cL
' End Sub
' Private Sub DrawNewBar(n1 As Byte)
'  barheight = vBot(n1) - (vBot(n1) - vTop(n1) - 4) * (vVal(n1) - vMin(n1)) / (vMax(n1) - vMin(n1)) - 2
'  Form1.Line (vLeft(n1) + 1, barheight)-(vRight(n1), barheight), barcol 'barcol set in Form_Load
' End Sub

' Private Sub rainbowcharacters()
'  Form1.CurrentX = 40 + Rnd * 60
'  Form1.CurrentY = Form1.ScaleHeight - 140 + Rnd * 60
'  Form1.ForeColor = RGB(r, g, b)
'  Form1.FontSize = 14
'  Form1.Print Chr(Rnd * 255)

'  r = r + ir * vVal(1)
'  If r < 0 Then
'   ir = -ir: r = 0
'  ElseIf r > 255 Then
'   ir = -ir: r = 255: End If

'  g = g + ig * vVal(1)
'  If g < 0 Then
'   ig = -ig: g = 0
'  ElseIf g > 255 Then
'   ig = -ig: g = 255: End If

'  b = b + ib * vVal(1)
'  If b < 0 Then
'   ib = -ib: b = 0
'  ElseIf b > 255 Then
'   ib = -ib: b = 255: End If
' End Sub

