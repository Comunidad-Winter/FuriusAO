VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mcLCD"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'  Name:       mcLCD
'  Version:    1.0
'  Date:       97-07-26
'  Author:     Martin Carlsson (martin@comports.com)
'  Homepage:   http://www.algonet.se/~mmcc/
'
'  Using this class module and only two lines of your own code you can create a nice-looking,
'  non-flickering LCD display. Do like this (or check the sample application) :
'
'  In the declarations section of the form, insert this code
'
'     Dim lcdTest As New mcLCD
'
'  Put a picture-box called Picture1 on the form and add this code to the Form_Load event
'
'     lcdTest.NewLCD Picture1
'
'  And you're ready to go! Now, to set the caption of the LCD-display just change the
'  Caption property of the LCD-"control", like this
'
'     lcdTest.Caption = "01:43:10"
'
'  I almost forgot... you might want to change the back- and foreground color of the
'  picture-box too (or set the BackColor and ForeColor properties of lcdTest).
'
'  Currently, only digits 0-9 and colon are supported. And the digits tend to become a little
'  bit thin if you make the picture-box bigger - the source could be modified to adjust the
'  line thickness, but that's up to you.
'
'  This file is provided "AS IS". You can't hold me responsible for any damage that might
'  occur by using this code in any way. When distributing this sourcecode, all the
'  original files and this notice must be included. Please do not distribute modified versions.
'
'  This is freeware. You may even use it for free in your commercial products, but please
'  include a small notice like "Parts of this program written by Martin Carlsson" or at least
'  send me an e-mail. Thank you.
'
'  Copyright © 1997 Martin Carlsson

Private Type Coordinate
   X As Integer
   Y As Integer
End Type

Dim BasePoint As Coordinate

Dim SegWidth As Integer
Dim SegHeight As Integer

Dim p As PictureBox
Property Let BackColor(Color As Long)

   p.BackColor = Color

End Property

Private Sub DrawNumber(Number As Integer)

   Select Case Number
   Case 0
      DrawSegment (1)
      DrawSegment (2)
      DrawSegment (3)
      DrawSegment (4)
      DrawSegment (5)
      DrawSegment (6)
   Case 1
      DrawSegment (2)
      DrawSegment (3)
   Case 2
      DrawSegment (1)
      DrawSegment (2)
      DrawSegment (7)
      DrawSegment (5)
      DrawSegment (4)
   Case 3
      DrawSegment (1)
      DrawSegment (2)
      DrawSegment (7)
      DrawSegment (3)
      DrawSegment (4)
   Case 4
      DrawSegment (2)
      DrawSegment (3)
      DrawSegment (7)
      DrawSegment (6)
   Case 5
      DrawSegment (1)
      DrawSegment (6)
      DrawSegment (7)
      DrawSegment (3)
      DrawSegment (4)
   Case 6
      DrawSegment (1)
      DrawSegment (6)
      DrawSegment (7)
      DrawSegment (3)
      DrawSegment (4)
      DrawSegment (5)
   Case 7
      DrawSegment (1)
      DrawSegment (2)
      DrawSegment (3)
   Case 8
      DrawSegment (1)
      DrawSegment (2)
      DrawSegment (3)
      DrawSegment (4)
      DrawSegment (5)
      DrawSegment (6)
      DrawSegment (7)
   Case 9
      DrawSegment (1)
      DrawSegment (2)
      DrawSegment (3)
      DrawSegment (4)
      DrawSegment (6)
      DrawSegment (7)
   End Select

End Sub


Private Sub DrawSegment(SegNum As Integer)

'
'      1
'     ___
'    |   |
' 6  |   |  2
'    |-7-|
' 5  |   |  3
'    |___|
'
'      4
'

   Select Case SegNum
   Case 1
      p.Line (BasePoint.X + 1, BasePoint.Y)-(BasePoint.X + SegWidth - 1, BasePoint.Y)
      p.Line (BasePoint.X + 2, BasePoint.Y + 1)-(BasePoint.X + SegWidth - 2, BasePoint.Y + 1)
      p.Line (BasePoint.X + 3, BasePoint.Y + 2)-(BasePoint.X + SegWidth - 3, BasePoint.Y + 2)
   Case 2
      p.Line (BasePoint.X + SegWidth - 1, BasePoint.Y + 1)-(BasePoint.X + SegWidth - 1, BasePoint.Y + (SegHeight \ 2) - 1)
      p.Line (BasePoint.X + SegWidth - 2, BasePoint.Y + 2)-(BasePoint.X + SegWidth - 2, BasePoint.Y + (SegHeight \ 2))
      p.Line (BasePoint.X + SegWidth - 3, BasePoint.Y + 3)-(BasePoint.X + SegWidth - 3, BasePoint.Y + (SegHeight \ 2) - 1)
   Case 3
      p.Line (BasePoint.X + SegWidth - 1, BasePoint.Y + (SegHeight \ 2) + 2)-(BasePoint.X + SegWidth - 1, BasePoint.Y + SegHeight)
      p.Line (BasePoint.X + SegWidth - 2, BasePoint.Y + (SegHeight \ 2) + 1)-(BasePoint.X + SegWidth - 2, BasePoint.Y + SegHeight - 1)
      p.Line (BasePoint.X + SegWidth - 3, BasePoint.Y + (SegHeight \ 2) + 2)-(BasePoint.X + SegWidth - 3, BasePoint.Y + SegHeight - 2)
   Case 4
      p.Line (BasePoint.X + 3, BasePoint.Y + SegHeight - 2)-(BasePoint.X + SegWidth - 3, BasePoint.Y + SegHeight - 2)
      p.Line (BasePoint.X + 2, BasePoint.Y + SegHeight - 1)-(BasePoint.X + SegWidth - 2, BasePoint.Y + SegHeight - 1)
      p.Line (BasePoint.X + 1, BasePoint.Y + SegHeight)-(BasePoint.X + SegWidth - 1, BasePoint.Y + SegHeight)
   Case 5
      p.Line (BasePoint.X, BasePoint.Y + (SegHeight \ 2) + 2)-(BasePoint.X, BasePoint.Y + SegHeight)
      p.Line (BasePoint.X + 1, BasePoint.Y + (SegHeight \ 2) + 1)-(BasePoint.X + 1, BasePoint.Y + SegHeight - 1)
      p.Line (BasePoint.X + 2, BasePoint.Y + (SegHeight \ 2) + 2)-(BasePoint.X + 2, BasePoint.Y + SegHeight - 2)
   Case 6
      p.Line (BasePoint.X, BasePoint.Y + 1)-(BasePoint.X, BasePoint.Y + (SegHeight \ 2) - 1)
      p.Line (BasePoint.X + 1, BasePoint.Y + 2)-(BasePoint.X + 1, BasePoint.Y + (SegHeight \ 2))
      p.Line (BasePoint.X + 2, BasePoint.Y + 3)-(BasePoint.X + 2, BasePoint.Y + (SegHeight \ 2) - 1)
   Case 7
      p.Line (BasePoint.X + 3, BasePoint.Y + (SegHeight \ 2) - 1)-(BasePoint.X + SegWidth - 3, BasePoint.Y + (SegHeight \ 2) - 1)
      p.Line (BasePoint.X + 2, BasePoint.Y + (SegHeight \ 2))-(BasePoint.X + SegWidth - 2, BasePoint.Y + (SegHeight \ 2))
      p.Line (BasePoint.X + 3, BasePoint.Y + (SegHeight \ 2) + 1)-(BasePoint.X + SegWidth - 3, BasePoint.Y + (SegHeight \ 2) + 1)
   End Select

End Sub


Public Property Let Caption(ByVal Value As String)
Dim OrigX As Integer

   OrigX = BasePoint.X
   p.Cls

   While Value <> ""
      If Left$(Value, 1) <> ":" Then
         DrawNumber (Val(Left$(Value, 1)))
         BasePoint.X = BasePoint.X + SegWidth + 3
      Else
         p.Line (BasePoint.X + (SegWidth \ 2) - 4, BasePoint.Y + (SegHeight \ 2) - 6)-(BasePoint.X + (SegWidth \ 2), BasePoint.Y + (SegHeight \ 2) - 3), , BF
         p.Line (BasePoint.X + (SegWidth \ 2) - 4, BasePoint.Y + (SegHeight \ 2) + 4)-(BasePoint.X + (SegWidth \ 2), BasePoint.Y + (SegHeight \ 2) + 7), , BF
         BasePoint.X = BasePoint.X + SegWidth
      End If
      Value = Right$(Value, Len(Value) - 1)
   Wend

   BasePoint.X = OrigX

End Property
Property Let ForeColor(Color As Long)

   p.ForeColor = Color

End Property

Public Sub NewLCD(PBox As PictureBox)

   Set p = PBox
   
   p.ScaleMode = 3               ' pixel
   p.AutoRedraw = True
   
   BasePoint.X = 2
   BasePoint.Y = 2
   
   SegHeight = p.ScaleHeight - 6
   SegWidth = (SegHeight \ 2) + 2

End Sub


