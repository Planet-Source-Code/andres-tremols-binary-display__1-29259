Attribute VB_Name = "Module1"
Option Explicit


Public Sub SetNibble(nibblenum As Integer, hexvalue As String)
Debug.Print "me llamaron, hexvalue =" & hexvalue; "y nibblenum" & nibblenum

Dim i As Integer
Dim offset As Integer

offset = 16 - (nibblenum * 4) + (4 - nibblenum)

Dim value As String
Dim strnibble As String
strnibble = cnibble(hexvalue)
For i = 0 To 3
value = Mid(strnibble, (i + 1), 1)
Form1.wordgrid.TextMatrix(0, (offset + i)) = value
Next
End Sub

Public Function cnibble(hexn As String) As String
Select Case hexn
Case "0"
cnibble = "0000"
Case "1"
cnibble = "0001"
Case "2"
cnibble = "0010"
Case "3"
cnibble = "0011"
Case "4"
cnibble = "0100"
Case "5"
cnibble = "0101"
Case "6"
cnibble = "0110"
Case "7"
cnibble = "0111"
Case "8"
cnibble = "1000"
Case "9"
cnibble = "1001"
Case "A"
cnibble = "1010"
Case "B"
cnibble = "1011"
Case "C"
cnibble = "1100"
Case "D"
cnibble = "1101"
Case "E"
cnibble = "1110"
Case "F"
cnibble = "1111"
End Select
End Function

Public Sub SetNibbles(xInput As Integer)
Dim HexStr As String
HexStr = Hex(xInput)
Dim i As Integer
Dim singleDigit As String
Dim hexStrlen As Integer
hexStrlen = Len(HexStr)
For i = 1 To 4 Step 1
If hexStrlen < 1 Then
Call SetNibble(i, "0")
Else
singleDigit = Mid(HexStr, hexStrlen, 1)
hexStrlen = hexStrlen - 1
Call SetNibble(i, singleDigit)
End If

Next

End Sub
