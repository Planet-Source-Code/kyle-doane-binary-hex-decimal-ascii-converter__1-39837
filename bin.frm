VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "lego Converter"
   ClientHeight    =   1890
   ClientLeft      =   3240
   ClientTop       =   2700
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4410
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Option4 
         Caption         =   "Binary"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1150
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Hex Decimal"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   800
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   400
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ASCII"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   80
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function DecToBin(dec)
Dim ret As String
Dim num As Variant
Dim bin As Byte

'############## Decimal To Binary ##############
num = dec
bin = 0
ret$ = ""
Do: DoEvents
    num = num / 2
        'If the division of num gives you a
        'fraction, then subtract .5 and set the
        'value of bin equal to 1
        If InStr(1, num, ".", vbBinaryCompare) Then
            num = num - 0.5
            bin = 1
        'If the division of num doesn't give
        'you a fraction, the set bin equal
        'to 0
        Else
            bin = 0
        End If
    'Input the binary first, and then the rest
    'of the value of ret
    ret = bin & ret$
Loop Until num = 0
'Make sure the return value is 8 digits long
Do While Len(ret$) <> 8: DoEvents
    ret = "0" & ret$
Loop
DecToBin = ret$
End Function


Function BinToDec(bin)

'############## Binary To Decimal ##############

Do While Len(bin) <> 8: DoEvents
    bin = "0" & bin
Loop
bin1 = Right(bin, 1) * 2 ^ 0
bin2 = Mid(bin, 7, 1) * 2 ^ 1
bin3 = Mid(bin, 6, 1) * 2 ^ 2
bin4 = Mid(bin, 5, 1) * 2 ^ 3
bin5 = Mid(bin, 4, 1) * 2 ^ 4
bin6 = Mid(bin, 3, 1) * 2 ^ 5
bin7 = Mid(bin, 2, 1) * 2 ^ 6
bin8 = Left(bin, 1) * 2 ^ 7
BinToDec = bin1 + bin2 + bin3 + bin4 + bin5 + bin6 + bin7 + bin8
End Function

Function BinToHex(hexval)
Dim hex1, hex2, hexd1, hexd2 As Variant
Dim hexd11, hexd12, hexd21, hexd22 As Variant
Dim hex11, hex12, hex13, hex14, hex21 As Integer
Dim hex22, hex23, hex24 As Integer

'############### Binary To Hex #################

'Get first group of four from binary
hex1 = Left(hexval, 4)
'Get second group of four from binary
hex2 = Right(hexval, 4)
'Get decimal of first hex
hex11 = Right(hex1, 1) * 2 ^ 0
hex12 = Mid(hex1, 3, 1) * 2 ^ 1
hex13 = Mid(hex1, 2, 1) * 2 ^ 2
hex14 = Left(hex1, 1) * 2 ^ 3
hexd1 = hex11 + hex12 + hex13 + hex14
'Get decimal of second hex
hex21 = Right(hex2, 1) * 2 ^ 0
hex22 = Mid(hex2, 3, 1) * 2 ^ 1
hex23 = Mid(hex2, 2, 1) * 2 ^ 2
hex24 = Left(hex2, 1) * 2 ^ 3
hexd2 = hex21 + hex22 + hex23 + hex24
'Convert the values of 10 - 15 into hex form
Select Case hexd1
    Case 10
        hexd1 = "a"
    Case 11
        hexd1 = "b"
    Case 12
        hexd1 = "c"
    Case 13
        hexd1 = "d"
    Case 14
        hexd1 = "e"
    Case 15
        hexd1 = "f"
    Case Is > 15
        'If the value is greater than 15,
        'separate the two digits, add one to the
        'left most and subtract 6 from the right
        'most
        hexd11 = Left(hexd1, 1) + 1
        hexd12 = Right(hexd1, 1) - 6
        hexd1 = hexd11 & hexd12
End Select
'Convert the values of 10 - 15 into hex form
Select Case hexd2
    Case 10
        hexd2 = "a"
    Case 11
        hexd2 = "b"
    Case 12
        hexd2 = "c"
    Case 13
        hexd2 = "d"
    Case 14
        hexd2 = "e"
    Case 15
        hexd2 = "f"
    Case Is > 15
        'If the value is greater than 15,
        'separate the two digits, add one to the
        'left most and subtract 6 from the right
        'most
        hexd21 = Left(hexd2, 1) + 1
        hexd22 = Right(hexd2, 1) - 6
        hexd2 = hexd21 & hexd22
End Select
BinToHex = hexd1 & hexd2
End Function

Function HexToBin(hex)

'################ Hex To Binary ################

hex1 = Left(hex, 1)
hex2 = Right(hex, 1)
Select Case LCase(hex1)
    Case "a"
        hex1 = 10
    Case "b"
        hex1 = 11
    Case "c"
        hex1 = 12
    Case "d"
        hex1 = 13
    Case "e"
        hex1 = 14
    Case "f"
        hex1 = 15
End Select
hex1 = DecToBin(hex1)
hex1 = Right(hex1, 4)
Select Case LCase(hex2)
    Case "a"
        hex2 = 10
    Case "b"
        hex2 = 11
    Case "c"
        hex2 = 12
    Case "d"
        hex2 = 13
    Case "e"
        hex2 = 14
    Case "f"
        hex2 = 15
End Select
hex2 = DecToBin(hex2)
hex2 = Right(hex2, 4)
HexToBin = hex1 & hex2
End Function

Private Sub Command1_Click()
If Option1.Value = True And Text1.Text = "" Then Exit Sub
If Option2.Value = True And Text2.Text = "" Then Exit Sub
If Option3.Value = True And Text3.Text = "" Then Exit Sub
If Option4.Value = True And Text4.Text = "" Then Exit Sub

'Option1
If Option1.Value = True Then
    Text2.Text = Asc(Text1.Text)
    Text4.Text = DecToBin(Text2.Text)
    Text3.Text = BinToHex(Text4.Text)
End If

'Option2
If Option2.Value = True Then
    Text4.Text = DecToBin(Text2.Text)
    Text3.Text = BinToHex(Text4.Text)
    Text1.Text = Chr(Text2.Text)
End If

'Option3
If Option3.Value = True Then
    If Len(Text3.Text) <> 2 Then Text3.Text = "0" & Text3.Text
    Text4.Text = HexToBin(Text3.Text)
    Text2.Text = BinToDec(Text4.Text)
    Text1.Text = Chr(Text2.Text)
End If

'Option4
If Option4.Value = True Then
    Do While Len(Text4.Text) <> 8: DoEvents
        Text4.Text = "0" & Text4.Text
    Loop
    Text2.Text = BinToDec(Text4.Text)
    Text3.Text = BinToHex(Text4.Text)
    Text1.Text = Chr(Text2.Text)
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Option1_Click()
Text1.BackColor = &H80000009
Text1.Locked = False
Text1.TabStop = True
Text1.SetFocus

Text2.BackColor = &H8000000F
Text2.Locked = True
Text2.TabStop = False

Text3.BackColor = &H8000000F
Text3.Locked = True
Text3.TabStop = False

Text4.BackColor = &H8000000F
Text4.Locked = True
Text4.TabStop = False
End Sub

Private Sub Option2_Click()
Text1.BackColor = &H8000000F
Text1.Locked = True
Text1.TabStop = False

Text2.BackColor = &H80000009
Text2.Locked = False
Text2.TabStop = True
Text2.SetFocus

Text3.BackColor = &H8000000F
Text3.Locked = True
Text3.TabStop = False

Text4.BackColor = &H8000000F
Text4.Locked = True
Text4.TabStop = False
End Sub


Private Sub Option3_Click()
Text1.BackColor = &H8000000F
Text1.Locked = True
Text1.TabStop = False

Text2.BackColor = &H8000000F
Text2.Locked = True
Text2.TabStop = False

Text3.BackColor = &H80000009
Text3.Locked = False
Text3.TabStop = True
Text3.SetFocus

Text4.BackColor = &H8000000F
Text4.Locked = True
Text4.TabStop = False
End Sub


Private Sub Option4_Click()
Text1.BackColor = &H8000000F
Text1.Locked = True
Text1.TabStop = False

Text2.BackColor = &H8000000F
Text2.Locked = True
Text2.TabStop = False

Text3.BackColor = &H8000000F
Text3.Locked = True
Text3.TabStop = False

Text4.BackColor = &H80000009
Text4.Locked = False
Text4.TabStop = True
Text4.SetFocus
End Sub


Private Sub Text2_Change()
On Error GoTo non
'If they didn't input nothing, then don't do
'anything.
If Text2.Text = "" Then Exit Sub
'If the value entered is more than 255 then
'delete the last digit and display an error.
If Text2.Text > 255 Then
    Text2.Text = Left(Text2.Text, 2)
    Text2.SelStart = 2
    Exit Sub
End If
Exit Sub

non:
'If there is an error then delete the last digit
'and display an error.
Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
Text2.SelStart = Len(Text2.Text)
End Sub


Private Sub Text4_Change()
On Error GoTo non
If Text4.Text = "" Then Exit Sub
If Right(Text4.Text, 1) > 1 Then
    Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
    Text4.SelStart = Len(Text4.Text)
    Beep
End If
Exit Sub
non:
Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
Text4.SelStart = Len(Text4.Text)
Beep
End Sub


