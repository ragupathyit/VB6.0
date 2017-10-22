Attribute VB_Name = "modConvertNumber"
Option Explicit
'============================================'
'    CONVERT NUMBER TO ENGLISH               '
'    BY : IECH SETHA                         '
'    E-Mail: iech_setha@yahoo.com            '
'============================================'
Private Const P_ENG_SEXTILLION As String = "sextillion"
Private Const P_ENG_QUINTILLION As String = "quintillion"
Private Const P_ENG_QUADRILLION As String = "quadrillion"
Private Const P_ENG_TRILLION As String = "trillion"
Private Const P_ENG_BILLION As String = "billion"
Private Const P_ENG_MILLION As String = "million"
Private Const P_ENG_THOUSAND As String = "thousand"
Private Const P_ENG_HUNDRED As String = "hundred"

Private Function P_ENG_CONVERT(ByVal pNum As String) As String
'On Error Resume Next
Dim MyOneNum, MyTwoNum
Dim strHun As String
Dim FixLen As Integer

MyTwoNum = Array("", "", "twenty", "thirty", _
                 "fourty", "fifty", _
                 "sixty", "seventy", _
                 "eighty", "ninety")
  
MyOneNum = Array("", "one", "two", "three", _
                "four", "five", "six", _
                "seven", "eight", "nine", _
                "ten", "eleven", "twelve", _
                "thirteen", "fourteen", "fifteen", _
                "sixteen", "seventeen", "eighteen", "nineteen")
                
    FixLen = GetFixLen(pNum)
    
    Select Case FixLen
        Case 2: strHun = P_ENG_HUNDRED
        Case 3: strHun = P_ENG_THOUSAND
        Case 6: strHun = P_ENG_MILLION
        Case 9: strHun = P_ENG_BILLION
        Case 12: strHun = P_ENG_TRILLION
        Case 15: strHun = P_ENG_QUADRILLION
        Case 18: strHun = P_ENG_QUINTILLION
        Case 21: strHun = P_ENG_SEXTILLION
        Case Else: strHun = ""
    End Select
    'if it is in plural form
    If FixLen <> 0 Then
        If Len(pNum) >= 3 And Len(Left(pNum, _
           Len(pNum) - FixLen)) > 1 Then
           strHun = strHun & "s"
        Else
            If Len(Left(pNum, Len(pNum) - FixLen)) = 1 Then
            End If
        End If
    End If
    'if num lenght is more than 2
    If FixLen > 0 Then
       Dim strConv As String
       strConv = P_ENG_CONVERT( _
                 Left(pNum, Len(pNum) - FixLen))
       If strConv <> "" Then _
            P_ENG_CONVERT = P_ENG_CONVERT( _
                 Left(pNum, Len(pNum) - FixLen)) & _
                 " " & strHun
       If CLng(Len(pNum)) > FixLen + 1 Or _
          (CLng(Len(pNum)) = FixLen + 1 And _
          Right(pNum, FixLen) <> "") Then
          P_ENG_CONVERT = P_ENG_CONVERT & ", " & _
          P_ENG_CONVERT(Right(pNum, FixLen))
       End If
    'if number is between 20 to 99
    Else
        If pNum <> "" Then
            If CLng(pNum) >= 20 Then
                P_ENG_CONVERT = MyTwoNum(CInt(Left(pNum, 1))) & _
                            " " & P_ENG_CONVERT(Right(pNum, 1))
            Else
                P_ENG_CONVERT = MyOneNum(CInt(pNum))
            End If
        End If
    'if number is less than 20
    End If
    'if the end of string is ","
    If Right(P_ENG_CONVERT, 2) = ", " Then
        P_ENG_CONVERT = Left(P_ENG_CONVERT, _
                        Len(P_ENG_CONVERT) - 2)
    End If
End Function

Public Function ConNumToEngLish(ByVal pNum As Variant) As String
'On Error Resume Next
Dim PostNum As Long
Dim SignNum As String
If pNum & "" <> "" Then
    If Left(pNum, 1) = "-" Then
        pNum = Right(pNum, Len(pNum) - 1)
        SignNum = "minus "
    End If
   
'    pNum = CDec(pNum):
    pNum = CStr(pNum)
    PostNum = InStr(1, pNum, ".", vbBinaryCompare)
    
    If PostNum <> 0 Then
       ConNumToEngLish = CallConvert(Left(pNum, _
                        PostNum - 1)) & " point " & _
                       CallConvert(Right(pNum, _
                       Len(pNum) - PostNum), "AF")
    Else
       ConNumToEngLish = CallConvert(pNum)
    End If
    
    ConNumToEngLish = SignNum & ConNumToEngLish
    ConNumToEngLish = UCase(Left(ConNumToEngLish, 1)) & _
                    Right(ConNumToEngLish, Len(ConNumToEngLish) - 1)
  End If
End Function

Public Function CallConvert(ByVal pNum As String, Optional pBAP As String) As String
    Dim PostNum As Integer
    
    If CInt(Left(pNum, 1)) = 0 And pBAP = "AF" Then
       pNum = "1" & Right(pNum, Len(pNum) - 1)
       If Len(pNum) = 2 Then
        CallConvert = "zero " & P_ENG_CONVERT(Right(pNum, 1))
       Else
        CallConvert = P_ENG_CONVERT(pNum)
        CallConvert = "zero " & Right(CallConvert, _
                      Len(CallConvert) - 3)
       End If
    Else
        
       CallConvert = P_ENG_CONVERT(pNum)
       If CallConvert = "" Then CallConvert = "zero"
    End If
    
    PostNum = MyInStr(Len(CallConvert), _
               CallConvert, ",")
    If PostNum <> 0 Then _
    CallConvert = Left(CallConvert, PostNum - 1) & _
                       " and" & Right(CallConvert, _
                       Len(CallConvert) - PostNum)
End Function

Function GetFixLen(ByRef pNum As String) As Integer
    Dim FixLen As Integer
    Select Case Len(pNum)
        Case 4 To 6: FixLen = 3
        Case 3: FixLen = 2
        Case 7 To 9: FixLen = 6
        Case 10 To 12: FixLen = 9
        Case 13 To 15: FixLen = 12
        Case 16 To 18: FixLen = 15
        Case 19 To 21: FixLen = 18
        Case Is > 21: FixLen = 21
        Case Else:    FixLen = 0
    End Select
    '
    If Len(Left(pNum, Len(pNum) - FixLen)) = 1 Then
       If CInt(Left(pNum, Len(pNum) - FixLen)) = 0 Then
            pNum = Right(pNum, FixLen)
            FixLen = GetFixLen(pNum)
       End If
    End If
    '
    GetFixLen = FixLen
End Function

Function MyInStr(pStop As Long, Str1 As String _
                 , Str2 As String) As Long
   If Len(Str2) <= Len(Str1) Then
    Dim i As Long, wsCount As Long
    Dim FindI As Boolean: FindI = False
    For i = pStop To 1 Step -1
        wsCount = wsCount + 1
        If Mid(Str1, i - Len(Str2) + 1, Len(Str2)) = CStr(Str2) Then
            FindI = True
            Exit For
        End If
    Next i
    If FindI = True Then MyInStr = pStop - wsCount + 1
  Else
    MsgBox "The length of string one must longer than string two.", vbInformation, "Information..."
  End If
End Function

