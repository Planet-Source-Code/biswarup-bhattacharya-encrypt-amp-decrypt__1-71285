Attribute VB_Name = "Module1"
Option Explicit

Public Const sDefaultWHEEL1 As String = "ABCDEFGHIJKLMNOPQRSTVUWXYZ_1234567890qwertyuiopasd!@#$%^&*(),. ~`-=\?/'""fghjklzxcvbnm"
Public Const sDefaultWHEEL2 As String = "IWEHJKTLZVOPFG_1234567890qwerBNMQRYUASDXCfghjklzxc ~`-=\?/'""!@#$%^&*(),.vbnmtyuiopasd"

Const sDEF_PREFIX As String = "---START-ENIGMA-MESSAGE---"
Const sDEF_SUFFIX As String = "----END-ENIGMA-MESSAGE----"

' Encrypts the string:
Function Encrypt_PRO(ByVal sInput As String, sPASSWORD As String, bExtra As Boolean) As String

    Dim sWHEEL1 As String
    Dim sWHEEL2 As String

    sWHEEL1 = sDefaultWHEEL1
    sWHEEL2 = sDefaultWHEEL2

    ' fixing line break bug
    sInput = Replace(sInput, Chr(10), "")
    sInput = Replace(sInput, Chr(13), "")
    sInput = Replace(sInput, vbTab, "")
    
    If bExtra Then sInput = Replace(sInput, " ", "")
    

    ' We use password to scramble the wheels:
    ScrambleWheels sWHEEL1, sWHEEL2, sPASSWORD
    
    

    Dim k As Long ' keeps index of the character on the wheel.
    Dim c As String ' to keep single character.
    
    Dim i As Long ' for current character index of source string.
    
    Dim sResult As String ' the result.
    sResult = ""
    
    For i = 1 To Len(sInput)
    
            ' Get character(i)
            c = Mid(sInput, i, 1)
    
            ' Find character(i) on the first wheel:
            k = InStr(1, sWHEEL1, c, vbBinaryCompare)
            
            If k > 0 Then
                ' Get the character with that index from the second
                ' wheel, and add it to result:
                sResult = sResult & Mid(sWHEEL2, k, 1)
            Else
                ' not found on the wheel, leave as it is:
                sResult = sResult & c
            End If
    
            ' Rotate first wheel to the left:
            sWHEEL1 = LeftShift(sWHEEL1)
            
            ' Rotate second wheel to the right:
            sWHEEL2 = RightShift(sWHEEL2)
    
    Next i
    
    Encrypt_PRO = sDEF_PREFIX & vbNewLine & cut_lines(sResult, 40) & sDEF_SUFFIX & vbNewLine
    
End Function


Private Function cut_lines(sInput As String, lBreakAfter As Long) As String

On Error GoTo err1

    Dim sResult As String
    
    sResult = ""
    
    Dim l As Long
    
    For l = 1 To Len(sInput) Step lBreakAfter
        
        sResult = sResult & Mid(sInput, l, lBreakAfter) & vbNewLine
    
    Next l
    
    cut_lines = sResult
        
    Exit Function
err1:
    cut_lines = sInput ' unchanged
    Debug.Print "cut_lines:" & Err.Description

End Function

' Decrypts the string.
' you may note that the only difference is
' the sWHEEL1 and sWHEEL2 exchange, instead of
' looking for character in sWHEEL1 we look it
' in sWHEEL2:
Function Decrypt_PRO(ByVal sInput As String, sPASSWORD As String) As String


    Dim sWHEEL1 As String
    Dim sWHEEL2 As String

    sWHEEL1 = sDefaultWHEEL1
    sWHEEL2 = sDefaultWHEEL2


    ' fixing the addition of prefix char bug
    Dim l As Long
    l = InStr(1, sInput, sDEF_PREFIX, vbTextCompare)
    If l > 0 Then
        sInput = Mid(sInput, l + Len(sDEF_PREFIX))
    End If
    l = InStr(1, sInput, sDEF_SUFFIX, vbTextCompare)
    If l > 0 Then
        sInput = Mid(sInput, 1, l - 1)
    End If
        

    ' fixing line break bug
    sInput = Replace(sInput, Chr(10), "")
    sInput = Replace(sInput, Chr(13), "")
    sInput = Replace(sInput, vbTab, "")

    ' We use password to "de"-scramble the wheels:
    ScrambleWheels sWHEEL1, sWHEEL2, sPASSWORD

    Dim k As Long ' keeps index of the character on the wheel.
    
    Dim i As Long ' for current character index of source string.
    Dim c As String ' to keep single character.
    
    Dim sResult As String ' the result.
    sResult = ""
    
    For i = 1 To Len(sInput)
    
            ' Get character(i)
            c = Mid(sInput, i, 1)
    
            ' Find character(i) on the second wheel:
            k = InStr(1, sWHEEL2, c, vbBinaryCompare)
            
            If k > 0 Then
                ' Get the character with that index from the first
                ' wheel, and add it to result:
                sResult = sResult & Mid(sWHEEL1, k, 1)
            Else
                ' not found on the wheel, leave as it is:
                sResult = sResult & c
            End If
    
            ' Rotate first wheel to the left:
            sWHEEL1 = LeftShift(sWHEEL1)
            
            ' Rotate second wheel to the right:
            sWHEEL2 = RightShift(sWHEEL2)
    
    Next i
    
    Decrypt_PRO = sResult
    
End Function

' Rotates the wheel (string).
' the first character goes to the end, all
' other characters go one step to the left side.
' For example:
'     "ABCD"
' will be
'     "BCDA"
' after rotation.
Function LeftShift(s As String) As String
    ' tricky way :)
    If Len(s) > 0 Then LeftShift = Mid(s, 2, Len(s) - 1) & Mid(s, 1, 1)
End Function


' Rotates the wheel (string).
' the last character goes to the beginning, all
' other characters go one step to the right side.
' For example:
'     "ABCD"
' will be
'     "DABC"
' after rotation.
Function RightShift(s As String) As String
    ' tricky way :)
    If Len(s) > 0 Then RightShift = Mid(s, Len(s), 1) & Mid(s, 1, Len(s) - 1)
End Function


' This sub scrambles the wheels.
' Wheels should be set to the same position
' for both encryption and decryption !
' (and this can be achieved by using the same password :)
' Bigger password = better scramble!
Sub ScrambleWheels(ByRef sW1 As String, ByRef sW2 As String, sPASSWORD As String)

Dim i As Long
Dim k As Long

For i = 1 To Len(sPASSWORD)
    
    For k = 1 To Asc(Mid(sPASSWORD, i, 1)) * i
        sW1 = LeftShift(sW1)
        sW2 = RightShift(sW2)
    Next k

Next i

' Who said there are no pointers in VB?

End Sub

