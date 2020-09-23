Attribute VB_Name = "DDIO2"
' The MethodNumber on Encryption and Decryption must be the same when processing to result correctly
Public Function EncryptDDIO2(Estring As String, MethodNumber As Long, dDelimiter)
Dim x As Long
Dim EREV, bHold, wHold As String
EREV = StrReverse(Estring) 'Reverse the string
bHold = "": wHold = ""
For x = 1 To Len(EREV) 'Create a number code for each character
bHold = Mid(EREV, x, 1) 'Get individual letter
wHold = wHold & Asc(bHold) + MethodNumber & dDelimiter 'Change individual letter into a number + MethodNumber followed by the delimeter
Next x
EncryptDDIO2 = wHold 'Finished encrypting
End Function
Public Function DecryptDDIO2(Dstring As String, MethodNumber As Long, dDelimiter)
Dim xi, num, bHold1, bHold As Long
Dim wHold As String
num = 0: wHold = ""
On Error GoTo gotit 'Errors when cannot get next item between delimeters
Do
bHold = Split(Dstring, dDelimiter)(num) 'Get the next item between the delimeter
bHold1 = bHold - MethodNumber 'Subtract MethodNumber from the number because you added MethodNumber in the encryption
wHold = wHold & Chr(bHold1) 'Change the number into a character
num = num + 1 'Add for next item
Loop
gotit:
DecryptDDIO2 = StrReverse(wHold) 'Reverse the final string
End Function
