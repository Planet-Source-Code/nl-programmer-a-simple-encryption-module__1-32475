<div align="center">

## A simple encryption module


</div>

### Description

This is a quick little encryption file I put togeather in about half an hour. It's reasonably secure for whatever you want to use it for

Feel free to add this to whatever program you want

to.
 
### More Info
 
The "Encrypt" call turns the message you call it with into a set of numbers and the "Decrypt" call takes those numbers and turns them back into a string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[NL\_Programmer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nl-programmer.md)
**Level**          |Beginner
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nl-programmer-a-simple-encryption-module__1-32475/archive/master.zip)





### Source Code

```
' Both of these functions return a string value
' just call the encrypt and it returns a set of numbers
'
' decrypt takes the set of numbers and decrypts it
' sweet and simple but it works well to make save files (At least thats what I use it for)
'
' P.S. Don't mess with anything. It works as it does and I
' Don't feel like comenting every line
Public Function encrypt(Message As String) As String
Randomize
On Error GoTo errorcheck
Dim tempmessage As String
Dim basea As Integer
Dim tempbasea As String
Message = Reverse_String(Message)
tempmessage = CStr(Message)
basea = Int(Rnd * 75) + 25
If basea < 0 Then
  tempbasea = CStr(basea)
  tempbasea = Right(tempbasea, Len(tempbasea) - 1)
  basea = CInt(tempbasea)
End If
basea = basea / 2
encrypt = CStr(basea) + ";"
For x = 1 To Len(tempmessage)
  encrypt = encrypt + CStr(Asc(Left(tempmessage, x)) - basea) + ";"
  basea = basea + 1
  tempmessage = Right(tempmessage, Len(tempmessage) - 1)
Next x
errorcheck:
End Function
Public Function decrypt(code As String) As String
On Error GoTo errorcheck
Dim basea As Integer
Dim tempcode As String
Do Until Left(code, 1) = ";"
  tempcode = tempcode + Left(code, 1)
  code = Right(code, Len(code) - 1)
Loop
basea = CInt(tempcode)
tempcode = ""
code = Right(code, Len(code) - 1)
Do Until code = ""
Do Until Left(code, 1) = ";"
  tempcode = tempcode + Left(code, 1)
  code = Right(code, Len(code) - 1)
Loop
  decrypt = decrypt + Chr(CLng(tempcode) + basea)
  code = Right(code, Len(code) - 1)
  tempcode = ""
  basea = basea + 1
Loop
decrypt = Reverse_String(decrypt)
errorcheck:
End Function
Public Function Reverse_String(Message As String) As String
For x = 1 To Len(Message)
  Reverse_String = Reverse_String + Left(Right(Message, x), 1)
Next x
End Function
```

