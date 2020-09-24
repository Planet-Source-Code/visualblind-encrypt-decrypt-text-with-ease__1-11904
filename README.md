<div align="center">

## Encrypt/Decrypt Text with Ease


</div>

### Description

This is an awesome encryption algorithm that you can you really use to encrypt and decrypt text. It works on multiline textboxes and singleline as well...Have PHUN!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VisualBlind](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/visualblind.md)
**Level**          |Intermediate
**User Rating**    |3.4 (17 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/visualblind-encrypt-decrypt-text-with-ease__1-11904/archive/master.zip)





### Source Code

```
'.......................................................
'to encrypt:
'right before you save you go:
'password = Encrypt(password) password is the variable
'.......................................................
Public Function Encrypt(ByVal Plain As String)
  Dim Letter As String
  For i = 1 To Len(Plain)
    Letter = Mid$(Plain, i, 1)
    Mid$(Plain, i, 1) = Chr(Asc(Letter) + 111)
  Next i
  Encrypt = Plain
End Function
'password = Encrypt(Text1)
'Text2 = password
'...................................................
'to decrypt:
'right before you load the password you go:
'password = Decrypt(password)
'....................................................
Public Function Decrypt(ByVal Encrypted As String)
Dim Letter As String
  For i = 1 To Len(Encrypted)
    Letter = Mid$(Encrypted, i, 1)
    Mid$(Encrypted, i, 1) = Chr(Asc(Letter) - 111)
  Next i
  Decrypt = Encrypted
End Function
```

