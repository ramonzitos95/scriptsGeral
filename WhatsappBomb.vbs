' InputBoxes
Contact = InputBox("Com qual contato você quer fazer DDos?", "WhatsApp DDos")
Message = InputBox("Qual é a mensagem a ser enviada?","WhatsApp DDos")
T = InputBox("Quantas vezes a mensagem precisa ser enviada?","WhatsApp DDos")
If MsgBox("Você preencheu tudo corretamente", 1024 + vbSystemModal, "WhatsApp DDos") = vbOk Then
 
' Go To WhatsApp
Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("https://web.whatsapp.com/", 1)
 
' Loading Time
 
If MsgBox("O WhatsApp Web está aberto?" & vbNewLine & vbNewLine & "Aperte Não para cancelar", vbYesNo + vbQuestion + vbSystemModal, "WhatsApp DDos") = vbYes Then
 
' Go To The WhatsApp Search Bar
WScript.Sleep 50
WshShell.SendKeys "{TAB}"
 
' Go To The Contacts Chat
WScript.Sleep 50
WshShell.SendKeys Contact
WScript.Sleep 50
WshShell.SendKeys "{ENTER}"
 
' The Loop For The Messages
For i = 0 to T
WScript.Sleep 5
WshShell.SendKeys Message
WScript.Sleep 5
WshShell.SendKeys "{ENTER}"
Next
 
' End Of The Script
WScript.Sleep 3000
MsgBox "DDosing no " + Contact + " foi feito com sucesso", 1024 + vbSystemModal, "DDos feito"
 
' Canceled Script
Else
MsgBox "O processo foi cancelado com sucesso", vbSystemModal, "DDos Cancelado"
End If
Else
End If

