# Mohawk_Password_Reset
Automating the Unlock and Password Reset Process 
#'key Press functiondelayVar = 250
Sub TypeText_QA(s) Delay(delayVar) l = Length(s)                For i = 1 To l       ch = Mid(s, i, 1)  Select Case ch   Case "`"    KeyPress("Oemtilde")   Case "~"    KeyPress("Oemtilde, Shift")   Case "!"    KeyPress("Shift, D1")   Case "@"    KeyPress("Shift, D2")   Case "#"    KeyPress("Shift, D3")   Case "$"    KeyPress("Shift, D4")   Case "%"    KeyPress("Shift, D5")   Case "^"    KeyPress("Shift, D6")   Case "&"    KeyPress("Shift, D7")   Case "*"    KeyPress("Shift, D8")   Case "("    KeyPress("Shift, D9")   Case ")"    KeyPress("Shift, D0")   Case "-"    KeyPress("OemMinus")   Case "="    KeyPress("Oemplus")   Case "_"    KeyPress("Shift, OemMinus")   Case "+"    KeyPress("Shift, Oemplus")   Case " "    KeyPress("Space")   Case "["    KeyPress("OemOpenBrackets")   Case "]"    KeyPress("Oem6")   Case ";"    KeyPress("Oem1")   Case "'"    KeyPress("Oem7")   Case "\\"    KeyPress("Oem5")   Case ","    KeyPress("Oemcomma")   Case "."    KeyPress("OemPeriod")   Case "/"    KeyPress("OemQuestion")   Case "{"    KeyPress("OemOpenBrackets, Shift")   Case "}"    KeyPress("Oem6, Shift")   Case ":"    KeyPress("Oem1, Shift")   Case "\""    KeyPress("Oem7, Shift")   Case "|"    KeyPress("Oem5, Shift")   Case "<"    KeyPress("Oemcomma, Shift")   Case ">"    KeyPress("OemPeriod, Shift")   Case "?"    KeyPress("OemQuestion, Shift")   Case else    If ch = Lower(ch) Then     KeyPress(Upper(ch))    Else     KeyPress("Shift, " + ch)    End If  End Select Next                Delay(delayVar*3) KeyPress("Enter") Delay(delayVar)End Sub#


'Created: 11/6/2017 1:56 PM'Windows Version: Windows 10 Enterprise N (x64)'Processor: Intel(R) Xeon(R) CPU           E5410  @ 2.33GHz'Installed Memory (RAM): 4 GB
Script.SetContext("Password Reset/1.0")Script.RunApp()

TestData = OpenRecordset("LocalDatasheet000") 
AdminUsername = GetRowValue(TestData, "Admin Username")
Adminpassword = GetRowValue(TestData, "Admin Password")
Requestedfunction = GetRowValue(TestData, "Function to Perform")
Window("One Identity Password Manager").EditBox("UserName").SetText(AdminUsername)
Window("One Identity Password Manager").EditBox("Password").Checkpoint("Text", "Password", True, "")Window("One Identity Password Manager").EditBox("Password2").Click()
TypeText_QA(AdminPassword)
Keypress("Enter")
Window("One Identity Password Manager 3").EditBox("UserName").SetText("bill-raynor")
Window("One Identity Password Manager 3").ComboBox("DomainHash").Item("Staff").Select()
Window("One Identity Password Manager 3").Button("DoSearchUserBtn").Click()
Window("One Identity Password Manager 3").HTMLLink("Raynor, Bill").Click()
If Requestedfunction = 0 then   
Window("One Identity Password Manager 3").HTMLLink("scenario2D2E0C69137C451EB7BC4A6A4039C9E7").Click() 
Window("One Identity Password Manager 3").HTMLLink("logoutmenuitemlink").Click() 
Browser("One Identity Password Manager 3").Navigate("https://mypassword.mohawkcollege.ca/PMHelpdesk/Authentication/Login?ReturnUrl=%2FPMHelpdesk") 
Browser("One Identity Password Manager 3").Close()
else
 Window("One Identity Password Manager 3").HTMLLink("scenario74CDB89BDFD340BEB9BD8C53CA31D381").Click() 
End If


           '------  change Password  -------'
Window("One Identity Password Manager 3").EditBox("Password").SetText("@*TvFlQNVaMbLjBpU7YW64wA==")
Window("One Identity Password Manager 3").EditBox("PasswordConfirm").SetText("@*TvFlQNVaMbLjBpU7YW64wA==")
Window("One Identity Password Manager 3").Button("next").Click()
Window("One Identity Password Manager 3").HTMLLink("statusperformAnotherTaskbutton").Click()
 
 '------  Question reset Code  ------'
Window("One Identity Password Manager 3").HTMLLink("scenarioBAD9019F2D9342FF9ED67FFEBCCE6D74").Click()
Window("One Identity Password Manager 3").RadioButton("Controls.0.Value").Click()
Window("One Identity Password Manager 3").RadioButton("Controls.0.Value1").Click()
Window("One Identity Password Manager 3").RadioButton("Controls.0.Value").Click()
Window("One Identity Password Manager 3").Button("next").Click()
Window("One Identity Password Manager 3").Button("next").Click()
Window("One Identity Password Manager 3").EditBox("Passcode").Range(5, 5)
Window("One Identity Password Manager 3").EditBox("MinutesToExpire").Range(5, 5)
Window("One Identity Password Manager 3").Button("next").Click()
Window("One Identity Password Manager 3").HTMLLink("statusperformAnotherTaskbutton").Click()
Window("One Identity Password Manager 3").HTMLLink("logoutmenuitemlink").Click()
Browser("One Identity Password Manager 3").Navigate("https://mypassword.mohawkcollege.ca/PMHelpdesk/Authentication/Login?ReturnUrl=%2FPMHelpdesk")
Browser("One Identity Password Manager 3").Close("Window")
