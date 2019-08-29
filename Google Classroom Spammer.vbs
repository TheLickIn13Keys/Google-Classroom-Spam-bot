set shell = createobject ("wscript.shell")

strtext = inputbox ("What will the spam text be?")
strtimes = inputbox ("How many messages should be sent?")
strspeed = inputbox ("How fast (milliseconds)?
strtimeneed = inputbox ("How many seconds do you need to get to the location you want to spam in?")
If not isnumeric (strtimes & strspeed & strtimeneed) then
msgbox "Make sure all values apart from the message are numbers."
wscript.quit
End If
strtimeneed2 = strtimeneed * 1000
do
msgbox "You have " & strtimeneed & " seconds to get to your input area where you are going to spam."
wscript.sleep strtimeneed2
for i=0 to strtimes
shell.sendkeys (strtext & "{tab}" & "{tab}" & "{enter}")
WScript.Sleep(1000)
shell.sendkeys ("{enter}")
wscript.sleep strspeed
Next
wscript.sleep strspeed * strtimes / 10
returnvalue=MsgBox ("Want to spam again with the same info?",36)
If returnvalue=6 Then
Msgbox "Ok Spambot will activate again"
End If
If returnvalue=7 Then
msgbox "Spambot is shutting down"
wscript.quit
End IF
loop