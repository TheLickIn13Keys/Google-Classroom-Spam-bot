set shell = createobject ("wscript.shell")

text = inputbox ("What will the spam text be?")
times = inputbox ("How many messages should be sent?")
delay = inputbox ("How long between messages? (Anything lower that 1500 will make your computer go haywire! Recommended is 1800-2000, but it may vary based on network performance. The worse your network, the higher the number)")
timeneed = inputbox ("How many seconds do you need to get to the comment section of a Classroom post?")
If not isnumeric (times & timeneed) then
msgbox "Make sure all values apart from the message are numbers."
wscript.quit
End If
timeneed2 = timeneed * 1000
do
msgbox "After you click ok you will have " & timeneed & " seconds to get to the comment section of a Classroom post (Remember don't spam unless you have permission from the classroom owner!)."
wscript.sleep(timeneed2)
for i=0 to times
shell.sendkeys (text & "{tab}" & "{enter}")
WScript.Sleep(delay)
Next
wscript.sleep 1000 * times / 10
returnvalue=MsgBox ("Want to spam again with the same info?",36)
If returnvalue=6 Then
Msgbox "Ok Spambot will activate again"
End If
If returnvalue=7 Then
msgbox "Spambot is shutting down"
wscript.quit
End IF
loop
