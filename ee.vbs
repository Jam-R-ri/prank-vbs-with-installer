Dim message, sapi
message=("Your computer is now compermised by jam. Enjoy finding where this is hidden!")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
do

MsgBox ("Compermised by jam :)")                

loop
