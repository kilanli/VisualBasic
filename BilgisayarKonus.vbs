dim msg,sapi
msg=InputBox("Ne soyletmek istersin")
set sapi=CreateObject("sapi.spvoice")
sapi.Speak msg