Dim speaks, speech, voice, i
speaks = "Welcome to PSG World"

' Create the speech object
Set speech = CreateObject("sapi.spvoice")


' Set to a female voice (assuming the second voice is female)
' You might need to change the index based on what voices are available on your system
If speech.GetVoices.Count > 1 Then
    Set speech.Voice = speech.GetVoices.Item(1)
End If

' Speak the text
speech.Speak speaks