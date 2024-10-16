Dim f
Set f = CreateObject("WSHSHELL.GUI")

f.OpenWindow
f.SetWindowTitle "This is a sample window"
f.SetWindowLocation 10, 50
f.SetWindowSize 320, 200
f.LinearGradient 1, 0, 0, 255, 0, 0, 0

f.NewButton 1, "Push This!", 1, 1

While Not f.XIsPushed
   Wscript.Sleep 100
   If f.IsPushed("btn", 1) Then
      f.FMsgBox "You pressed the button."
   End If
Wend

f.CloseWindow
Set f = Nothing