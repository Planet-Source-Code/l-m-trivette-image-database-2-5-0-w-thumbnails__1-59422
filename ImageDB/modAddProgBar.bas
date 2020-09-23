Attribute VB_Name = "modAddProgBar"


Public Function AddProgBar(pb As ProgressBar, sb As StatusBar, lPan As Long)
    ' make sure that when the form is resized that the
    ' statusbar is rsized before we continue
    sb.Align = 2
    sb.Refresh
    
    ' set the properties of the progressbar
    ' flat with no border seems to look the best
    ' also set the progressbar to the top of the zorder
    pb.ZOrder 0
    pb.Appearance = ccFlat
    pb.BorderStyle = ccNone
    
    ' now resize the progressbar1 to fit in the statusbar panel
    pb.Left = sb.Panels(lPan).Left + 25
    pb.Width = sb.Panels(lPan).Width - 45
    pb.Top = sb.Top + 45
    pb.Height = sb.Height - 75
End Function

Public Sub TimeOut(duration)
    Dim starttime As Date
    Dim X As Variant
    starttime = Timer
    Do While Timer - starttime < duration
        X = DoEvents()
    Loop
End Sub
