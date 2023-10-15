'# MWS Version: Version 2022.4 - Apr 26 2022 - ACIS 31.0.1 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 0.0 fmax = 0.0
'# created = '[VERSION]2022.4|31.0.1|20220426[/VERSION]


'@ Background Initial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Background
.ResetBackground
.Type "PEC"
End With

'@ Unit Initial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Units
.Geometry "mm"
.Frequency "ghz"
.Time "ns"
End With

'@ Boundary Initial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Boundary
.Xmin "electric"
.Xmax "electric"
.Ymin "electric"
.Ymax "electric"
.Zmin "electric"
.Zmax "electric"
.Xsymmetry "none"
.Ysymmetry "none"
.Zsymmetry "none"
End With

