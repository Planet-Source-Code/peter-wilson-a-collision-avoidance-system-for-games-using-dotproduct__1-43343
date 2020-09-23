Attribute VB_Name = "mMisc"
Option Explicit


Private Const g_sngPIDivideBy180 = 0.0174533!


Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Public Function GetRNDNumberBetween(Min As Variant, Max As Variant) As Single

    GetRNDNumberBetween = (Rnd * (Max - Min)) + Min

End Function

