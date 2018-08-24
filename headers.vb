<Assembly: CLSCompliant(False)>
<Assembly: System.Runtime.InteropServices.ComVisible(False)>

Public Class ParamNameValue
    Public Property Name As String
    Public Property Value As Double
    Public Property Type As ParamType
    Public Property IsNorthAxis As Boolean

    Public Sub New(newName As String, newValue As Double, newType As ParamType, newIsNorthAxis As Boolean)
        Name = newName
        Value = newValue
        Type = newType
        IsNorthAxis = newIsNorthAxis
    End Sub
End Class

Public Enum ParamType
    Generic
    EastNorth
    LatLong
    Angle
    ScaleFactor
    TrueFalse
    UtmFuse
End Enum