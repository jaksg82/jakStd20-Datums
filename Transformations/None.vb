Imports MathExt

Public Class None
    Inherits Transformations

    Private iSrcEll, iTgtEll As Ellipsoid
    Private idx, idy, idz As Double
    Private iRotConv As RotationConventions

    Overrides ReadOnly Property Type As Methods
        Get
            Return Methods.None
        End Get
    End Property

    Public Overrides Property RotationConvention As RotationConventions
        Get
            Return iRotConv
        End Get
        Set(value As RotationConventions)
            iRotConv = value
        End Set
    End Property

    Public Sub New()
        FullName = "None"
        ShortName = "None"
        iSrcEll = New Ellipsoid
        iTgtEll = New Ellipsoid
        idx = 0.0
        idy = 0.0
        idz = 0.0
        iRotConv = RotationConventions.PositionVector
    End Sub

    Public Overrides Function ToWGS84(point As Point3D) As Point3D
        Return point
    End Function

    Public Overrides Function FromWGS84(point As Point3D) As Point3D
        Return point
    End Function

    Public Overrides Function ToXml() As XElement
        'New(Source As Ellipsoid, FullName As String, ShortName As String, DeltaX As Double, DeltaY As Double, DeltaZ As Double)
        Dim x As New XmlTags
        Dim ellx As New XElement(x.Transformation)
        ellx.Add(New XElement(x.Type, Type.ToString))
        ellx.Add(New XElement(x.FullName, FullName))
        ellx.Add(New XElement(x.ShortName, ShortName))
        ellx.Add(New XElement(x.DeltaX, idx))
        ellx.Add(New XElement(x.DeltaY, idy))
        ellx.Add(New XElement(x.DeltaZ, idz))
        ellx.Add(New XElement(x.RotationConvention, iRotConv))
        ellx.Add(New XElement(x.SourceEllipsoid, iSrcEll.ToXml(False)))
        ellx.Add(New XElement(x.TargetEllipsoid, iTgtEll.ToXml(False)))

        Return ellx

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Delta X", idx, ParamType.Generic, True),
            New ParamNameValue("Delta Y", idy, ParamType.Generic, True),
            New ParamNameValue("Delta Z", idz, ParamType.Generic, True)
        }

        Return tmpList

    End Function

End Class