Imports MathExt

Public Class Geocentric3p
    Inherits Transformations

    Private iSrcEll, iTgtEll As Ellipsoid
    Private idx, idy, idz As Double
    Private iRotConv As RotationConventions

    Overrides ReadOnly Property Type As Methods
        Get
            Return Methods.Geocentric3Parameter
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
        FullName = "Default"
        ShortName = "Default"
        iSrcEll = New Ellipsoid
        iTgtEll = New Ellipsoid
        idx = 0.0
        idy = 0.0
        idz = 0.0
        iRotConv = RotationConventions.PositionVector
    End Sub

    Public Sub New(sourceEllipsoid As Ellipsoid, longName As String, compactName As String, deltaX As Double, deltaY As Double, deltaZ As Double)
        FullName = longName
        ShortName = compactName
        iSrcEll = sourceEllipsoid
        iTgtEll = New Ellipsoid
        idx = deltaX
        idy = deltaY
        idz = deltaZ
        iRotConv = RotationConventions.PositionVector
    End Sub

    Public Sub New(sourceEllipsoid As Ellipsoid, targetEllipsoid As Ellipsoid, longName As String, compactName As String, deltaX As Double, deltaY As Double, deltaZ As Double)
        FullName = longName
        ShortName = compactName
        iSrcEll = sourceEllipsoid
        iTgtEll = targetEllipsoid
        idx = deltaX
        idy = deltaY
        idz = deltaZ
        iRotConv = RotationConventions.PositionVector
    End Sub

    Public Overrides Function ToWGS84(point As Point3D) As Point3D
        Dim TmpResult, TmpGC1, TmpGC2 As New Point3D

        Try
            'Source Ellipsoid costants
            Dim axis, ecc, axis2 As Double
            axis = iSrcEll.SemiMayorAxis
            ecc = iSrcEll.Eccentricity
            axis2 = iSrcEll.SemiMinorAxis

            'From LL to XYZ
            Dim V1 As Double
            V1 = iSrcEll.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            TmpGC1.X = (V1 + point.Z) * Math.Cos(point.Y) * Math.Cos(point.X)
            TmpGC1.Y = (V1 + point.Z) * Math.Cos(point.Y) * Math.Sin(point.X)
            TmpGC1.Z = ((1 - ecc ^ 2) * V1 + point.Z) * Math.Sin(point.Y)

            'From XYZ to XYZ
            TmpGC2.X = TmpGC1.X + idx
            TmpGC2.Y = TmpGC1.Y + idy
            TmpGC2.Z = TmpGC1.Z + idz

            'Target Ellipsoid costants
            axis = iTgtEll.SemiMayorAxis
            ecc = iTgtEll.Eccentricity
            axis2 = iTgtEll.SemiMinorAxis

            'From XYZ to LL
            Dim eps, p, q, V2 As Double
            eps = ecc ^ 2 / (1 - ecc ^ 2)
            p = (TmpGC2.X ^ 2 + TmpGC2.Y ^ 2) ^ 0.5
            q = Math.Atan((TmpGC2.Z * axis) / (p * axis2))

            TmpResult.X = Math.Atan(TmpGC2.Y / TmpGC2.X)
            TmpResult.Y = Math.Atan((TmpGC2.Z + eps * axis2 * Math.Sin(q) ^ 3) / (p - ecc ^ 2 * axis * Math.Cos(q) ^ 3))
            'Calculate the elevetion
            V2 = iTgtEll.GetRadiuosOfCurvatureInThePrimeVertical(TmpResult.Y)
            TmpResult.Z = (p / Math.Cos(TmpResult.Y)) - V2

            Return TmpResult
        Catch ex As Exception
            Return New Point3D
        End Try

    End Function

    Public Overrides Function FromWGS84(point As Point3D) As Point3D
        Dim TmpResult, TmpGC1, TmpGC2 As New Point3D

        Try
            'Target Ellipsoid costants
            Dim axis, ecc, axis2 As Double
            axis = iTgtEll.SemiMayorAxis
            ecc = iTgtEll.Eccentricity
            axis2 = iTgtEll.SemiMinorAxis

            'From LL to XYZ
            Dim V1 As Double
            V1 = iSrcEll.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            TmpGC1.X = (V1 + point.Z) * Math.Cos(point.Y) * Math.Cos(point.X)
            TmpGC1.Y = (V1 + point.Z) * Math.Cos(point.Y) * Math.Sin(point.X)
            TmpGC1.Z = ((1 - ecc ^ 2) * V1 + point.Z) * Math.Sin(point.Y)

            'From XYZ to XYZ
            TmpGC2.X = TmpGC1.X - idx
            TmpGC2.Y = TmpGC1.Y - idy
            TmpGC2.Z = TmpGC1.Z - idz

            'Source Ellipsoid costants
            axis = iSrcEll.SemiMayorAxis
            ecc = iSrcEll.Eccentricity
            axis2 = iSrcEll.SemiMinorAxis

            'From XYZ to LL
            Dim eps, p, q, V2 As Double
            eps = ecc ^ 2 / (1 - ecc ^ 2)
            p = (TmpGC2.X ^ 2 + TmpGC2.Y ^ 2) ^ 0.5
            q = Math.Atan((TmpGC2.Z * axis) / (p * axis2))

            TmpResult.X = Math.Atan(TmpGC2.Y / TmpGC2.X)
            TmpResult.Y = Math.Atan((TmpGC2.Z + eps * axis2 * Math.Sin(q) ^ 3) / (p - ecc ^ 2 * axis * Math.Cos(q) ^ 3))
            'Calculate the elevetion
            V2 = iTgtEll.GetRadiuosOfCurvatureInThePrimeVertical(TmpResult.Y)
            TmpResult.Z = (p / Math.Cos(TmpResult.Y)) - V2

            Return TmpResult
        Catch ex As Exception
            Return New Point3D
        End Try

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