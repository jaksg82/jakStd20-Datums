Imports MathExt

Public Class AbridgedMolodensky
    Inherits Transformations

    Dim iSrcEll, iTgtEll As Ellipsoid
    Private idx, idy, idz As Double
    Dim iRotConv As RotationConventions

    Overrides ReadOnly Property Type As Methods
        Get
            Return Methods.AbridgedMolodensky
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

    Public Sub New(sourceEllipsoid As Ellipsoid, longName As String, compactName As String, deltaX As Double, deltaY As Double, deltaZ As Double)
        FullName = longName
        ShortName = compactName
        iSrcEll = sourceEllipsoid
        iTgtEll = New Ellipsoid(7030)
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
        Dim TmpResult As New Point3D

        Try
            'Source Ellipsoid costants
            Dim axis, flat As Double
            axis = iSrcEll.SemiMayorAxis
            flat = iSrcEll.Flattening

            'From LL to XYZ
            Dim V1, P1, Da, Df, Dlat, Dlon, Dh As Double
            V1 = iSrcEll.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            P1 = iSrcEll.GetRadiuosOfCurvatureInTheMeridian(point.Y)
            Da = iTgtEll.SemiMayorAxis - iSrcEll.SemiMayorAxis
            Df = iTgtEll.Flattening - iSrcEll.Flattening
            Dlat = (-idx * Math.Sin(point.Y) * Math.Cos(point.X) - idy * Math.Sin(point.Y) * Math.Sin(point.X) _
                + idz * Math.Cos(point.Y) + (axis * Df + flat * Da) * Math.Sin(2 * point.Y)) / (P1 * Math.Sin(DegRad(1 / 3600)))

            Dlon = (-idx * Math.Sin(point.X) + idy * Math.Cos(point.X)) / (V1 * Math.Cos(point.Y) * Math.Sin(DegRad(1 / 3600)))

            Dh = idx * Math.Cos(point.Y) * Math.Cos(point.X) + idy * Math.Cos(point.Y) * Math.Sin(point.X) _
                + idz * Math.Sin(point.Y) + (axis * Df + flat * Da) * Math.Sin(point.Y) ^ 2 - Da

            TmpResult.X = point.X + Dlon
            TmpResult.Y = point.Y + Dlat
            TmpResult.Z = point.Z + Dh

            Return TmpResult
        Catch ex As Exception
            Return New Point3D
        End Try

    End Function

    Public Overrides Function FromWGS84(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'Source Ellipsoid costants
            Dim axis, flat As Double
            axis = iTgtEll.SemiMayorAxis
            flat = iTgtEll.Flattening

            'From LL to XYZ
            Dim V1, P1, Da, Df, Dlat, Dlon, Dh As Double
            V1 = iTgtEll.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            P1 = iTgtEll.GetRadiuosOfCurvatureInTheMeridian(point.Y)
            Da = iSrcEll.SemiMayorAxis - iTgtEll.SemiMayorAxis
            Df = iSrcEll.Flattening - iTgtEll.Flattening
            Dlat = (idx * Math.Sin(point.Y) * Math.Cos(point.X) + idy * Math.Sin(point.Y) * Math.Sin(point.X) _
                - idz * Math.Cos(point.Y) + (axis * Df + flat * Da) * Math.Sin(2 * point.Y)) / (P1 * Math.Sin(DegRad(1 / 3600)))

            Dlon = (idx * Math.Sin(point.X) - idy * Math.Cos(point.X)) / (V1 * Math.Cos(point.Y) * Math.Sin(DegRad(1 / 3600)))

            Dh = -idx * Math.Cos(point.Y) * Math.Cos(point.X) - idy * Math.Cos(point.Y) * Math.Sin(point.X) _
                - idz * Math.Sin(point.Y) + (axis * Df + flat * Da) * Math.Sin(point.Y) ^ 2 - Da

            TmpResult.X = point.X + Dlon
            TmpResult.Y = point.Y + Dlat
            TmpResult.Z = point.Z + Dh

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