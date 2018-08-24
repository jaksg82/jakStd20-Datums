Imports jakStd20_MathExt
Imports jakStd20_StringFormat

Public Class EquidistantCylindrical
    Inherits Projections

    Dim iLat1, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.EquidistantCylindrical
        End Get
    End Property

    ReadOnly Property LatitudeOfFirstStandardParallel As Double
        Get
            Return iLat1
        End Get
    End Property

    ReadOnly Property LongitudeOfOrigin As Double
        Get
            Return iLonO
        End Get
    End Property

    ReadOnly Property FalseEasting As Double
        Get
            Return iEastO
        End Get
    End Property

    ReadOnly Property FalseNorthing As Double
        Get
            Return iNorthO
        End Get
    End Property

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfFirstStandardParallel As Double,
                   lonOfOrigin As Double, falseEast As Double, falseNorth As Double)
        Me.FullName = longName
        Me.ShortName = compactName
        iLat1 = latOfFirstStandardParallel
        iLonO = lonOfOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        Me.BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            Dim axis, ecc As Double
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity

            'From LL to EN
            Dim V1, M, M0, M1, M2, M3, M4, M5, M6, M7 As Double
            V1 = BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLat1)
            M0 = (1 - (1 / 4) * ecc ^ 2 - (3 / 64) * ecc ^ 4 - (5 / 256) * ecc ^ 6 - (175 / 16384) * ecc ^ 8 - (441 / 65536) * ecc ^ 10 -
                (4851 / 1048576) * ecc ^ 12 - (14157 / 4194304) * ecc ^ 14) * point.Y
            M1 = (-(3 / 8) * ecc ^ 2 - (3 / 32) * ecc ^ 4 - (45 / 1024) * ecc ^ 6 - (105 / 4096) * ecc ^ 8 - (2205 / 131072) * ecc ^ 10 -
                (6237 / 524288) * ecc ^ 12 - (297297 / 33554432) * ecc ^ 14) * Math.Sin(2 * point.Y)
            M2 = ((15 / 256) * ecc ^ 4 + (45 / 1024) * ecc ^ 6 + (525 / 16384) * ecc ^ 8 + (1575 / 65536) * ecc ^ 10 +
                (155925 / 8388608) * ecc ^ 12 + (495495 / 33554432) * ecc ^ 14) * Math.Sin(4 * point.Y)
            M3 = (-(35 / 3072) * ecc ^ 6 - (175 / 12288) * ecc ^ 8 - (3675 / 262144) * ecc ^ 10 - (13475 / 1048576) * ecc ^ 12 - (385385 / 33554432) * ecc ^ 14) * Math.Sin(6 * point.Y)
            M4 = ((315 / 131072) * ecc ^ 8 + (2205 / 524288) * ecc ^ 10 + (43659 / 8388608) * ecc ^ 12 + (189189 / 33554432) * ecc ^ 14) * Math.Sin(8 * point.Y)
            M5 = (-(693 / 1310720) * ecc ^ 10 - (6237 / 5242880) * ecc ^ 12 - (297297 / 167772160) * ecc ^ 14) * Math.Sin(10 * point.Y)
            M6 = ((1001 / 8388608) * ecc ^ 12 + (11011 / 33554432) * ecc ^ 14) * Math.Sin(12 * point.Y)
            M7 = (-(6435 / 234881024) * ecc ^ 14) * Math.Sin(14 * point.Y)
            M = axis * (M0 + M1 + M2 + M3 + M4 + M5 + M6 + M7)

            TmpResult.X = iEastO + V1 * Math.Cos(iLat1) * (point.X - iLonO)
            TmpResult.Y = iNorthO + M
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            Dim axis, ecc As Double
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity

            'From EN to LL
            Dim X, Y, u, n, u1, u2, u3, u4, u5, u6, u7, V1 As Double
            V1 = BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLat1)
            X = point.X - iEastO
            Y = point.Y - iNorthO
            n = (1 - (1 - ecc ^ 2) ^ 0.5) / (1 + (1 - ecc ^ 2) ^ 0.5)
            u = Y / (axis * (1 - (1 / 4) * ecc ^ 2 - (3 / 64) * ecc ^ 4 - (5 / 256) * ecc ^ 6 - (175 / 16384) * ecc ^ 8 -
                (441 / 65536) * ecc ^ 10 - (4851 / 1048576) * ecc ^ 12 - (14157 / 4194304) * ecc ^ 14))
            u1 = ((3 / 2) * n - (27 / 32) * n ^ 3 + (269 / 512) * n ^ 5 - (6607 / 24576) * n ^ 7) * Math.Sin(2 * u)
            u2 = ((21 / 16) * n ^ 2 - (55 / 32) * n ^ 4 + (6759 / 4096) * n ^ 6) * Math.Sin(4 * u)
            u3 = ((151 / 96) * n ^ 3 - (417 / 128) * n ^ 5 + (87963 / 20480) * n ^ 7) * Math.Sin(6 * u)
            u4 = ((1097 / 512) * n ^ 4 - (15543 / 2560) * n ^ 6) * Math.Sin(8 * u)
            u5 = ((8011 / 2560) * n ^ 5 - (69119 / 6144) * n ^ 7) * Math.Sin(10 * u)
            u6 = ((293393 / 61440) * n ^ 6) * Math.Sin(12 * u)
            u7 = ((6845701 / 860160) * n ^ 7) * Math.Sin(14 * u)

            TmpResult.X = iLonO + X / (V1 * Math.Cos(iLat1))
            TmpResult.Y = u + u1 + u2 + u3 + u4 + u5 + u6 + u7
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToXml() As XElement
        Dim x As New XmlTags
        Dim xproj As New XElement(x.Projection)
        xproj.Add(New XElement(x.FullName, FullName))
        xproj.Add(New XElement(x.ShortName, ShortName))
        xproj.Add(New XElement(x.Type, Type))
        xproj.Add(New XElement(x.FirstLatitude, iLat1))
        xproj.Add(New XElement(x.OriginLongitude, iLonO))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue)

        tmpList.Add(New ParamNameValue("Latitude of first standard parallel", LatitudeOfFirstStandardParallel, ParamType.LatLong, True))
        tmpList.Add(New ParamNameValue("Longitude of Origin", LongitudeOfOrigin, ParamType.LatLong, False))
        tmpList.Add(New ParamNameValue("False Easting at Origin", FalseEasting, ParamType.EastNorth, False))
        tmpList.Add(New ParamNameValue("False Northing at Origin", FalseNorthing, ParamType.EastNorth, True))

        Return tmpList

    End Function

    Public Overrides Function ToP190Header() As List(Of String)
        Dim ret As New List(Of String)
        Dim ps As New ProjStrings

        'Projection Code and Name
        Dim H1800 As String = ps.H1800
        H1800 = H1800 & " 999" & If(FullName.Length > 44, FullName.Substring(0, 44), FullName.PadRight(44, " "c))
        ret.Add(H1800)
        'Projection Type for the not defined projections (Code 999)
        Dim H2600 As String = ps.H2600ProjType & ps.TypeString(Type).PadRight(44, " "c)
        ret.Add(H2600)

        'Latitude of standard parallel(s), (d.m.s. N/S)
        Dim H2100 As String = ps.H2100
        H2100 = H2100 & FormatDMS(LatitudeOfFirstStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100.PadRight(80, " "c)
        ret.Add(H2100)

        'Longitude of central meridian, (d.m.s. E/W)
        Dim H2200 As String = ps.H2200
        H2200 = H2200 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude), (d.m.s. N/E)
        Dim H2301 As String = ps.H2301
        H2301 = H2301 & FormatDMS(0.0, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2301 = H2301 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2301 = H2301.PadRight(80, " "c)
        ret.Add(H2301)

        'Grid co-ordinates at grid origin (E,N)
        Dim H2302 As String = ps.H2302
        H2302 = H2302 & FormatMetric(FalseEasting, MetricSign.Suffix, 2, False).PadLeft(20, " "c)
        H2302 = H2302 & FormatMetric(FalseNorthing, MetricSign.Suffix, 2, True).PadLeft(20, " "c)
        H2302 = H2302.PadRight(80, " "c)
        ret.Add(H2302)

        'Scale factor
        Dim H2401 As String = ps.H2401
        H2401 = H2401 & FormatNumber(1.0, 10).PadLeft(20, " "c)
        H2401 = H2401.PadRight(80, " "c)
        ret.Add(H2401)

        'Latitude/Longitude at which scale factor is defined
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfFirstStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class