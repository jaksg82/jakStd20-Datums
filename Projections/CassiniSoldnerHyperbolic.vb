Imports MathExt
Imports StringFormat

Public Class CassiniSoldnerHyperbolic
    Inherits Projections

    Private iLatO, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.CassiniSoldnerHyperbolic
        End Get
    End Property

    ReadOnly Property LatitudeOfOrigin As Double
        Get
            Return iLatO
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

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfOrigin As Double,
                   lonOfOrigin As Double, falseEast As Double, falseNorth As Double)
        Me.FullName = longName
        Me.ShortName = compactName
        iLatO = latOfOrigin
        iLonO = lonOfOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        Me.BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim axis, ecc As Double
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity
            'From LL to EN
            Dim X, A, T, C, V, P, M, Mo As Double
            V = BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            P = BaseEllipsoid.GetRadiuosOfCurvatureInTheMeridian(point.Y)
            A = (point.X - iLonO) * Math.Cos(point.Y)
            T = Math.Tan(point.Y) ^ 2
            C = ecc ^ 2 * Math.Cos(point.Y) ^ 2 / (1 - ecc ^ 2)
            'Splitted factors for the M computation
            Dim A1, A2, A3, A4 As Double
            A1 = (1 - ecc ^ 2 / 4 - 3 * ecc ^ 4 / 64 - 5 * ecc ^ 6 / 256)
            A2 = (3 * ecc ^ 2 / 8 + 3 * ecc ^ 4 / 32 + 45 * ecc ^ 6 / 1024)
            A3 = (15 * ecc ^ 4 / 256 + 45 * ecc ^ 6 / 1024)
            A4 = (35 * ecc ^ 6 / 3072)
            M = axis * (A1 * point.Y - A2 * Math.Sin(2 * point.Y) + A3 * Math.Sin(4 * point.Y) - A4 * Math.Sin(6 * point.Y))
            Mo = axis * (A1 * iLatO - A2 * Math.Sin(2 * iLatO) + A3 * Math.Sin(4 * iLatO) - A4 * Math.Sin(6 * iLatO))
            X = M - Mo + V * Math.Tan(point.Y) * (A ^ 2 / 2 + (5 - T + 6 * C) * A ^ 4 / 24)
            TmpResult.X = iEastO + V * (A - T * A ^ 3 / 6 - (8 - T + 8 * C) * T * A ^ 5 / 120)
            TmpResult.Y = iNorthO + X - (X ^ 3 / (6 * P * V))
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim axis, ecc As Double
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity
            'Splitted factors for the M computation
            Dim A1, A2, A3, A4 As Double
            A1 = (1 - ecc ^ 2 / 4 - 3 * ecc ^ 4 / 64 - 5 * ecc ^ 6 / 256)
            A2 = (3 * ecc ^ 2 / 8 + 3 * ecc ^ 4 / 32 + 45 * ecc ^ 6 / 1024)
            A3 = (15 * ecc ^ 4 / 256 + 45 * ecc ^ 6 / 1024)
            A4 = (35 * ecc ^ 6 / 3072)

            'From EN to LL
            Dim D, T1, M1, Mo, u1, e1, lat1, V1, P1, q, q1 As Double
            Mo = axis * (A1 * iLatO - A2 * Math.Sin(2 * iLatO) + A3 * Math.Sin(4 * iLatO) - A4 * Math.Sin(6 * iLatO))
            e1 = (1 - (1 - ecc ^ 2) ^ 0.5) / (1 + (1 - ecc ^ 2) ^ 0.5)
            lat1 = iLatO + (point.Y - iNorthO) / 315320
            V1 = BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(lat1)
            P1 = BaseEllipsoid.GetRadiuosOfCurvatureInTheMeridian(lat1)
            q1 = (point.Y - iNorthO) ^ 3 / 6 * P1 * V1
            q = (point.Y - iNorthO + q1) ^ 3 / 6 * P1 * V1
            M1 = Mo + (point.Y - iNorthO) + q
            u1 = M1 / (axis * (1 - ecc ^ 2 / 4 - 3 * ecc ^ 4 / 64 - 5 * ecc ^ 6 / 256))
            'Splitted factors for the lat1 computation
            Dim B1, B2, B3, B4 As Double
            B1 = 3 * e1 / 2 - 27 * e1 ^ 3 / 32
            B2 = 21 * e1 ^ 2 / 16 - 55 * e1 ^ 4 / 32
            B3 = 151 * e1 ^ 3 / 96
            B4 = 1097 * e1 ^ 4 / 512
            lat1 = u1 + B1 * Math.Sin(2 * u1) + B2 * Math.Sin(4 * u1) + B3 * Math.Sin(6 * u1) + B4 * Math.Sin(8 * u1)
            T1 = Math.Tan(lat1) ^ 2
            D = (point.X - iEastO) / V1

            TmpResult.Y = lat1 - (V1 * Math.Tan(lat1) / P1) * (D ^ 2 / 2 - (1 + 3 * T1) * D ^ 4 / 24)
            TmpResult.X = iLonO + (D - T1 * D ^ 3 / 3 + (1 + 3 * T1) * T1 * D ^ 5 / 15) / Math.Cos(lat1)
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
        xproj.Add(New XElement(x.OriginLatitude, iLatO))
        xproj.Add(New XElement(x.OriginLongitude, iLonO))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Latitude of Origin", LatitudeOfOrigin, ParamType.LatLong, True),
            New ParamNameValue("Longitude of Origin", LongitudeOfOrigin, ParamType.LatLong, False),
            New ParamNameValue("False Easting at Origin", FalseEasting, ParamType.EastNorth, False),
            New ParamNameValue("False Northing at Origin", FalseNorthing, ParamType.EastNorth, True)
        }

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

        'Longitude of central meridian, (d.m.s. E/W)
        Dim H2200 As String = ps.H2200
        H2200 = H2200 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude), (d.m.s. N/E)
        Dim H2301 As String = ps.H2301
        H2301 = H2301 & FormatDMS(LatitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
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
        H2402 = H2402 & FormatDMS(LatitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class