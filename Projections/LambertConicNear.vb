Imports MathExt
Imports StringFormat

Public Class LambertConicNear
    Inherits Projections

    Private iLatO, iK0, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.LambertConicNearConformal
        End Get
    End Property

    ReadOnly Property LatitudeOfNaturalOrigin As Double
        Get
            Return iLatO
        End Get
    End Property

    ReadOnly Property LongitudeOfNaturalOrigin As Double
        Get
            Return iLonO
        End Get
    End Property

    ReadOnly Property ScaleFactorAtNaturalOrigin As Double
        Get
            Return iK0
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

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfNaturalOrigin As Double, lonOfNaturalOrigin As Double,
                                            scaleAtNaturalOrigin As Double, falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLatO = latOfNaturalOrigin
        iLonO = lonOfNaturalOrigin
        iK0 = scaleAtNaturalOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim axis As Double
        Dim TmpResult As New Point3D
        'Dim Rf, mf, Ff, t0, tf, n, R0, theta As Double
        Dim A, Ai, Bi, Ci, Di, Ei, Ro, So, n, s, m, Mi, r, theta As Double

        Try
            'Ellipsoid
            axis = BaseEllipsoid.SemiMayorAxis
            'First compute constants for the projection
            n = BaseEllipsoid.Flattening / (2 - BaseEllipsoid.Flattening)
            A = 1 / (6 * BaseEllipsoid.GetRadiuosOfCurvatureInTheMeridian(iLatO) * BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLatO))
            Ai = axis * (1 - n + 5 * (n ^ 2 - n ^ 3) / 4 + 81 * (n ^ 4 - n ^ 5) / 64) * Math.PI / 180
            Bi = 3 * axis * (n - n ^ 2 + 7 * (n ^ 3 - n ^ 4) / 8 + 55 * n ^ 5 / 64) / 2
            Ci = 15 * axis * (n ^ 2 - n ^ 3 + 3 * (n ^ 4 - n ^ 5) / 4) / 16
            Di = 35 * axis * (n ^ 3 - n ^ 4 + 11 * n ^ 5 / 16) / 48
            Ei = 315 * axis * (n ^ 4 - n ^ 5) / 512
            Ro = iK0 * BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLatO) / Math.Tan(iLatO)
            'In the first term iLatF is in degrees, in the other terms iLatF is in radians.
            So = Ai * RadDeg(iLatO) - Bi * Math.Sin(2 * iLatO) + Ci * Math.Sin(4 * iLatO) - Di * Math.Sin(6 * iLatO) + Ei * Math.Sin(8 * iLatO)

            'From LL to EN
            s = Ai * RadDeg(point.Y) - Bi * Math.Sin(2 * point.Y) + Ci * Math.Sin(4 * point.Y) - Di * Math.Sin(6 * point.Y) + Ei * Math.Sin(8 * point.Y)
            m = (s - So)
            Mi = iK0 * (m + A * m ^ 3)
            r = Ro - Mi
            theta = (point.X - iLonO) * Math.Sin(iLatO)

            TmpResult.X = iEastO + r * Math.Sin(theta)
            TmpResult.Y = iNorthO + Mi + r * Math.Sin(theta) * Math.Tan(theta / 2)
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim axis As Double
        Dim TmpResult As New Point3D
        Dim A, Ai, Bi, Ci, Di, Ei, Ro, So, n, Si, m, Mi, DSi, lati, Ri, thetai As Double

        Try
            'Ellipsoid costants
            axis = BaseEllipsoid.SemiMayorAxis
            'First compute constants for the projection
            n = BaseEllipsoid.Flattening / (2 - BaseEllipsoid.Flattening)
            A = 1 / (6 * BaseEllipsoid.GetRadiuosOfCurvatureInTheMeridian(iLatO) * BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLatO))
            Ai = axis * (1 - n + 5 * (n ^ 2 - n ^ 3) / 4 + 81 * (n ^ 4 - n ^ 5) / 64) * Math.PI / 180
            Bi = 3 * axis * (n - n ^ 2 + 7 * (n ^ 3 - n ^ 4) / 8 + 55 * n ^ 5 / 64) / 2
            Ci = 15 * axis * (n ^ 2 - n ^ 3 + 3 * (n ^ 4 - n ^ 5) / 4) / 16
            Di = 35 * axis * (n ^ 3 - n ^ 4 + 11 * n ^ 5 / 16) / 48
            Ei = 315 * axis * (n ^ 4 - n ^ 5) / 512
            Ro = iK0 * BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLatO) / Math.Tan(iLatO)
            'In the first term iLatF is in degrees, in the other terms iLatF is in radians.
            So = Ai * RadDeg(iLatO) - Bi * Math.Sin(2 * iLatO) + Ci * Math.Sin(4 * iLatO) - Di * Math.Sin(6 * iLatO) + Ei * Math.Sin(8 * iLatO)

            'From EN to LL
            thetai = Math.Atan((point.X - iEastO) / (Ro - (point.Y - iNorthO)))
            Ri = Math.Abs(((point.X - iEastO) ^ 2 + (Ro - (point.Y - iNorthO)) ^ 2) ^ 0.5)
            If n < 0 Then Ri = -Ri
            Mi = Ro - Ri
            m = Mi
            For i = 0 To 10
                m = m - (Mi - iK0 * m - iK0 * A * m ^ 3) / (-iK0 - 3 * iK0 * A * m ^ 2)
            Next
            lati = iLatO + m / Ai * (Math.PI / 180)
            For i = 0 To 20
                lati = lati + (m + So - (Ai * lati * (180 / Math.PI) - Bi * Math.Sin(2 * lati) + Ci * Math.Sin(4 * lati) - Di * Math.Sin(6 * lati) + Ei * Math.Sin(8 * lati))) / Ai * (Math.PI / 180)
            Next
            'In the first term lati is in degrees, in the other terms lati is in radians.
            Si = Ai * RadDeg(lati) - Bi * Math.Sin(2 * lati) + Ci * Math.Sin(4 * lati) - Di * Math.Sin(6 * lati) + Ei * Math.Sin(8 * lati)
            DSi = Ai * (180 / Math.PI) - 2 * Bi * Math.Cos(2 * lati) + 4 * Ci * Math.Cos(4 * lati) - 6 * Di * Math.Cos(6 * lati) + 8 * Ei * Math.Cos(8 * lati)

            TmpResult.X = iLonO + thetai / Math.Sin(iLatO)
            TmpResult.Y = lati - ((m + So - Si) / (-DSi))
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
        xproj.Add(New XElement(x.OriginScale, iK0))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Latitude of Natural Origin", LatitudeOfNaturalOrigin, ParamType.LatLong, True),
            New ParamNameValue("Longitude of Natural Origin", LongitudeOfNaturalOrigin, ParamType.LatLong, False),
            New ParamNameValue("Scale factor at Natural Origin", ScaleFactorAtNaturalOrigin, ParamType.ScaleFactor, False),
            New ParamNameValue("False Easting at Natural Origin", FalseEasting, ParamType.EastNorth, False),
            New ParamNameValue("False Northing at Natural Origin", FalseNorthing, ParamType.EastNorth, True)
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
        H2200 = H2200 & FormatDMS(LongitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude), (d.m.s. N/E)
        Dim H2301 As String = ps.H2301
        H2301 = H2301 & FormatDMS(LatitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2301 = H2301 & FormatDMS(LongitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
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
        H2401 = H2401 & FormatNumber(ScaleFactorAtNaturalOrigin, 10).PadLeft(20, " "c)
        H2401 = H2401.PadRight(80, " "c)
        ret.Add(H2401)

        'Latitude/Longitude at which scale factor is defined
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class