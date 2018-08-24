Imports jakStd20_MathExt
Imports jakStd20_StringFormat

Public Class StereographicOblique
    Inherits Projections

    Dim iLatO, iK0, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.StereographicOblique
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
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            Dim ecc As Double
            ecc = BaseEllipsoid.Eccentricity
            'Projection costants
            Dim R, n, c, sinXo, w1, w2, S1, S2, Xo, Vo As Double
            R = (BaseEllipsoid.GetRadiuosOfCurvatureInTheMeridian(iLatO) * BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLatO)) ^ 0.5
            n = (1 + ((ecc ^ 2 * Math.Cos(iLatO) ^ 4) / (1 - ecc ^ 2))) ^ 0.5
            S1 = (1 + Math.Sin(iLatO)) / (1 - Math.Sin(iLatO))
            S2 = (1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))
            w1 = (S1 * S2 ^ ecc) ^ n
            sinXo = (w1 - 1) / (w1 + 1)
            c = (n + Math.Sin(iLatO)) * (1 - sinXo) / ((n - Math.Sin(iLatO)) * (1 + sinXo))
            w2 = c * w1
            Xo = Math.Asin((w2 - 1) / (w2 + 1))
            Vo = iLonO

            'From LL to EN
            Dim Xp, Vp, w, Sa, Sb, B As Double
            Sa = (1 + Math.Sin(point.Y)) / (1 - Math.Sin(point.Y))
            Sb = (1 - ecc * Math.Sin(point.Y)) / (1 + ecc * Math.Sin(point.Y))
            w = c * (Sa * Sb ^ ecc) ^ n
            Xp = Math.Asin((w - 1) / (w + 1))
            Vp = n * (point.X - Vo) + Vo
            B = (1 + Math.Sin(Xp) * Math.Sin(Xo) + Math.Cos(Xp) * Math.Cos(Xo) * Math.Cos(Vp - Vo))

            TmpResult.X = iEastO + 2 * R * iK0 * Math.Cos(Xp) * Math.Sin(Vp - Vo) / B
            TmpResult.Y = iNorthO + 2 * R * iK0 * (Math.Sin(Xp) * Math.Cos(Xo) - Math.Cos(Xp) * Math.Sin(Xo) * Math.Cos(Vp - Vo)) / B
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
            Dim ecc As Double
            ecc = BaseEllipsoid.Eccentricity
            'Projection costants
            Dim R, n, c, sinXo, w1, w2, S1, S2, Xo, Vo As Double
            R = (BaseEllipsoid.GetRadiuosOfCurvatureInTheMeridian(iLatO) * BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(iLatO)) ^ 0.5
            n = (1 + ((ecc ^ 2 * Math.Cos(iLatO) ^ 4) / (1 - ecc ^ 2))) ^ 0.5
            S1 = (1 + Math.Sin(iLatO)) / (1 - Math.Sin(iLatO))
            S2 = (1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))
            w1 = (S1 * S2 ^ ecc) ^ n
            sinXo = (w1 - 1) / (w1 + 1)
            c = (n + Math.Sin(iLatO)) * (1 - sinXo) / ((n - Math.Sin(iLatO)) * (1 + sinXo))
            w2 = c * w1
            Xo = Math.Asin((w2 - 1) / (w2 + 1))
            Vo = iLonO

            'From EN to LL
            Dim Xp, Vp, g, h, i, j As Double
            g = 2 * R * iK0 * Math.Tan(Math.PI / 4 - Xo / 2)
            h = 4 * R * iK0 * Math.Tan(Xo) + g
            i = Math.Atan((point.X - iEastO) / (h + (point.Y - iNorthO)))
            j = Math.Atan((point.X - iEastO) / (g - (point.Y - iNorthO))) - i
            Xp = Xo + 2 * Math.Atan(((point.Y - iNorthO) - (point.X - iEastO) * Math.Tan(j / 2)) / (2 * R * iK0))
            Vp = j + 2 * i + Vo
            'Iteration for the latitude
            Dim isoLat0, isoLat1, lat0, lat1 As Double
            isoLat0 = 0.5 * Math.Log((1 + Math.Sin(Xp)) / (c * (1 - Math.Sin(Xp)))) / n
            lat0 = 2 * Math.Atan(Math.E ^ isoLat0) - Math.PI / 2
            Do
                isoLat1 = Math.Log((Math.Tan(lat0 / 2 + Math.PI / 4)) * ((1 - ecc * Math.Sin(lat0)) / (1 + ecc * Math.Sin(lat0))) ^ (ecc / 2))
                lat1 = lat0 - (isoLat1 - isoLat0) * Math.Cos(lat0) * (1 - ecc ^ 2 * Math.Sin(lat0) ^ 2) / (1 - ecc ^ 2)
                If Math.Abs(lat0 - lat1) < 0.0000000001 Then Exit Do
                lat0 = lat1
            Loop

            TmpResult.X = (Vp - Vo) / n + Vo
            TmpResult.Y = lat1
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
        xproj.Add(New XElement(x.ScaleDifference, iK0))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue)

        tmpList.Add(New ParamNameValue("Latitude of Natural Origin", LatitudeOfNaturalOrigin, ParamType.LatLong, True))
        tmpList.Add(New ParamNameValue("Longitude of Natural Origin", LongitudeOfNaturalOrigin, ParamType.LatLong, False))
        tmpList.Add(New ParamNameValue("Scale factor at Natural Origin", ScaleFactorAtNaturalOrigin, ParamType.ScaleFactor, False))
        tmpList.Add(New ParamNameValue("False Easting at Natural Origin", FalseEasting, ParamType.EastNorth, False))
        tmpList.Add(New ParamNameValue("False Northing at Natural Origin", FalseNorthing, ParamType.EastNorth, True))

        Return tmpList

    End Function

    Public Overrides Function ToP190Header() As List(Of String)
        Dim ret As New List(Of String)
        Dim ps As New ProjStrings

        'Projection Code and Name
        Dim H1800 As String = ps.H1800
        H1800 = H1800 & " 010" & If(FullName.Length > 44, FullName.Substring(0, 44), FullName.PadRight(44, " "c))
        ret.Add(H1800)

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