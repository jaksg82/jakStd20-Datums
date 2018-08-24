Imports jakStd20_MathExt
Imports jakStd20_StringFormat

Public Class LambertConic1SP
    Inherits Projections

    Dim iLatO, iK0, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.LambertConicConformal1SP
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
        Dim a, e As Double
        Dim TmpResult As New Point3D
        Dim Rf, mf, Ff, t0, tf, n, R0, theta As Double

        Try
            'Ellipsoid costants
            a = BaseEllipsoid.SemiMayorAxis
            e = BaseEllipsoid.Eccentricity
            'From LL to EN
            mf = Math.Cos(iLatO) / Math.Sqrt(1 - e ^ 2 * Math.Sin(iLatO) ^ 2)
            t0 = Math.Tan(Math.PI / 4 - point.Y / 2) / ((1 - e * Math.Sin(point.Y)) / (1 + e * Math.Sin(point.Y))) ^ (e / 2)
            tf = Math.Tan(Math.PI / 4 - iLatO / 2) / ((1 - e * Math.Sin(iLatO)) / (1 + e * Math.Sin(iLatO))) ^ (e / 2)
            n = Math.Sin(iLatO)
            Ff = mf / (n * tf ^ n)
            Rf = a * Ff * tf ^ n * iK0
            R0 = a * Ff * t0 ^ n * iK0
            theta = n * (point.X - iLonO)
            TmpResult.X = iEastO + R0 * Math.Sin(theta)
            TmpResult.Y = iNorthO + Rf - R0 * Math.Cos(theta)
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim a, e As Double
        Dim TmpResult As New Point3D
        Dim Rf, mf, tf, n, Ri, Ff, Ti, thetai, tmplat As Double

        Try
            'Ellipsoid costants
            a = BaseEllipsoid.SemiMayorAxis
            e = BaseEllipsoid.Eccentricity
            'From EN to LL
            mf = Math.Cos(iLatO) / Math.Sqrt(1 - e ^ 2 * Math.Sin(iLatO) ^ 2)
            tf = Math.Tan(Math.PI / 4 - iLatO / 2) / ((1 - e * Math.Sin(iLatO)) / (1 + e * Math.Sin(iLatO))) ^ (e / 2)
            n = Math.Sin(iLatO)
            Ff = mf / (n * tf ^ n)
            Rf = a * Ff * tf ^ n * iK0
            Ri = Math.Abs(((point.X - iEastO) ^ 2 + (Rf - (point.Y - iNorthO)) ^ 2) ^ 0.5)
            If n < 0 Then Ri = -Ri
            Ti = (Ri / (a * iK0 * Ff)) ^ (1 / n)
            tmplat = Math.PI / 2 - 2 * Math.Atan(Ti)
            For i = 0 To 100
                tmplat = Math.PI / 2 - 2 * Math.Atan(Ti * ((1 - e * Math.Sin(tmplat)) / (1 + e * Math.Sin(tmplat))) ^ (e / 2))
            Next
            thetai = Math.Atan((point.X - iEastO) / (Rf - (point.Y - iNorthO)))
            TmpResult.X = thetai / n + iLonO
            TmpResult.Y = tmplat
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
        H1800 = H1800 & " 005" & If(FullName.Length > 44, FullName.Substring(0, 44), FullName.PadRight(44, " "c))
        ret.Add(H1800)

        'Latitude of standard parallel(s), (d.m.s. N/S)
        Dim H2100 As String = ps.H2100
        H2100 = H2100 & FormatDMS(LatitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100.PadRight(80, " "c)
        ret.Add(H2100)

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