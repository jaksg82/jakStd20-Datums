Imports MathExt
Imports StringFormat

Public Class LambertConicBelgium
    Inherits Projections

    Private iLatO, iLat1, iLat2, iLonO, iEastO, iNorthO As Double
    Private CorrA As Double = DegRad(29.2985 / 3600)

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.LambertConicConformalBelgium
        End Get
    End Property

    ReadOnly Property LatitudeOfFalseOrigin As Double
        Get
            Return iLatO
        End Get
    End Property

    ReadOnly Property LongitudeOfFalseOrigin As Double
        Get
            Return iLonO
        End Get
    End Property

    ReadOnly Property LatitudeOfFirstParallel As Double
        Get
            Return iLat1
        End Get
    End Property

    ReadOnly Property LatitudeOfSecondParallel As Double
        Get
            Return iLat2
        End Get
    End Property

    ReadOnly Property EastingAtFalseOrigin As Double
        Get
            Return iEastO
        End Get
    End Property

    ReadOnly Property NorthingAtFalseOrigin As Double
        Get
            Return iNorthO
        End Get
    End Property

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfFalseOrigin As Double, lonOfFalseOrigin As Double,
                                            latOfFirstParallel As Double, latOfSecondParallel As Double,
                                            falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLatO = latOfFalseOrigin
        iLonO = lonOfFalseOrigin
        iLat1 = latOfFirstParallel
        iLat2 = latOfSecondParallel
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim a, e As Double
        Dim TmpResult, tmpIter As New Point3D
        Dim f, Rf, m1, m2, t0, t1, t2, tf, n, r, theta As Double

        Try
            'Ellipsoid costants
            a = BaseEllipsoid.SemiMayorAxis
            e = BaseEllipsoid.Eccentricity
            'From LL to EN
            m1 = Math.Cos(iLat1) / (1 - e ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
            m2 = Math.Cos(iLat2) / (1 - e ^ 2 * Math.Sin(iLat2) ^ 2) ^ 0.5
            t0 = Math.Tan(Math.PI / 4 - point.Y / 2) / ((1 - e * Math.Sin(point.Y)) / (1 + e * Math.Sin(point.Y))) ^ (e / 2)
            t1 = Math.Tan(Math.PI / 4 - iLat1 / 2) / ((1 - e * Math.Sin(iLat1)) / (1 + e * Math.Sin(iLat1))) ^ (e / 2)
            t2 = Math.Tan(Math.PI / 4 - iLat2 / 2) / ((1 - e * Math.Sin(iLat2)) / (1 + e * Math.Sin(iLat2))) ^ (e / 2)
            tf = Math.Tan(Math.PI / 4 - iLatO / 2) / ((1 - e * Math.Sin(iLatO)) / (1 + e * Math.Sin(iLatO))) ^ (e / 2)
            n = (Math.Log(m1) - Math.Log(m2)) / (Math.Log(t1) - Math.Log(t2))
            f = m1 / (n * t1 ^ n)
            r = a * f * t0 ^ n
            Rf = a * f * tf ^ n
            theta = n * (point.X - iLonO)
            'Return the results
            TmpResult.X = iEastO + r * Math.Sin(theta - CorrA)
            TmpResult.Y = iNorthO + Rf - r * Math.Cos(theta - CorrA)
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim a, e As Double
        Dim TmpResult As New Point3D
        Dim f, Rf, m1, m2, t0, t1, t2, tf, n, Ri, thetainv, tmplat As Double

        Try
            'Ellipsoid costants
            a = BaseEllipsoid.SemiMayorAxis
            e = BaseEllipsoid.Eccentricity
            'From EN to LL
            m1 = Math.Cos(iLat1) / Math.Sqrt(1 - e ^ 2 * Math.Sin(iLat1) ^ 2)
            m2 = Math.Cos(iLat2) / Math.Sqrt(1 - e ^ 2 * Math.Sin(iLat2) ^ 2)
            t1 = Math.Tan(Math.PI / 4 - iLat1 / 2) / ((1 - e * Math.Sin(iLat1)) / (1 + e * Math.Sin(iLat1))) ^ (e / 2)
            t2 = Math.Tan(Math.PI / 4 - iLat2 / 2) / ((1 - e * Math.Sin(iLat2)) / (1 + e * Math.Sin(iLat2))) ^ (e / 2)
            tf = Math.Tan(Math.PI / 4 - iLatO / 2) / ((1 - e * Math.Sin(iLatO)) / (1 + e * Math.Sin(iLatO))) ^ (e / 2)
            n = (Math.Log(m1) - Math.Log(m2)) / (Math.Log(t1) - Math.Log(t2))
            f = m1 / (n * t1 ^ n)
            Rf = a * f * tf ^ n
            If n > 0 Then
                Ri = Math.Abs(((point.X - iEastO) ^ 2 + (Rf - (point.Y - iNorthO)) ^ 2) ^ 0.5)
            Else
                Ri = -Math.Abs(((point.X - iEastO) ^ 2 + (Rf - (point.Y - iNorthO)) ^ 2) ^ 0.5)
            End If
            t0 = (Ri / (a * f)) ^ (1 / n)
            tmplat = Math.PI / 2 - 2 * Math.Atan(t0)
            For i = 0 To 100
                tmplat = Math.PI / 2 - 2 * Math.Atan(t0 * ((1 - e * Math.Sin(tmplat)) / (1 + e * Math.Sin(tmplat))) ^ (e / 2))
            Next
            thetainv = Math.Atan((point.X - iEastO) / (Rf - (point.Y - iNorthO)))
            'Return the results
            TmpResult.X = (thetainv + CorrA) / n + iLonO
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
        xproj.Add(New XElement(x.FirstLatitude, iLat1))
        xproj.Add(New XElement(x.SecondLatitude, iLat2))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Latitude of False Origin", LatitudeOfFalseOrigin, ParamType.LatLong, True),
            New ParamNameValue("Longitude of False Origin", LongitudeOfFalseOrigin, ParamType.LatLong, False),
            New ParamNameValue("Latitude of first parallel", LatitudeOfFirstParallel, ParamType.LatLong, True),
            New ParamNameValue("Latitude of second parallel", LatitudeOfSecondParallel, ParamType.LatLong, True),
            New ParamNameValue("Easting at false Origin", EastingAtFalseOrigin, ParamType.EastNorth, False),
            New ParamNameValue("Northing at false Origin", NorthingAtFalseOrigin, ParamType.EastNorth, True)
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

        'Latitude of standard parallel(s), (d.m.s. N/S)
        Dim H2100 As String = ps.H2100
        H2100 = H2100 & FormatDMS(LatitudeOfFirstParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100 & FormatDMS(LatitudeOfSecondParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100.PadRight(80, " "c)
        ret.Add(H2100)

        'Longitude of central meridian, (d.m.s. E/W)
        Dim H2200 As String = ps.H2200
        H2200 = H2200 & FormatDMS(LongitudeOfFalseOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude), (d.m.s. N/E)
        Dim H2301 As String = ps.H2301
        H2301 = H2301 & FormatDMS(LatitudeOfFalseOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2301 = H2301 & FormatDMS(LongitudeOfFalseOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2301 = H2301.PadRight(80, " "c)
        ret.Add(H2301)

        'Grid co-ordinates at grid origin (E,N)
        Dim H2302 As String = ps.H2302
        H2302 = H2302 & FormatMetric(EastingAtFalseOrigin, MetricSign.Suffix, 2, False).PadLeft(20, " "c)
        H2302 = H2302 & FormatMetric(NorthingAtFalseOrigin, MetricSign.Suffix, 2, True).PadLeft(20, " "c)
        H2302 = H2302.PadRight(80, " "c)
        ret.Add(H2302)

        'Scale factor
        Dim H2401 As String = ps.H2401
        H2401 = H2401 & FormatNumber(1.0, 10).PadLeft(20, " "c)
        H2401 = H2401.PadRight(80, " "c)
        ret.Add(H2401)

        'Latitude/Longitude at which scale factor is defined
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfFirstParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfFalseOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class