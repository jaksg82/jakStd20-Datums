Imports MathExt
Imports StringFormat

Public Class AlbersConicEqualArea
    Inherits Projections

    Private iLatO, iLat1, iLat2, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.AlbersEqualArea
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
                                            eastAtFalseOrigin As Double, northAtFalseOrigin As Double)
        Me.FullName = longName
        Me.ShortName = compactName
        iLatO = latOfFalseOrigin
        iLonO = lonOfFalseOrigin
        If Math.Abs(latOfFirstParallel) > Math.Abs(latOfSecondParallel) Then
            iLat2 = latOfFirstParallel
            iLat1 = latOfSecondParallel
        Else
            iLat1 = latOfFirstParallel
            iLat2 = latOfSecondParallel
        End If
        iEastO = eastAtFalseOrigin
        iNorthO = northAtFalseOrigin
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
            Dim C, n, m1, m2, a, a1, a2, aO, p, pO, theta As Double
            m1 = Math.Cos(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
            m2 = Math.Cos(iLat2) / (1 - ecc ^ 2 * Math.Sin(iLat2) ^ 2) ^ 0.5
            a = (1 - ecc ^ 2) * ((Math.Sin(point.Y) / (1 - ecc ^ 2 * Math.Sin(point.Y) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(point.Y)) / (1 + ecc * Math.Sin(point.Y))))
            aO = (1 - ecc ^ 2) * ((Math.Sin(iLatO) / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))))
            a1 = (1 - ecc ^ 2) * ((Math.Sin(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLat1)) / (1 + ecc * Math.Sin(iLat1))))
            a2 = (1 - ecc ^ 2) * ((Math.Sin(iLat2) / (1 - ecc ^ 2 * Math.Sin(iLat2) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLat2)) / (1 + ecc * Math.Sin(iLat2))))
            n = (m1 ^ 2 - m2 ^ 2) / (a2 - a1)
            C = m1 ^ 2 + (n * a1)
            theta = n * (point.X - iLonO)
            p = (axis * (C - n * a) ^ 0.5) / n
            pO = (axis * (C - n * aO) ^ 0.5) / n

            'Return the results
            TmpResult.X = iEastO + (p * Math.Sin(theta))
            TmpResult.Y = iNorthO + pO - (p * Math.Cos(theta))
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
            Dim m1, m2, a1, a2, aO, ai, Beta1, p, n, C, pO, Theta As Double
            m1 = Math.Cos(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
            m2 = Math.Cos(iLat2) / (1 - ecc ^ 2 * Math.Sin(iLat2) ^ 2) ^ 0.5
            a1 = (1 - ecc ^ 2) * ((Math.Sin(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLat1)) / (1 + ecc * Math.Sin(iLat1))))
            a2 = (1 - ecc ^ 2) * ((Math.Sin(iLat2) / (1 - ecc ^ 2 * Math.Sin(iLat2) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLat2)) / (1 + ecc * Math.Sin(iLat2))))
            aO = (1 - ecc ^ 2) * ((Math.Sin(iLatO) / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)) - (1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))))
            n = (m1 ^ 2 - m2 ^ 2) / (a2 - a1)
            C = m1 ^ 2 + (n * a1)
            pO = (axis * (C - n * aO) ^ 0.5) / n
            Theta = Math.Atan((point.X - iEastO) / (pO - (point.Y - iNorthO)))
            p = ((point.X - iEastO) ^ 2 + (pO - (point.Y - iNorthO)) ^ 2) ^ 0.5
            ai = (C - (p ^ 2 * n ^ 2 / axis ^ 2)) / n
            Beta1 = Math.Asin(ai / (1 - ((1 - ecc ^ 2) / (2 * ecc)) * Math.Log((1 - ecc) / (1 + ecc))))

            'Splitted factors for the latitude computation
            Dim E1, E2, E3 As Double
            E1 = ecc ^ 2 / 3 + 31 * ecc ^ 4 / 180 + 517 * ecc ^ 6 / 5040
            E2 = 23 * ecc ^ 4 / 360 + 251 * ecc ^ 6 / 3780
            E3 = 761 * ecc ^ 6 / 45360

            'Return the results
            TmpResult.X = iLonO + (Theta / n)
            TmpResult.Y = Beta1 + (E1 * Math.Sin(2 * Beta1)) + (E2 * Math.Sin(4 * Beta1)) + (E3 * Math.Sin(6 * Beta1))
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
            New ParamNameValue("Easting at False Origin", EastingAtFalseOrigin, ParamType.EastNorth, False),
            New ParamNameValue("Northing at False Origin", NorthingAtFalseOrigin, ParamType.EastNorth, True)
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

        'Latitude/Longitude at which scale factor is defined 2*(d.m.s)
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfFirstParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfFalseOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class