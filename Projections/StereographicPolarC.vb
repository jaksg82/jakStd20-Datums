Imports jakStd20_MathExt
Imports jakStd20_StringFormat

Public Class StereographicPolarC
    Inherits Projections

    Dim iLat1, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.StereographicPolarC
        End Get
    End Property

    ReadOnly Property LongitudeOfOrigin As Double
        Get
            Return iLonO
        End Get
    End Property

    ReadOnly Property LatitudeOfStandardParallel As Double
        Get
            Return iLat1
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

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfStandardParallel As Double, lonOfOrigin As Double,
                                             falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLonO = lonOfOrigin
        iLat1 = latOfStandardParallel
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            Dim axis, ecc As Double
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity

            'From LL to EN
            Dim dE, dN, p, pf, t, tf, mf, ko As Double
            If iLat1 < 0 Then
                'South pole case
                tf = Math.Tan(Math.PI / 4 + iLat1 / 2) / (((1 + ecc * Math.Sin(iLat1)) / (1 - ecc * Math.Sin(iLat1))) ^ (ecc / 2))
                mf = Math.Cos(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
                ko = mf * (((1 + ecc) ^ (1 + ecc) * (1 - ecc) ^ (1 - ecc)) ^ 0.5) / (2 * tf)
                t = Math.Tan(Math.PI / 4 + point.Y / 2) / (((1 + ecc * Math.Sin(point.Y)) / (1 - ecc * Math.Sin(point.Y))) ^ (ecc / 2))
                pf = axis * mf
                p = pf * t / tf
                dE = p * Math.Sin(point.X - iLonO)
                dN = -pf + p * Math.Cos(point.X - iLonO)
            Else
                'North pole case
                tf = Math.Tan(Math.PI / 4 - iLat1 / 2) * (((1 + ecc * Math.Sin(iLat1)) / (1 - ecc * Math.Sin(iLat1))) ^ (ecc / 2))
                mf = Math.Cos(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
                ko = mf * (((1 + ecc) ^ (1 + ecc) * (1 - ecc) ^ (1 - ecc)) ^ 0.5) / (2 * tf)
                t = Math.Tan(Math.PI / 4 - point.Y / 2) * (((1 + ecc * Math.Sin(point.Y)) / (1 - ecc * Math.Sin(point.Y))) ^ (ecc / 2))
                pf = axis * mf
                p = pf * t / tf
                dE = p * Math.Sin(point.X - iLonO)
                dN = pf - p * Math.Cos(point.X - iLonO)
            End If

            TmpResult.X = dE + iEastO
            TmpResult.Y = dN + iNorthO
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
            Dim X, p, pf, t, E1, E2, E3, E4, tf, mf, ko As Double
            If iLat1 < 0 Then
                'South pole case
                tf = Math.Tan(Math.PI / 4 + iLat1 / 2) / (((1 + ecc * Math.Sin(iLat1)) / (1 - ecc * Math.Sin(iLat1))) ^ (ecc / 2))
                mf = Math.Cos(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
                ko = mf * (((1 + ecc) ^ (1 + ecc) * (1 - ecc) ^ (1 - ecc)) ^ 0.5) / (2 * tf)
                pf = axis * mf
                p = ((point.X - iEastO) ^ 2 + (point.Y - iNorthO + pf) ^ 2) ^ 0.5
                t = p * tf / pf
                X = 2 * Math.Atan(t) - Math.PI / 2
            Else
                'North pole case
                tf = Math.Tan(Math.PI / 4 - iLat1 / 2) * (((1 + ecc * Math.Sin(iLat1)) / (1 - ecc * Math.Sin(iLat1))) ^ (ecc / 2))
                mf = Math.Cos(iLat1) / (1 - ecc ^ 2 * Math.Sin(iLat1) ^ 2) ^ 0.5
                ko = mf * (((1 + ecc) ^ (1 + ecc) * (1 - ecc) ^ (1 - ecc)) ^ 0.5) / (2 * tf)
                pf = axis * mf
                p = ((point.X - iEastO) ^ 2 + (point.Y - iNorthO - pf) ^ 2) ^ 0.5
                t = p * tf / pf
                X = Math.PI / 2 - 2 * Math.Atan(t)
            End If
            E1 = ecc ^ 2 / 2 + 5 * ecc ^ 4 / 24 + ecc ^ 6 / 12 + 13 * ecc ^ 8 / 360
            E2 = 7 * ecc ^ 4 / 48 + 29 * ecc ^ 6 / 240 + 811 * ecc ^ 8 / 11520
            E3 = 7 * ecc ^ 6 / 120 + 81 * ecc ^ 8 / 1120
            E4 = 4279 * ecc ^ 8 / 161280

            If point.X = iEastO Then
                TmpResult.X = iLonO
            Else
                If iLat1 < 0 Then
                    'South pole case
                    TmpResult.X = iLonO + Math.Atan2(point.X - iEastO, point.Y - iNorthO + pf)
                Else
                    'North pole case
                    TmpResult.X = iLonO + Math.Atan2(point.X - iEastO, iNorthO + pf - point.Y)
                End If
            End If
            TmpResult.Y = X + E1 * Math.Sin(2 * X) + E2 * Math.Sin(4 * X) + E3 * Math.Sin(6 * X) + E4 * Math.Sin(8 * X)
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
        xproj.Add(New XElement(x.OriginLongitude, iLonO))
        xproj.Add(New XElement(x.FirstLatitude, iLat1))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue)

        tmpList.Add(New ParamNameValue("Latitude of standard parallel", LatitudeOfStandardParallel, ParamType.LatLong, True))
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
        H1800 = H1800 & " 010" & If(FullName.Length > 44, FullName.Substring(0, 44), FullName.PadRight(44, " "c))
        ret.Add(H1800)

        'Latitude of standard parallel(s), (d.m.s. N/S)
        Dim H2100 As String = ps.H2100
        H2100 = H2100 & FormatDMS(LatitudeOfStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        'H2100 = H2100 & JakStringFormats.FormatDMS(LatitudeOfSecondParallel, JakStringFormats.DmsFormat.UkooaDMS, JakStringFormats.DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100.PadRight(80, " "c)
        ret.Add(H2100)

        'Grid origin (Latitude, Longitude), (d.m.s. N/E)
        Dim H2301 As String = ps.H2301
        H2301 = H2301 & FormatDMS(If(iLat1 < 0, -90, 90), DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
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
        H2402 = H2402 & FormatDMS(LatitudeOfStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class