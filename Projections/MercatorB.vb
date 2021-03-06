﻿Imports MathExt
Imports StringFormat

Public Class MercatorB
    Inherits Projections

    Private iLat1, iLonO, iK0, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.MercatorVariantB
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

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, lonOfOrigin As Double,
                   latOfStandardParallel As Double, falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLonO = lonOfOrigin
        iLat1 = latOfStandardParallel
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid
    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim a, e As Double
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            a = BaseEllipsoid.SemiMayorAxis
            e = BaseEllipsoid.Eccentricity
            'Projection costants
            iK0 = Math.Cos(iLat1) / (1 - (e ^ 2 * Math.Sin(iLat1) ^ 2)) ^ 0.5
            'From LL to EN
            TmpResult.X = iEastO + a * iK0 * (point.X - iLonO)
            TmpResult.Y = iNorthO + a * iK0 * Math.Log(Math.Tan(Math.PI / 4 + (point.Y / 2)) * ((1 - e * Math.Sin(point.Y)) / (1 + e * Math.Sin(point.Y))) ^ (e / 2))
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim a, e As Double
        Dim X, t, A1, A2, A3, A4 As Double
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            a = BaseEllipsoid.SemiMayorAxis
            e = BaseEllipsoid.Eccentricity
            'Projection costants
            iK0 = Math.Cos(iLat1) / (1 - (e ^ 2 * Math.Sin(iLat1) ^ 2)) ^ 0.5
            A1 = (e ^ 2 / 2) + (5 * e ^ 4 / 24) + (e ^ 6 / 12) + (13 * e ^ 8 / 360)
            A2 = (7 * e ^ 4 / 48) + (29 * e ^ 6 / 240) + (811 * e ^ 8 / 11520)
            A3 = (7 * e ^ 6 / 120) + (81 * e ^ 8 / 1120)
            A4 = (4279 * e ^ 8 / 161280)
            'From EN to LL
            t = Math.E ^ ((iNorthO - point.Y) / (a * iK0))
            X = Math.PI / 2 - 2 * Math.Atan(t)
            TmpResult.Y = X + A1 * Math.Sin(2 * X) + A2 * Math.Sin(4 * X) + A3 * Math.Sin(6 * X) + A4 * Math.Sin(8 * X)
            TmpResult.X = ((point.X - iEastO) / (a * iK0)) + iLonO
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
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Longitude of False Origin", LongitudeOfOrigin, ParamType.LatLong, False),
            New ParamNameValue("Latitude of standard parallel", LatitudeOfStandardParallel, ParamType.LatLong, True),
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
        H1800 = H1800 & " 007" & If(FullName.Length > 44, FullName.Substring(0, 44), FullName.PadRight(44, " "c))
        ret.Add(H1800)

        'Latitude of standard parallel(s), (d.m.s. N/S)
        Dim H2100 As String = ps.H2100
        H2100 = H2100 & FormatDMS(LatitudeOfStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100 & FormatDMS(-LatitudeOfStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100.PadRight(80, " "c)
        ret.Add(H2100)

        'Longitude of central meridian, (d.m.s. E/W)
        Dim H2200 As String = ps.H2200
        H2200 = H2200 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude, (d.m.s. N/E)
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
        H2402 = H2402 & FormatDMS(LatitudeOfStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class