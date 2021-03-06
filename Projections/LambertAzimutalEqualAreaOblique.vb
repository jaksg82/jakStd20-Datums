﻿Imports MathExt
Imports StringFormat

Public Class LambertAzimutalEqualAreaOblique
    Inherits Projections

    Private iLatO, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.LambertAzimutalEqualArea
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
        FullName = longName
        ShortName = compactName
        iLatO = latOfOrigin
        iLonO = lonOfOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim axis, ecc As Double
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity
            'From LL to EN
            Dim B, D, Rq, Beta, BetaO, q, qP, qO, latP As Double
            latP = Math.PI / 2
            q = (1 - ecc ^ 2) * ((Math.Sin(point.Y) / (1 - ecc ^ 2 * Math.Sin(point.Y) ^ 2)) - ((1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(point.Y)) / (1 + ecc * Math.Sin(point.Y)))))
            qO = (1 - ecc ^ 2) * ((Math.Sin(iLatO) / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)) - ((1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO)))))
            qP = (1 - ecc ^ 2) * ((Math.Sin(latP) / (1 - ecc ^ 2 * Math.Sin(latP) ^ 2)) - ((1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(latP)) / (1 + ecc * Math.Sin(latP)))))
            Beta = Math.Asin(q / qP)
            BetaO = Math.Asin(qO / qP)
            Rq = axis * (qP / 2) ^ 0.5
            D = axis * (Math.Cos(iLatO) / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2) ^ 0.5) / (Rq * Math.Cos(BetaO))
            B = Rq * (2 / (1 + Math.Sin(BetaO) * Math.Sin(Beta) + (Math.Cos(BetaO) * Math.Cos(Beta) * Math.Cos(point.X - iLonO)))) ^ 0.5

            TmpResult.X = iEastO + ((B * D) * (Math.Cos(Beta) * Math.Sin(point.X - iLonO)))
            TmpResult.Y = iNorthO + (B / D) * ((Math.Cos(BetaO) * Math.Sin(Beta)) - (Math.Sin(BetaO) * Math.Cos(Beta) * Math.Cos(point.X - iLonO)))
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

            'From EN to LL
            Dim D, Rq, BetaO, Beta1, C, p, qP, qO, latP As Double
            latP = Math.PI / 2
            qO = (1 - ecc ^ 2) * ((Math.Sin(iLatO) / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)) - ((1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO)))))
            qP = (1 - ecc ^ 2) * ((Math.Sin(latP) / (1 - ecc ^ 2 * Math.Sin(latP) ^ 2)) - ((1 / (2 * ecc)) * Math.Log((1 - ecc * Math.Sin(latP)) / (1 + ecc * Math.Sin(latP)))))
            BetaO = Math.Asin(qO / qP)
            Rq = axis * (qP / 2) ^ 0.5
            D = axis * (Math.Cos(iLatO) / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2) ^ 0.5) / (Rq * Math.Cos(BetaO))
            p = (((point.X - iEastO) / D) ^ 2 + (D * (point.Y - iNorthO)) ^ 2) ^ 0.5
            C = 2 * Math.Asin(p / (2 * Rq))
            Beta1 = Math.Asin((Math.Cos(C) * Math.Sin(BetaO)) + ((D * (point.Y - iNorthO) * Math.Sin(C) * Math.Cos(BetaO)) / p))

            'Splitted factors for the latitude computation
            Dim E1, E2, E3 As Double
            E1 = ecc ^ 2 / 3 + 31 * ecc ^ 4 / 180 + 517 * ecc ^ 6 / 5040
            E2 = 23 * ecc ^ 4 / 360 + 251 * ecc ^ 6 / 3780
            E3 = 761 * ecc ^ 6 / 45360

            TmpResult.Y = Beta1 + (E1 * Math.Sin(2 * Beta1)) + (E2 * Math.Sin(4 * Beta1)) + (E3 * Math.Sin(6 * Beta1))
            TmpResult.X = iLonO + Math.Atan((point.X - iEastO) * Math.Sin(C) / (D * p * Math.Cos(BetaO) * Math.Cos(C) - D ^ 2 * (point.Y - iNorthO) * Math.Sin(BetaO) * Math.Sin(C)))
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
            New ParamNameValue("Latitude of projection center", LatitudeOfOrigin, ParamType.LatLong, True),
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