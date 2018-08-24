Imports jakStd20_MathExt
Imports jakStd20_StringFormat

Public Class TransverseMercatorZoned
    Inherits Projections

    Dim iLatO, iK0, iLonI, iZoneW, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.TransverseMercatorZoned
        End Get
    End Property

    ReadOnly Property LatitudeOfNaturalOrigin As Double
        Get
            Return iLatO
        End Get
    End Property

    ReadOnly Property InitialLongitude As Double
        Get
            Return iLonI
        End Get
    End Property

    ReadOnly Property ZoneWidth As Double
        Get
            Return iZoneW
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

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, initialLon As Double, widthZones As Double, latOfNaturalOrigin As Double,
                                            scaleAtNaturalOrigin As Double, falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLatO = latOfNaturalOrigin
        iLonI = initialLon
        iZoneW = widthZones
        iK0 = scaleAtNaturalOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            Dim axis, ecc, flat As Double
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity
            flat = BaseEllipsoid.Flattening
            'Projection costants
            Dim n, BI, h1, h2, h3, h4, Z, LonO As Double
            Z = ((point.X - iLonI) / iZoneW)
            Z = Math.Floor(Z) + If(point.X > iLonI, 1, 0)
            LonO = (Z * iZoneW) - (iLonI + (iZoneW / 2))

            n = flat / (2 - flat)
            BI = (axis / (1 + n)) * (1 + n ^ 2 / 4 + n ^ 4 / 64)
            h1 = n / 2 - (2 / 3) * n ^ 2 + (5 / 16) * n ^ 3 + (41 / 180) * n ^ 4
            h2 = (13 / 48) * n ^ 2 - (3 / 5) * n ^ 3 + (557 / 1440) * n ^ 4
            h3 = (61 / 240) * n ^ 3 - (103 / 140) * n ^ 4
            h4 = (49561 / 161280) * n ^ 4

            Dim Mo, Qo, BetaO, xio0, xio1, xio2, xio3, xio4, xio As Double
            If iLatO = 0 Then
                Mo = 0
            ElseIf iLatO = (Math.PI / 2) Then
                Mo = BI * (Math.PI / 2)
            ElseIf iLatO = -(Math.PI / 2) Then
                Mo = BI * -(Math.PI / 2)
            Else
                Qo = ASinH(Math.Tan(iLatO)) - (ecc * ATanH(ecc * Math.Sin(iLatO)))
                BetaO = Math.Atan(Math.Sinh(Qo))
                xio0 = Math.Asin(Math.Sin(BetaO))
                xio1 = h1 * Math.Sin(2 * xio0)
                xio2 = h2 * Math.Sin(4 * xio0)
                xio3 = h3 * Math.Sin(6 * xio0)
                xio4 = h4 * Math.Sin(8 * xio0)
                xio = xio0 + xio1 + xio2 + xio3 + xio4
                Mo = BI * xio
            End If

            'From LL to EN
            Dim Q, Beta, xi0, xi1, xi2, xi3, xi4, xi, Eta0, Eta1, Eta2, Eta3, Eta4, Eta As Double
            Q = ASinH(Math.Tan(point.Y)) - (ecc * ATanH(ecc * Math.Sin(point.Y)))
            Beta = Math.Atan(SinH(Q))
            Eta0 = ATanH(Math.Cos(Beta) * Math.Sin(point.X - LonO))
            xi0 = ASin(Math.Sin(Beta) * CosH(Eta0))
            xi1 = h1 * Math.Sin(2 * xi0) * CosH(2 * Eta0)
            xi2 = h2 * Math.Sin(4 * xi0) * CosH(4 * Eta0)
            xi3 = h3 * Math.Sin(6 * xi0) * CosH(6 * Eta0)
            xi4 = h4 * Math.Sin(8 * xi0) * CosH(8 * Eta0)
            xi = xi0 + xi1 + xi2 + xi3 + xi4
            Eta1 = h1 * Math.Cos(2 * xi0) * SinH(2 * Eta0)
            Eta2 = h2 * Math.Cos(4 * xi0) * SinH(4 * Eta0)
            Eta3 = h3 * Math.Cos(6 * xi0) * SinH(6 * Eta0)
            Eta4 = h4 * Math.Cos(8 * xi0) * SinH(8 * Eta0)
            Eta = Eta0 + Eta1 + Eta2 + Eta3 + Eta4
            TmpResult.X = Z * 10 ^ 6 + iEastO + iK0 * BI * Eta
            TmpResult.Y = iNorthO + iK0 * (BI * xi - Mo)
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
            Dim axis, ecc, flat As Double
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity
            flat = BaseEllipsoid.Flattening
            'Projection costants
            Dim n, BI, h1, h2, h3, h4, h1i, h2i, h3i, h4i, Z, LonO As Double
            Z = Math.Floor((point.X - iEastO) / 10 ^ 6) + 1
            LonO = (Z * iZoneW) - (iLonI + (iZoneW / 2))

            n = flat / (2 - flat)
            BI = (axis / (1 + n)) * (1 + n ^ 2 / 4 + n ^ 4 / 64)
            h1 = n / 2 - (2 / 3) * n ^ 2 + (5 / 16) * n ^ 3 + (41 / 180) * n ^ 4
            h2 = (13 / 48) * n ^ 2 - (3 / 5) * n ^ 3 + (557 / 1440) * n ^ 4
            h3 = (61 / 240) * n ^ 3 - (103 / 140) * n ^ 4
            h4 = (49561 / 161280) * n ^ 4
            h1i = n / 2 - (2 / 3) * n ^ 2 + (37 / 96) * n ^ 3 - (1 / 360) * n ^ 4
            h2i = (1 / 48) * n ^ 2 + (1 / 15) * n ^ 3 - (437 / 1440) * n ^ 4
            h3i = (17 / 480) * n ^ 3 - (37 / 840) * n ^ 4
            h4i = (4397 / 161280) * n ^ 4

            Dim Mo, Qo, BetaO, xio0, xio1, xio2, xio3, xio4, xio As Double
            If iLatO = 0 Then
                Mo = 0
            ElseIf iLatO = (Math.PI / 2) Then
                Mo = BI * (Math.PI / 2)
            ElseIf iLatO = -(Math.PI / 2) Then
                Mo = BI * -(Math.PI / 2)
            Else
                Qo = ASinH(Math.Tan(iLatO)) - (ecc * ATanH(ecc * Math.Sin(iLatO)))
                BetaO = Math.Atan(Math.Sinh(Qo))
                xio0 = Math.Asin(Math.Sin(BetaO))
                xio1 = h1 * Math.Sin(2 * xio0)
                xio2 = h2 * Math.Sin(4 * xio0)
                xio3 = h3 * Math.Sin(6 * xio0)
                xio4 = h4 * Math.Sin(8 * xio0)
                xio = xio0 + xio1 + xio2 + xio3 + xio4
                Mo = BI * xio
            End If

            'From EN to LL
            Dim Q, Qi, Beta, xi0, xi1, xi2, xi3, xi4, xi, Eta0, Eta1, Eta2, Eta3, Eta4, Eta As Double
            Eta = (point.X - iEastO - Z * 10 ^ 6) / (BI * iK0)
            xi = ((point.Y - iNorthO) + iK0 * Mo) / (BI * iK0)
            xi1 = h1i * Math.Sin(2 * xi) * Math.Cosh(2 * Eta)
            xi2 = h2i * Math.Sin(4 * xi) * Math.Cosh(4 * Eta)
            xi3 = h3i * Math.Sin(6 * xi) * Math.Cosh(6 * Eta)
            xi4 = h4i * Math.Sin(8 * xi) * Math.Cosh(8 * Eta)
            xi0 = xi - (xi1 + xi2 + xi3 + xi4)
            Eta1 = h1i * Math.Cos(2 * xi) * Math.Sinh(2 * Eta)
            Eta2 = h2i * Math.Cos(4 * xi) * Math.Sinh(4 * Eta)
            Eta3 = h3i * Math.Cos(6 * xi) * Math.Sinh(6 * Eta)
            Eta4 = h4i * Math.Cos(8 * xi) * Math.Sinh(8 * Eta)
            Eta0 = Eta - (Eta1 + Eta2 + Eta3 + Eta4)
            Beta = Math.Asin(Math.Sin(xi0) / Math.Cosh(Eta0))
            Q = ASinH(Math.Tan(Beta))
            Qi = Q + (ecc * ATanH(ecc * Math.Tanh(Q)))
            For i = 0 To 100
                Qi = Q + (ecc * ATanH(ecc * Math.Tanh(Qi)))
            Next
            TmpResult.X = LonO + Math.Asin(Math.Tanh(Eta0) / Math.Cos(Beta))
            TmpResult.Y = Math.Atan(SinH(Qi))
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
        xproj.Add(New XElement(x.InitialLongitude, iLonI))
        xproj.Add(New XElement(x.ZoneWidth, iZoneW))
        xproj.Add(New XElement(x.OriginLatitude, iLatO))
        xproj.Add(New XElement(x.OriginScale, iK0))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue)

        tmpList.Add(New ParamNameValue("Initial longitude", InitialLongitude, ParamType.LatLong, False))
        tmpList.Add(New ParamNameValue("Zone width", ZoneWidth, ParamType.Angle, False))
        tmpList.Add(New ParamNameValue("Latitude of Natural Origin", LatitudeOfNaturalOrigin, ParamType.LatLong, True))
        tmpList.Add(New ParamNameValue("Scale factor at Natural Origin", ScaleFactorAtNaturalOrigin, ParamType.ScaleFactor, False))
        tmpList.Add(New ParamNameValue("False Easting at Natural Origin", FalseEasting, ParamType.EastNorth, False))
        tmpList.Add(New ParamNameValue("False Northing at Natural Origin", FalseNorthing, ParamType.EastNorth, True))

        Return tmpList

    End Function

    Public Overrides Function ToP190Header() As List(Of String)
        Dim ret As New List(Of String)
        ret.Add("H2600 Projection Not Supported".PadRight(80))
        Return ret
    End Function

End Class