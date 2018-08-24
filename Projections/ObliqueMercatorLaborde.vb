Imports jakStd20_MathExt
Imports jakStd20_StringFormat

Public Class ObliqueMercatorLaborde
    Inherits Projections

    Dim iLatO, iAzim, iLonO, iEastO, iNorthO, iK0 As Double
    Dim PM As Double = DegRad(2.337229167)

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.ObliqueMercatorLaborde
        End Get
    End Property

    ReadOnly Property LatitudeOfProjectionCenter As Double
        Get
            Return iLatO
        End Get
    End Property

    ReadOnly Property LongitudeOfProjectionCenter As Double
        Get
            Return iLonO
        End Get
    End Property

    ReadOnly Property AzimutOfInitialLine As Double
        Get
            Return iAzim
        End Get
    End Property

    ReadOnly Property ScaleFactorOnInitialLine As Double
        Get
            Return iK0
        End Get
    End Property

    ReadOnly Property EastingAtProjectionCenter As Double
        Get
            Return iEastO
        End Get
    End Property

    ReadOnly Property NorthingAtProjectionCenter As Double
        Get
            Return iNorthO
        End Get
    End Property

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfProjectionCenter As Double, lonOfProjectionCenter As Double,
                                            scaleOnLine As Double, azimutOfLine As Double,
                                            falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLatO = latOfProjectionCenter
        iLonO = lonOfProjectionCenter
        iK0 = scaleOnLine
        iAzim = azimutOfLine
        iEastO = falseEast
        iNorthO = falseNorth
        BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim axis, ecc As Double
        Dim TmpResult As New Point3D

        Try
            'Reduce longitude to the Paris Meridian
            Dim PointLon As Double = point.X - PM

            'Ellipsoid costants
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity
            'Projection costants
            Dim B, LatS, R, C As Double
            B = (1 + (ecc ^ 2 * Math.Cos(iLatO) ^ 4) / (1 - ecc ^ 2)) ^ 0.5
            LatS = Math.Asin(Math.Sin(iLatO) / B)
            R = axis * iK0 * ((1 - ecc ^ 2) ^ 0.5 / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2))
            C = Math.Log(Math.Tan(Math.PI / 4 + LatS / 2)) - B * Math.Log(Math.Tan(Math.PI / 4 + iLatO / 2) * ((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))) ^ (ecc / 2))

            'From LL to EN
            Dim L, q, P, U, V, W, d, Li, Pi As Double
            L = B * (PointLon - iLonO)
            q = C + B * Math.Log(Math.Tan(Math.PI / 4 + point.Y / 2) * ((1 - ecc * Math.Sin(point.Y)) / (1 + ecc * Math.Sin(point.Y))) ^ (ecc / 2))
            P = 2 * Math.Atan(Math.E ^ q) - Math.PI / 2
            U = Math.Cos(P) * Math.Cos(L) * Math.Cos(LatS) + Math.Sin(P) * Math.Sin(LatS)
            V = Math.Cos(P) * Math.Cos(L) * Math.Sin(LatS) - Math.Sin(P) * Math.Cos(LatS)
            W = Math.Cos(P) * Math.Sin(L)
            d = (U ^ 2 + V ^ 2) ^ 0.5
            If d <> 0 Then
                Li = 2 * Math.Atan(V / (U + d))
                Pi = Math.Atan(W / d)
            Else
                Li = 0
                Pi = Math.Sign(W) * Math.PI / 2
            End If
            Dim H, G, HG As Numerics.Complex
            H = New Numerics.Complex(-Li, Math.Log(Math.Tan(Math.PI / 4 + Pi / 2)))
            G = New Numerics.Complex(1 - Math.Cos(2 * iAzim), Math.Sin(2 * iAzim)) / 12
            HG = H + G * Numerics.Complex.Pow(H, 3)

            'Return the results
            TmpResult.X = iEastO + R * HG.Imaginary
            TmpResult.Y = iNorthO + R * HG.Real
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
            'Projection costants
            Dim B, LatS, R, C, i As Double
            B = (1 + (ecc ^ 2 * Math.Cos(iLatO) ^ 4) / (1 - ecc ^ 2)) ^ 0.5
            LatS = Math.Asin(Math.Sin(iLatO) / B)
            R = axis * iK0 * ((1 - ecc ^ 2) ^ 0.5 / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2))
            C = Math.Log(Math.Tan(Math.PI / 4 + LatS / 2)) - B * Math.Log(Math.Tan(Math.PI / 4 + iLatO / 2) * ((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))) ^ (ecc / 2))
            i = Math.Sqrt(-1)

            'From EN to LL
            Dim H0, H1, H1a, H1b, G, HG As Numerics.Complex
            Dim L, P, U, V, W, d, ChkLoop, q, Lat0, Lat1 As Double
            G = New Numerics.Complex(1 - Math.Cos(2 * iAzim), Math.Sin(2 * iAzim)) / 12
            H0 = New Numerics.Complex((point.Y - iNorthO) / R, (point.X - iEastO) / R)
            'H1 = H0 / (H0 + G * Numerics.Complex.Pow(H0, 3))
            H1 = Numerics.Complex.Divide(H0, Numerics.Complex.Add(H0, Numerics.Complex.Multiply(G, Numerics.Complex.Pow(H0, 3))))
            Do
                'H1 = (H0 + 2 * G * Numerics.Complex.Pow(H1, 3)) / (3 * G * Numerics.Complex.Pow(H1, 2))
                H1a = Numerics.Complex.Add(H0, Numerics.Complex.Multiply(2, Numerics.Complex.Multiply(G, Numerics.Complex.Pow(H1, 3))))
                H1b = Numerics.Complex.Multiply(3, Numerics.Complex.Multiply(G, Numerics.Complex.Pow(H1, 2)))
                H1 = Numerics.Complex.Divide(H1a, H1b)
                'HG = H0 - H1 - G * Numerics.Complex.Pow(H1, 3)
                HG = Numerics.Complex.Subtract(H0, Numerics.Complex.Subtract(H1, Numerics.Complex.Multiply(G, Numerics.Complex.Pow(H1, 3))))
                ChkLoop = Math.Abs(HG.Real)
                If ChkLoop < 0.00000000001 Then Exit Do
            Loop
            L = -1 * H1.Real
            P = 2 * Math.Atan(Math.E ^ H1.Imaginary) - Math.PI / 2
            U = Math.Cos(P) * Math.Cos(L) * Math.Cos(LatS) + Math.Cos(P) * Math.Sin(L) * Math.Sin(LatS)
            V = Math.Sin(P)
            W = Math.Cos(P) * Math.Cos(L) * Math.Sin(LatS) - Math.Cos(P) * Math.Sin(L) * Math.Cos(LatS)
            d = (U ^ 2 + V ^ 2) ^ 0.5
            If d <> 0 Then
                L = 2 * Math.Atan(V / (U + d))
                P = Math.Atan(W / d)
            Else
                L = 0
                P = Math.Sign(W) * Math.PI / 2
            End If
            q = (Math.Log(Math.Tan(Math.PI / 4 + P / 2)) - C) / B
            Lat0 = 2 * Math.Atan(Math.E ^ q) - Math.PI / 2
            Do
                Lat1 = 2 * Math.Atan(((1 + ecc * Math.Sin(Lat0)) / (1 - ecc * Math.Sin(Lat0))) ^ (ecc / 2) * Math.E ^ q) - Math.PI / 2
                If Math.Abs(Lat1 - Lat0) < 0.00000000001 Then Exit Do
                Lat0 = Lat1
            Loop

            'Return the results
            TmpResult.X = iLonO + (L / B) + PM
            TmpResult.Y = Lat1
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
        xproj.Add(New XElement(x.AzimutLine, iAzim))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue)

        tmpList.Add(New ParamNameValue("Latitude of projection center", LatitudeOfProjectionCenter, ParamType.LatLong, True))
        tmpList.Add(New ParamNameValue("Longitude of projection center", LongitudeOfProjectionCenter, ParamType.LatLong, False))
        tmpList.Add(New ParamNameValue("Scale factor on initial line", ScaleFactorOnInitialLine, ParamType.ScaleFactor, False))
        tmpList.Add(New ParamNameValue("Azimut of the initial line", AzimutOfInitialLine, ParamType.Angle, True))
        tmpList.Add(New ParamNameValue("Easting at projection center", EastingAtProjectionCenter, ParamType.EastNorth, False))
        tmpList.Add(New ParamNameValue("Northing at projection center", NorthingAtProjectionCenter, ParamType.EastNorth, True))

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
        H2301 = H2301 & FormatDMS(LatitudeOfProjectionCenter, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2301 = H2301 & FormatDMS(LongitudeOfProjectionCenter, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2301 = H2301.PadRight(80, " "c)
        ret.Add(H2301)

        'Grid co-ordinates at grid origin (E,N)
        Dim H2302 As String = ps.H2302
        H2302 = H2302 & FormatMetric(EastingAtProjectionCenter, MetricSign.Suffix, 2, False).PadLeft(20, " "c)
        H2302 = H2302 & FormatMetric(NorthingAtProjectionCenter, MetricSign.Suffix, 2, True).PadLeft(20, " "c)
        H2302 = H2302.PadRight(80, " "c)
        ret.Add(H2302)

        'Scale factor
        Dim H2401 As String = ps.H2401
        H2401 = H2401 & FormatNumber(ScaleFactorOnInitialLine, 10).PadLeft(20, " "c)
        H2401 = H2401.PadRight(80, " "c)
        ret.Add(H2401)

        'Latitude/Longitude at which scale factor is defined
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfProjectionCenter, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfProjectionCenter, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        'Circular bearing of initial line of projection (d.m.s) or (grads)
        Dim H2507 As String = ps.H2507
        H2507 = H2507 & FormatNumber(AzimutOfInitialLine, 7).PadLeft(12, " "c)

        Return ret
    End Function

End Class