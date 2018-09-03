Imports MathExt
Imports StringFormat

Public Class ObliqueMercatorHotineB
    Inherits Projections

    Private iLatO, iAzim, iAngle, iLonO, iEastO, iNorthO, iK0 As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.ObliqueMercatorHotineB
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

    ReadOnly Property AngleFromRectifiedToSkewedGrid As Double
        Get
            Return iAngle
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
                                            scaleOnLine As Double, azimutOfLine As Double, angleGrids As Double,
                                            falseEast As Double, falseNorth As Double)
        FullName = longName
        ShortName = compactName
        iLatO = latOfProjectionCenter
        iLonO = lonOfProjectionCenter
        iK0 = scaleOnLine
        iAzim = azimutOfLine
        iAngle = angleGrids
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
            'Projection costants
            Dim B, A, TOr, D, F, H, G, GammaO, LonO, vc, uc As Double
            B = (1 + (ecc ^ 2 * Math.Cos(iLatO) ^ 4 / (1 - ecc ^ 2))) ^ 0.5
            A = axis * B * iK0 * (1 - ecc ^ 2) ^ 0.5 / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)
            TOr = Math.Tan(Math.PI / 4 - iLatO / 2) / ((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))) ^ (ecc / 2)
            D = B * (1 - ecc ^ 2) ^ 0.5 / (Math.Cos(iLatO) * (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2) ^ 0.5)
            If D < 1 Then
                F = D + (1 - 1) ^ 0.5 * Math.Sign(iLatO)
            Else
                F = D + (D ^ 2 - 1) ^ 0.5 * Math.Sign(iLatO)
            End If
            H = F * TOr ^ B
            G = (F - 1 / F) / 2
            GammaO = Math.Asin(Math.Sin(iAzim) / D)
            LonO = iLonO - (Math.Asin(G * Math.Tan(GammaO))) / B
            vc = 0
            If iAzim = Math.PI / 2 Then
                uc = A * (iLonO - LonO)
            Else
                uc = (A / B) * Math.Atan((D ^ 2 - 1) ^ 0.5 / Math.Cos(iAzim)) * Math.Sign(iLatO)
            End If

            'From LL to EN
            Dim t1, Q, S, T2, V1, U1, v2, u2 As Double
            t1 = Math.Tan(Math.PI / 4 - point.Y / 2) / ((1 - ecc * Math.Sin(point.Y)) / (1 + ecc * Math.Sin(point.Y))) ^ (ecc / 2)
            Q = H / t1 ^ B
            S = (Q - 1 / Q) / 2
            T2 = (Q + 1 / Q) / 2
            V1 = Math.Sin(B * (point.X - LonO))
            U1 = (-V1 * Math.Cos(GammaO) + S * Math.Sin(GammaO)) / T2
            v2 = A * Math.Log((1 - U1) / (1 + U1)) / (2 * B)
            If iAzim = Math.PI / 2 Then
                If point.X = iLonO Then
                    u2 = 0
                Else
                    u2 = (A * Math.Atan((S * Math.Cos(GammaO) + V1 * Math.Sin(GammaO)) / Math.Cos(B * (point.X - LonO))) / B) - (Math.Abs(uc) * Math.Sign(iLatO) * Math.Sign(iLonO - point.X))
                End If
            Else
                u2 = (A * Math.Atan((S * Math.Cos(GammaO) + V1 * Math.Sin(GammaO)) / Math.Cos(B * (point.X - LonO))) / B) - (Math.Abs(uc) * Math.Sign(iLatO))
            End If

            'Return the results
            TmpResult.X = v2 * Math.Cos(iAngle) + u2 * Math.Sin(iAngle) + iEastO
            TmpResult.Y = u2 * Math.Cos(iAngle) - v2 * Math.Sin(iAngle) + iNorthO
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
            Dim B, A, TOr, D, F, H, G, GammaO, LonO, vc, uc As Double
            B = (1 + (ecc ^ 2 * Math.Cos(iLatO) ^ 4 / (1 - ecc ^ 2))) ^ 0.5
            A = axis * B * iK0 * (1 - ecc ^ 2) ^ 0.5 / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)
            TOr = Math.Tan(Math.PI / 4 - iLatO / 2) / ((1 - ecc * Math.Sin(iLatO)) / (1 + ecc * Math.Sin(iLatO))) ^ (ecc / 2)
            D = B * (1 - ecc ^ 2) ^ 0.5 / (Math.Cos(iLatO) * (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2) ^ 0.5)
            If D < 1 Then
                F = D + (1 - 1) ^ 0.5 * Math.Sign(iLatO)
            Else
                F = D + (D ^ 2 - 1) ^ 0.5 * Math.Sign(iLatO)
            End If
            H = F * TOr ^ B
            G = (F - 1 / F) / 2
            GammaO = Math.Asin(Math.Sin(iAzim) / D)
            LonO = iLonO - (Math.Asin(G * Math.Tan(GammaO))) / B
            vc = 0
            If iAzim = Math.PI / 2 Then
                uc = A * (iLonO - LonO)
            Else
                uc = (A / B) * Math.Atan((D ^ 2 - 1) ^ 0.5 / Math.Cos(iAzim)) * Math.Sign(iLatO)
            End If

            'From EN to LL
            Dim v1, u1, Qi, Si, Ti, Vi, Ui, t1i, X As Double
            v1 = (point.X - iEastO) * Math.Cos(iAngle) - (point.Y - iNorthO) * Math.Sin(iAngle)
            u1 = (point.Y - iNorthO) * Math.Cos(iAngle) + (point.X - iEastO) * Math.Sin(iAngle) + (Math.Abs(uc) * Math.Sign(iLatO))

            Qi = Math.E ^ -(B * v1 / A)
            Si = (Qi - 1 / Qi) / 2
            Ti = (Qi + 1 / Qi) / 2
            Vi = Math.Sin(B * u1 / A)
            Ui = (Vi * Math.Cos(GammaO) + Si * Math.Sin(GammaO)) / Ti
            t1i = (H / ((1 + Ui) / (1 - Ui)) ^ 0.5) ^ (1 / B)
            X = Math.PI / 2 - 2 * Math.Atan(t1i)
            'Ellipsoid costants for latitude formula
            Dim E1, E2, E3, E4 As Double
            E1 = ecc ^ 2 / 2 + 5 * ecc ^ 4 / 24 + ecc ^ 6 / 12 + 13 * ecc ^ 8 / 360
            E2 = 7 * ecc ^ 4 / 48 + 29 * ecc ^ 6 / 240 + 811 * ecc ^ 8 / 11520
            E3 = 7 * ecc ^ 6 / 120 + 81 * ecc ^ 8 / 1120
            E4 = 4279 * ecc ^ 8 / 161280

            'Return the results
            TmpResult.X = LonO - Math.Atan((Si * Math.Cos(GammaO) - Vi * Math.Sin(GammaO)) / Math.Cos(B * u1 / A)) / B
            TmpResult.Y = X + Math.Sin(2 * X) * E1 + Math.Sin(4 * X) * E2 + Math.Sin(6 * X) * E3 + Math.Sin(8 * X) * E4
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
        xproj.Add(New XElement(x.AngleGrid, iAngle))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Latitude of projection center", LatitudeOfProjectionCenter, ParamType.LatLong, True),
            New ParamNameValue("Longitude of projection center", LongitudeOfProjectionCenter, ParamType.LatLong, False),
            New ParamNameValue("Scale factor on initial line", ScaleFactorOnInitialLine, ParamType.ScaleFactor, False),
            New ParamNameValue("Azimut of the initial line", AzimutOfInitialLine, ParamType.Angle, True),
            New ParamNameValue("Angle from rectified to skewed grid", AngleFromRectifiedToSkewedGrid, ParamType.Angle, True),
            New ParamNameValue("Easting at projection center", EastingAtProjectionCenter, ParamType.EastNorth, False),
            New ParamNameValue("Northing at projection center", NorthingAtProjectionCenter, ParamType.EastNorth, True)
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

        'Angle from skew to rectified grid (d.m.s.) or (grads)
        Dim H2509 As String = ps.H2509
        H2509 = H2509 & FormatNumber(AngleFromRectifiedToSkewedGrid, 7).PadLeft(12, " "c)

        Return ret
    End Function

End Class