Imports MathExt
Imports StringFormat

Public Class KrovakModifiedNorth
    Inherits Projections

    Private iLatO, iK0, iLonO, iEastO, iNorthO, iAzim, iLat1 As Double
    Private Xo As Double = 1089000.0
    Private Yo As Double = 654000.0
    Private C1 As Double = 0.02946529277
    Private C2 As Double = 0.02515965696
    Private C3 As Double = 0.0000001193845912
    Private C4 As Double = -0.0000004668270147
    Private C5 As Double = 0.000000000009233980362
    Private C6 As Double = 0.000000000001523735715
    Private C7 As Double = 1.696780024E-18
    Private C8 As Double = 4.408314235E-18
    Private C9 As Double = -8.331083518E-24
    Private C10 As Double = -3.689471323E-24

    Public Overrides ReadOnly Property Type As Method
        Get
            Return Method.KrovakModifiedNorth
        End Get
    End Property

    ReadOnly Property LatitudeOfProjectionCenter As Double
        Get
            Return iLatO
        End Get
    End Property

    ReadOnly Property LongitudeOfOrigin As Double
        Get
            Return iLonO
        End Get
    End Property

    ReadOnly Property AzimutOfInitialLine As Double
        Get
            Return iAzim
        End Get
    End Property

    ReadOnly Property LatitudeOfPseudoStandardParallel As Double
        Get
            Return iLat1
        End Get
    End Property

    ReadOnly Property ScaleFactorOnPseudoStandardParallel As Double
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

    ReadOnly Property EvaluationPoint As Point3D
        Get
            Dim tmp As New Point3D With {
                .X = Xo,
                .Y = Yo,
                .Z = 0.0
            }
            Return tmp
        End Get
    End Property

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfProjectionCenter As Double, lonOfOrigin As Double,
                                            azimutOfLine As Double, latOfPseudoStandardParallel As Double, scaleOnPseudoStandardParallel As Double,
                                            falseEast As Double, falseNorth As Double)
        Me.FullName = longName
        Me.ShortName = compactName
        iLatO = latOfProjectionCenter
        iLonO = lonOfOrigin
        iAzim = azimutOfLine
        iLat1 = latOfPseudoStandardParallel
        iK0 = scaleOnPseudoStandardParallel
        iEastO = falseEast
        iNorthO = falseNorth
        Me.BaseEllipsoid = refEllipsoid

    End Sub

    Public Sub SetAdditionalParameters(paramXo As Double, paramYo As Double, paramC1 As Double, paramC2 As Double, paramC3 As Double, paramC4 As Double,
                                       paramC5 As Double, paramC6 As Double, paramC7 As Double, paramC8 As Double, paramC9 As Double, paramC10 As Double)
        Xo = paramXo
        Yo = paramYo
        C1 = paramC1
        C2 = paramC2
        C3 = paramC3
        C4 = paramC4
        C5 = paramC5
        C6 = paramC6
        C7 = paramC7
        C8 = paramC8
        C9 = paramC9
        C10 = paramC10
    End Sub

    Public Function GetCostant(id As Integer) As Double
        Dim resValue As Double = Double.NaN
        If id = 1 Then resValue = C1
        If id = 2 Then resValue = C2
        If id = 3 Then resValue = C3
        If id = 4 Then resValue = C4
        If id = 5 Then resValue = C5
        If id = 6 Then resValue = C6
        If id = 7 Then resValue = C7
        If id = 8 Then resValue = C8
        If id = 9 Then resValue = C9
        If id = 10 Then resValue = C10
        Return resValue
    End Function

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        'Ellipsoid costants
        Dim axis As Double = BaseEllipsoid.SemiMayorAxis
        Dim ecc As Double = BaseEllipsoid.Eccentricity

        'Projection costants
        Dim A, B, GammaOr, tOr, n, rOr As Double
        A = axis * (1 - ecc ^ 2) ^ 0.5 / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)
        B = (1 + (ecc ^ 2 * Math.Cos(iLatO) ^ 4 / (1 - ecc ^ 2))) ^ 0.5
        GammaOr = Math.Asin(Math.Sin(iLatO) / B)
        tOr = Math.Tan(Math.PI / 4 + GammaOr / 2) * ((1 + ecc * Math.Sin(iLatO)) / (1 - ecc * Math.Sin(iLatO))) ^ (ecc * B / 2) / (Math.Tan(Math.PI / 4 + iLatO / 2)) ^ B
        n = Math.Sin(iLat1)
        rOr = iK0 * A / Math.Tan(iLat1)
        'From LL to EN
        Dim U, V, T, D, Theta, r, Xp, Yp As Double
        U = 2 * (Math.Atan(tOr * Math.Tan(point.Y / 2 + Math.PI / 4) ^ B / ((1 + ecc * Math.Sin(point.Y)) / (1 - ecc * Math.Sin(point.Y))) ^ (ecc * B / 2)) - Math.PI / 4)
        V = B * (iLonO - point.X)
        T = Math.Asin(Math.Cos(iAzim) * Math.Sin(U) + Math.Sin(iAzim) * Math.Cos(U) * Math.Cos(V))
        D = Math.Asin(Math.Cos(U) * Math.Sin(V) / Math.Cos(T))
        Theta = n * D
        r = rOr * Math.Tan(Math.PI / 4 + iLat1 / 2) ^ n / Math.Tan(T / 2 + Math.PI / 4) ^ n
        Xp = r * Math.Cos(Theta)
        Yp = r * Math.Sin(Theta)
        'Modified part
        Dim Xr, Yr, dX, dY As Double
        Xr = Xp - Xo
        Yr = Yp - Yo
        dX = C1 + C3 * Xr - C4 * Yr - 2 * C6 * Xr * Yr + C5 * (Xr ^ 2 - Yr ^ 2) + C7 * Xr * (Xr ^ 2 - 3 * Yr ^ 2) 'Splitted for clarity
        dX = dX - C8 * Yr * (3 * Xr ^ 2 - Yr ^ 2) + 4 * C9 * Xr * Yr * (Xr ^ 2 - Yr ^ 2) + C10 * (Xr ^ 4 + Yr ^ 4 - 6 * Xr ^ 2 * Yr ^ 2)
        dY = C2 + C3 * Yr + C4 * Xr + 2 * C5 * Xr * Yr + C6 * (Xr ^ 2 - Yr ^ 2) + C8 * Xr * (Xr ^ 2 - 3 * Yr ^ 2) 'Splitted for clarity
        dY = dY + C7 * Yr * (3 * Xr ^ 2 - Yr) - 4 * C10 * Xr * Yr * (Xr ^ 2 - Yr ^ 2) + C9 * (Xr ^ 4 + Yr ^ 4 - 6 * Xr ^ 2 * Yr ^ 2)

        'Return the results
        Dim TmpResult As New Point3D With {
            .Y = -(Xp - dX + iNorthO), 'Southing
            .X = -(Yp - dY + iEastO), 'Westing
            .Z = point.Z
        }
        Return TmpResult

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        'Ellipsoid costants
        Dim axis As Double = BaseEllipsoid.SemiMayorAxis
        Dim ecc As Double = BaseEllipsoid.Eccentricity

        'Projection costants
        Dim A, B, GammaOr, tOr, n, rOr As Double
        A = axis * (1 - ecc ^ 2) ^ 0.5 / (1 - ecc ^ 2 * Math.Sin(iLatO) ^ 2)
        B = (1 + (ecc ^ 2 * Math.Cos(iLatO) ^ 4 / (1 - ecc ^ 2))) ^ 0.5
        GammaOr = Math.Asin(Math.Sin(iLatO) / B)
        tOr = Math.Tan(Math.PI / 4 + GammaOr / 2) * ((1 + ecc * Math.Sin(iLatO)) / (1 - ecc * Math.Sin(iLatO))) ^ (ecc * B / 2) / (Math.Tan(Math.PI / 4 + iLatO / 2)) ^ B
        n = Math.Sin(iLat1)
        rOr = iK0 * A / Math.Tan(iLat1)
        'Modified part
        Dim Xr, Yr, dX, dY As Double
        Xr = (-point.Y - iNorthO) - Xo
        Yr = (-point.X - iEastO) - Yo
        dX = C1 + C3 * Xr - C4 * Yr - 2 * C6 * Xr * Yr + C5 * (Xr ^ 2 - Yr ^ 2) + C7 * Xr * (Xr ^ 2 - 3 * Yr ^ 2) 'Splitted for clarity
        dX = dX - C8 * Yr * (3 * Xr ^ 2 - Yr ^ 2) + 4 * C9 * Xr * Yr * (Xr ^ 2 - Yr ^ 2) + C10 * (Xr ^ 4 + Yr ^ 4 - 6 * Xr ^ 2 * Yr ^ 2)
        dY = C2 + C3 * Yr + C4 * Xr + 2 * C5 * Xr * Yr + C6 * (Xr ^ 2 - Yr ^ 2) + C8 * Xr * (Xr ^ 2 - 3 * Yr ^ 2) 'Splitted for clarity
        dY = dY + C7 * Yr * (3 * Xr ^ 2 - Yr) - 4 * C10 * Xr * Yr * (Xr ^ 2 - Yr ^ 2) + C9 * (Xr ^ 4 + Yr ^ 4 - 6 * Xr ^ 2 * Yr ^ 2)

        'From EN to LL
        Dim U, V, T, D, Theta, r, Xp, Yp As Double
        Xp = (-point.Y - iNorthO) + dX  'Southing
        Yp = (-point.X - iEastO) + dY  'Westing
        r = (Xp ^ 2 + Yp ^ 2) ^ 0.5
        Theta = Math.Atan(Yp / Xp)
        D = Theta / Math.Sin(iLat1)
        T = 2 * (Math.Atan((rOr / r) ^ (1 / n) * Math.Tan(Math.PI / 4 + iLat1 / 2)) - Math.PI / 4)
        U = Math.Asin(Math.Cos(iAzim) * Math.Sin(T) - Math.Sin(iAzim) * Math.Cos(T) * Math.Cos(D))
        V = Math.Asin(Math.Cos(T) * Math.Sin(D) / Math.Cos(U))
        Dim tempLat As Double = U
        For i = 0 To 20
            tempLat = 2 * (Math.Atan(tOr ^ (-1 / B) * Math.Tan(U / 2 + Math.PI / 4) ^ (1 / B) * ((1 + ecc * Math.Sin(tempLat)) / (1 - ecc * Math.Sin(tempLat))) ^ (ecc / 2)) - Math.PI / 4)
        Next
        'Return the results
        Dim TmpResult As New Point3D With {
            .X = iLonO - V / B,
            .Y = tempLat,
            .Z = point.Z
        }
        Return TmpResult

    End Function

    Public Overrides Function ToXml() As XElement
        Dim x As New XmlTags
        Dim xproj As New XElement(x.Projection)
        xproj.Add(New XElement(x.FullName, FullName))
        xproj.Add(New XElement(x.ShortName, ShortName))
        xproj.Add(New XElement(x.Type, Type))
        xproj.Add(New XElement(x.OriginLatitude, iLatO))
        xproj.Add(New XElement(x.OriginLongitude, iLonO))
        xproj.Add(New XElement(x.AzimutLine, iAzim))
        xproj.Add(New XElement(x.FirstLatitude, iLat1))
        xproj.Add(New XElement(x.ScaleDifference, iK0))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Latitude of projection center", LatitudeOfProjectionCenter, ParamType.LatLong, True),
            New ParamNameValue("Longitude of Origin", LongitudeOfOrigin, ParamType.LatLong, False),
            New ParamNameValue("Azimut of initial line", AzimutOfInitialLine, ParamType.Angle, False),
            New ParamNameValue("Latitude of pseudo standard parallel", LatitudeOfPseudoStandardParallel, ParamType.LatLong, True),
            New ParamNameValue("Scale factor on pseudo standard parallel", ScaleFactorOnPseudoStandardParallel, ParamType.ScaleFactor, False),
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
        H2301 = H2301 & FormatDMS(LatitudeOfProjectionCenter, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
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
        H2401 = H2401 & FormatNumber(ScaleFactorOnPseudoStandardParallel, 10).PadLeft(20, " "c)
        H2401 = H2401.PadRight(80, " "c)
        ret.Add(H2401)

        'Latitude/Longitude at which scale factor is defined
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfPseudoStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        'Circular bearing of initial line of projection (d.m.s) or (grads)
        Dim H2507 As String = ps.H2507
        H2507 = H2507 & FormatNumber(AzimutOfInitialLine, 7).PadLeft(12, " "c)

        Return ret
    End Function

End Class