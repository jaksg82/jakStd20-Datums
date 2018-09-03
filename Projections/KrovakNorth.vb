Imports MathExt
Imports StringFormat

Public Class KrovakNorth
    Inherits Projections

    Private iLatO, iK0, iLonO, iEastO, iNorthO, iAzim, iLat1 As Double

    Public Overrides ReadOnly Property Type As Method
        Get
            Return Method.KrovakNorth
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
        'Return the results
        Dim TmpResult As New Point3D With {
            .Y = -(Xp + iNorthO), 'Southing to Nothing
            .X = -(Yp + iEastO), 'Westing to Easting
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
        'From EN to LL
        Dim U, V, T, D, Theta, r, Xp, Yp As Double
        Xp = -point.Y - iNorthO 'Southing
        Yp = -point.X - iEastO 'Westing
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