Imports MathExt
Imports StringFormat

Public Class AmericanPolyconic
    Inherits Projections

    Private iLatO, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Projections.Method.AmericanPolyconic
        End Get
    End Property

    ReadOnly Property LatitudeOfNaturalOrigin As Double
        Get
            Return iLatO
        End Get
    End Property

    ReadOnly Property LongitudeOfNaturalOrigin As Double
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

    Public Sub New(RefEllipsoid As Ellipsoid, longName As String, compactName As String, latOfNaturalOrigin As Double, lonOfNaturalOrigin As Double,
                                             falseEast As Double, falseNorth As Double)
        Me.FullName = longName
        Me.ShortName = compactName
        iLatO = latOfNaturalOrigin
        iLonO = lonOfNaturalOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        Me.BaseEllipsoid = RefEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'Ellipsoid costants
            Dim axis, ecc As Double
            axis = BaseEllipsoid.SemiMayorAxis
            ecc = BaseEllipsoid.Eccentricity

            'From LL to EN
            Dim E1, E2, E3, E4, L, M, Mo, V As Double
            E1 = 1 - ecc ^ 2 / 4 - 3 * ecc ^ 4 / 64 - 5 * ecc ^ 6 / 256
            E2 = 3 * ecc ^ 2 / 8 + 3 * ecc ^ 4 / 32 + 45 * ecc ^ 6 / 1024
            E3 = 15 * ecc ^ 4 / 256 + 45 * ecc ^ 6 / 1024
            E4 = 35 * ecc ^ 6 / 3072
            L = (point.X - iEastO) * Math.Sin(point.Y)
            V = BaseEllipsoid.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            M = axis * (E1 * point.Y - E2 * Math.Sin(2 * point.Y) + E3 * Math.Sin(4 * point.Y) - E4 * Math.Sin(6 * point.Y))
            Mo = axis * (E1 * iLatO - E2 * Math.Sin(2 * iLatO) + E3 * Math.Sin(4 * iLatO) - E4 * Math.Sin(6 * iLatO))
            If point.Y = 0 Then
                TmpResult.X = iEastO + axis * (point.X - iLonO)
                TmpResult.Y = iNorthO - Mo
            Else
                TmpResult.X = iEastO + V * CTan(point.Y) * Math.Sin(L)
                TmpResult.Y = iNorthO + M - Mo + V * CTan(point.Y) * (1 - Math.Cos(L))
            End If

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
            Dim E1, E2, E3, E4, M, Mo, A, B, C, J, LatN, LatN1 As Double
            E1 = 1 - ecc ^ 2 / 4 - 3 * ecc ^ 4 / 64 - 5 * ecc ^ 6 / 256
            E2 = 3 * ecc ^ 2 / 8 + 3 * ecc ^ 4 / 32 + 45 * ecc ^ 6 / 1024
            E3 = 15 * ecc ^ 4 / 256 + 45 * ecc ^ 6 / 1024
            E4 = 35 * ecc ^ 6 / 3072
            Mo = axis * (E1 * iLatO - E2 * Math.Sin(2 * iLatO) + E3 * Math.Sin(4 * iLatO) - E4 * Math.Sin(6 * iLatO))
            If Mo = (point.Y - iNorthO) Then
                TmpResult.Y = 0.0
                TmpResult.X = iEastO + (point.X - iEastO) / axis
            Else
                A = (Mo + (point.Y - iNorthO)) / axis
                B = A ^ 2 + (((point.X - iEastO) ^ 2) / axis ^ 2)
                LatN = A
                Do
                    C = (1 - ecc ^ 2 * Math.Sin(LatN) ^ 2) ^ 0.5 * Math.Tan(LatN)
                    M = axis * (E1 * LatN - E2 * Math.Sin(2 * LatN) + E3 * Math.Sin(4 * LatN) - E4 * Math.Sin(6 * LatN))
                    J = M / axis
                    LatN1 = LatN - (A * (C * J + 1) - J - 0.5 * C * (J ^ 2 + B)) / (ecc ^ 2 * Math.Sin(2 * LatN) * (J ^ 2 + B - 2 * A * J) / (4 * C) + (A - J) * (C * M - (2 / Math.Sin(2 * LatN)) - M))
                    If Math.Abs(LatN - LatN1) < 0.0000000001 Then Exit Do
                    LatN = LatN1
                Loop
                TmpResult.Y = LatN
                TmpResult.X = iEastO + (Math.Asin((point.X - iEastO) * C / axis)) / Math.Sin(LatN)

            End If

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
            New ParamNameValue("Latitude of Natural Origin", LatitudeOfNaturalOrigin, ParamType.LatLong, True),
            New ParamNameValue("Longitude of Natural Origin", LongitudeOfNaturalOrigin, ParamType.LatLong, False),
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

        'Longitude of central meridian, (d.m.s E/W)
        Dim H2200 As String = ps.H2200
        H2200 = H2200 & FormatDMS(LongitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude), (d.m.s N/E)
        Dim H2301 As String = ps.H2301
        H2301 = H2301 & FormatDMS(LatitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2301 = H2301 & FormatDMS(LongitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
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

        'Latitude/Longitude at which scale factor is defined 2*(d.m.s)
        Dim H2402 As String = ps.H2402
        H2402 = H2402 & FormatDMS(LatitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfNaturalOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class