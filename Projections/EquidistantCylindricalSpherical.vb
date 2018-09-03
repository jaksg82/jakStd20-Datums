Imports MathExt
Imports StringFormat

Public Class EquidistantCylindricalSpherical
    Inherits Projections

    Private iLat1, iLonO, iEastO, iNorthO As Double

    Overrides ReadOnly Property Type As Method
        Get
            Return Method.EquidistantCylindricalSpherical
        End Get
    End Property

    ReadOnly Property LatitudeOfFirstStandardParallel As Double
        Get
            Return iLat1
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

    Public Sub New(refEllipsoid As Ellipsoid, longName As String, compactName As String, latOfFirstStandardParallel As Double,
                   lonOfOrigin As Double, falseEast As Double, falseNorth As Double)
        Me.FullName = longName
        Me.ShortName = compactName
        iLat1 = latOfFirstStandardParallel
        iLonO = lonOfOrigin
        iEastO = falseEast
        iNorthO = falseNorth
        Me.BaseEllipsoid = refEllipsoid

    End Sub

    Public Overrides Function FromGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'From LL to EN
            Dim R As Double
            R = BaseEllipsoid.GetRadiuosOfConformalSphere(iLat1)

            TmpResult.X = iEastO + R * (point.X - iLonO) * Math.Cos(iLat1)
            TmpResult.Y = iNorthO + R * point.Y
            TmpResult.Z = point.Z
            Return TmpResult
        Catch ex As Exception
            Return New Point3D

        End Try

    End Function

    Public Overrides Function ToGeographic(point As Point3D) As Point3D
        Dim TmpResult As New Point3D

        Try
            'From EN to LL
            Dim R As Double
            R = BaseEllipsoid.GetRadiuosOfConformalSphere(iLat1)

            TmpResult.X = iLonO + ((point.X - iEastO) / (R * Math.Cos(iLat1)))
            TmpResult.Y = (point.Y - iNorthO) / R
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
        xproj.Add(New XElement(x.FirstLatitude, iLat1))
        xproj.Add(New XElement(x.OriginLongitude, iLonO))
        xproj.Add(New XElement(x.OriginEasting, iEastO))
        xproj.Add(New XElement(x.OriginNorthing, iNorthO))
        xproj.Add(BaseEllipsoid.ToXml)
        Return xproj

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Latitude of first standard parallel", LatitudeOfFirstStandardParallel, ParamType.LatLong, True),
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

        'Latitude of standard parallel(s), (d.m.s. N/S)
        Dim H2100 As String = ps.H2100
        H2100 = H2100 & FormatDMS(LatitudeOfFirstStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2100 = H2100.PadRight(80, " "c)
        ret.Add(H2100)

        'Longitude of central meridian, (d.m.s. E/W)
        Dim H2200 As String = ps.H2200
        H2200 = H2200 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20)
        H2200 = H2200.PadRight(80, " "c)
        ret.Add(H2200)

        'Grid origin (Latitude, Longitude), (d.m.s. N/E)
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
        H2402 = H2402 & FormatDMS(LatitudeOfFirstStandardParallel, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, True).PadLeft(20, " "c)
        H2402 = H2402 & FormatDMS(LongitudeOfOrigin, DmsFormat.UkooaDMS, DmsSign.Suffix, 3, False).PadLeft(20, " "c)
        H2402 = H2402.PadRight(80, " "c)
        ret.Add(H2402)

        Return ret
    End Function

End Class