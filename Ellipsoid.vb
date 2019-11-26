Imports System.Math
Imports MathExt

Public Class Ellipsoid
    Private iSphere As Boolean
    Private iEpsg As Integer
    Private iAxis As Double
    Private iInvFlat As Double

    Public Property IsSphere As Boolean
        Get
            Return iSphere
        End Get
        Set(value As Boolean)
            iSphere = value
        End Set
    End Property

    Public Property FullName As String
    Public Property ShortName As String

    Public Property EpsgId As Integer
        Get
            Return iEpsg
        End Get
        Set(value As Integer)
            iEpsg = value
        End Set
    End Property

    Public Property SemiMayorAxis As Double
        Get
            Return iAxis
        End Get
        Set(value As Double)
            If IsFinite(value) Then iAxis = value
        End Set
    End Property

    Public Property InverseFlattening As Double
        Get
            Return iInvFlat
        End Get
        Set(value As Double)
            If IsFinite(value) Then iInvFlat = value
        End Set
    End Property

    Public Property ToWgs84 As Transformations

    Public ReadOnly Property SemiMinorAxis As Double
        Get
            Return If(IsSphere, SemiMayorAxis, (1 - Flattening) * SemiMayorAxis)
        End Get
    End Property

    Public ReadOnly Property Flattening As Double
        Get
            Return If(IsSphere, 1.0, 1 / InverseFlattening)
        End Get
    End Property

    Public ReadOnly Property Eccentricity As Double
        Get
            Return Sqrt((2 * Flattening) - Flattening ^ 2)
        End Get
    End Property

    Public ReadOnly Property SecondEccentricity As Double
        Get
            Return Sqrt(Eccentricity ^ 2 / (1 - Eccentricity ^ 2))
        End Get
    End Property

    ''' <summary>
    ''' Create a default ellipsoid (World Geodetic System 1984 EPSG:7030)
    ''' </summary>
    Public Sub New()
        iEpsg = 7030
        FullName = "World Geodetic System 1984"
        ShortName = "WGS_1984"
        iAxis = 6378137
        iInvFlat = 298.2572236
        iSphere = False
        ToWgs84 = New None

    End Sub

    ''' <summary>
    ''' Create a default ellipsoid (World Geodetic System 1984 EPSG:7030)
    ''' </summary>
    ''' <param name="datum">Datum Transformation Method</param>
    Public Sub New(datum As Transformations)
        iEpsg = 7030
        FullName = "World Geodetic System 1984"
        ShortName = "WGS_1984"
        iAxis = 6378137
        iInvFlat = 298.2572236
        iSphere = False
        ToWgs84 = datum

    End Sub

    ''' <summary>
    ''' Create an ellipsoid from the given paramenters
    ''' </summary>
    ''' <param name="id">Epsg Id</param>
    ''' <param name="fullName">Complete name of the Ellipsoid</param>
    ''' <param name="ShortName">Short name max 12 characters</param>
    ''' <param name="SemiMayorAxis">Semi mayor axis</param>
    ''' <param name="InverseFlattening">Inverse flattening. Input zero or NaN to define a sphere.</param>
    Public Sub New(id As Integer, fullName As String, shortName As String, semiMayorAxis As Double, inverseFlattening As Double)
        iEpsg = id
        Me.FullName = fullName
        If shortName Is Nothing Then
            Me.ShortName = "DEFAULT"
        Else
            If shortName.Length > 12 Then
                Me.ShortName = shortName.Substring(0, 12)
            Else
                Me.ShortName = shortName
            End If

        End If
        If IsFinite(semiMayorAxis) Then
            iAxis = semiMayorAxis
        Else
            iAxis = 6378137
        End If

        If IsFinite(inverseFlattening) Then
            If inverseFlattening = 0.0 Then
                Me.InverseFlattening = inverseFlattening
                IsSphere = True
            Else
                Me.InverseFlattening = inverseFlattening
                IsSphere = False
            End If
        Else
            Me.InverseFlattening = inverseFlattening
            IsSphere = True
        End If
        ToWgs84 = New None

    End Sub

    ''' <summary>
    ''' Create an ellipsoid from the given paramenters
    ''' </summary>
    ''' <param name="id">Epsg Id</param>
    ''' <param name="fullName">Complete name of the Ellipsoid</param>
    ''' <param name="ShortName">Short name max 12 characters</param>
    ''' <param name="SemiMayorAxis">Semi mayor axis</param>
    ''' <param name="InverseFlattening">Inverse flattening. Input zero or NaN to define a sphere.</param>
    ''' <param name="datum">Datum Transformation Method</param>
    Public Sub New(id As Integer, fullName As String, shortName As String, semiMayorAxis As Double, inverseFlattening As Double, datum As Transformations)
        iEpsg = id
        Me.FullName = fullName
        If shortName Is Nothing Then
            Me.ShortName = "DEFAULT"
        Else
            If shortName.Length > 12 Then
                Me.ShortName = shortName.Substring(0, 12)
            Else
                Me.ShortName = shortName
            End If

        End If
        If IsFinite(semiMayorAxis) Then
            iAxis = semiMayorAxis
        Else
            iAxis = 6378137
        End If

        If IsFinite(inverseFlattening) Then
            If inverseFlattening = 0.0 Then
                Me.InverseFlattening = inverseFlattening
                IsSphere = True
            Else
                Me.InverseFlattening = inverseFlattening
                IsSphere = False
            End If
        Else
            Me.InverseFlattening = inverseFlattening
            IsSphere = True
        End If
        ToWgs84 = datum

    End Sub

    ''' <summary>
    ''' Create an ellipsoid from the selected EPSG ellipsoid.
    ''' </summary>
    ''' <param name="id">Epsg Id</param>
    Public Sub New(id As Integer)
        Dim epsgell As New EpsgEllipsoids
        Dim tmpell As Ellipsoid = epsgell.GetFromEpsgID(id)

        iEpsg = tmpell.EpsgId
        Me.FullName = tmpell.FullName
        Me.ShortName = tmpell.ShortName
        iAxis = tmpell.SemiMayorAxis
        iInvFlat = tmpell.InverseFlattening
        If IsFinite(tmpell.InverseFlattening) Then
            If tmpell.InverseFlattening = 0.0 Then
                IsSphere = True
            Else
                IsSphere = False
            End If
        Else
            IsSphere = True
        End If
        ToWgs84 = New None

    End Sub

    ''' <summary>
    ''' Create an ellipsoid from the selected EPSG ellipsoid.
    ''' </summary>
    ''' <param name="id">Epsg Id</param>
    ''' <param name="datum">Datum Transformation Method</param>
    Public Sub New(id As Integer, datum As Transformations)
        Dim epsgell As New EpsgEllipsoids
        Dim tmpell As Ellipsoid = epsgell.GetFromEpsgID(id)

        iEpsg = tmpell.EpsgId
        Me.FullName = tmpell.FullName
        Me.ShortName = tmpell.ShortName
        iAxis = tmpell.SemiMayorAxis
        iInvFlat = tmpell.InverseFlattening
        If IsFinite(tmpell.InverseFlattening) Then
            If tmpell.InverseFlattening = 0.0 Then
                IsSphere = True
            Else
                IsSphere = False
            End If
        Else
            IsSphere = True
        End If
        ToWgs84 = datum

    End Sub

    ''' <summary>
    ''' Radious of curvature of the ellipsoid in the plane of the meridian at a certain latitude.
    ''' </summary>
    ''' <param name="latitude">Latitude of where to calculate the radious.</param>
    ''' <returns>Return the computed radious.</returns>
    Public Function GetRadiuosOfCurvatureInTheMeridian(latitude As Double) As Double
        Dim p As Double
        p = (SemiMayorAxis * (1 - Eccentricity ^ 2)) / ((1 - ((Eccentricity ^ 2) * Sin(latitude) ^ 2)) ^ (3 / 2))
        Return p
    End Function

    ''' <summary>
    ''' Radious of curvature of the ellipsoid perpendicular to the meridian at a certain latitude.
    ''' </summary>
    ''' <param name="latitude">Latitude of where to calculate the radious.</param>
    ''' <returns>Return the computed radious.</returns>
    Public Function GetRadiuosOfCurvatureInThePrimeVertical(latitude As Double) As Double
        Dim v As Double
        v = SemiMayorAxis / ((1 - ((Eccentricity ^ 2) * Sin(latitude) ^ 2)) ^ (1 / 2))
        Return v
    End Function

    ''' <summary>
    ''' Radious of sphere having same surface area as ellipsoid.
    ''' </summary>
    ''' <returns>Return the computed radious.</returns>
    ReadOnly Property GetRadiuosOfAuthalicSphere As Double
        Get
            Dim Ra As Double
            Ra = SemiMayorAxis * ((1 - ((1 - Eccentricity ^ 2) / (2 * Eccentricity)) * (Log((1 - Eccentricity) / (1 + Eccentricity)))) * 0.5) ^ 0.5
            Return Ra
        End Get
    End Property

    ''' <summary>
    ''' Radious of the conformal sphere.
    ''' </summary>
    ''' <param name="latitude"></param>
    ''' <returns>Return the computed radious.</returns>
    Public Function GetRadiuosOfConformalSphere(latitude As Double) As Double
        Dim v, p, Rc As Double
        v = SemiMayorAxis / ((1 - ((Eccentricity ^ 2) * Sin(latitude) ^ 2)) ^ (1 / 2))
        p = (SemiMayorAxis * (1 - Eccentricity ^ 2)) / ((1 - ((Eccentricity ^ 2) * Sin(latitude) ^ 2)) ^ (3 / 2))
        Rc = Sqrt(p * v)
        Return Rc
    End Function

    Public Function Distance(point1 As Point3D, ByVal point2 As Point3D, Optional ByRef alpha1 As Double = 0.0, Optional ByRef alpha2 As Double = 0.0) As Double
        Dim smax, smix, flatt, lat1, lat2, U1, U2, deltaLon, lon1, lon2, dist, sinSigma, cosSigma, Sigma, lambda As Double
        Dim sinAlpha, cos2Alpha, cos2sigma, Cc, lambdaPrev, Usquare, Aa, Bb, k1, DeltaSigma As Double
        Dim maxIter As Integer = 100000
        Dim Iter As Integer = 0

        If point1.X = point2.X And point1.Y = point2.Y Then
            Return 0
        Else
            smax = Me.SemiMayorAxis
            flatt = Me.Flattening
            smix = Me.SemiMinorAxis
            lat1 = point1.Y
            lat2 = point2.Y
            U1 = Atan((1 - flatt) * Tan(lat1))
            U2 = Atan((1 - flatt) * Tan(lat2))
            lon1 = point1.X
            lon2 = point2.X
            deltaLon = lon2 - lon1
            lambda = lon2 - lon1
            If lat1 = 0 And lat2 = 0 Then
                'Geodesic runs along equator
                dist = smax * lambda
                If lon1 > lon2 Then
                    alpha1 = DegRad(270)
                    alpha2 = alpha1
                Else
                    alpha1 = DegRad(90)
                    alpha2 = alpha1
                End If
            Else
                Do
                    sinSigma = Sqrt((Cos(U2) * Sin(lambda)) ^ 2 + (Cos(U1) * Sin(U2) - Sin(U1) * Cos(U2) * Cos(lambda)) ^ 2)
                    cosSigma = Sin(U1) * Sin(U2) + Cos(U1) * Cos(U2) * Cos(lambda)
                    Sigma = Atan2(sinSigma, cosSigma)
                    sinAlpha = (Cos(U1) * Cos(U2) * Sin(lambda)) / sinSigma
                    cos2Alpha = 1 - sinAlpha ^ 2
                    cos2sigma = cosSigma - ((2 * Sin(U1) * Sin(U2)) / cos2Alpha)
                    Cc = (flatt / 16) * cos2Alpha * (4 + flatt * (4 - 3 * cos2Alpha))
                    lambdaPrev = lambda
                    lambda = deltaLon + (1 - Cc) * flatt * sinAlpha * (Sigma + Cc * sinSigma * (cos2sigma + Cc * cosSigma * (-1 + 2 * cos2sigma ^ 2)))
                    Iter = Iter + 1
                Loop Until Abs(lambda - lambdaPrev) < 0.0000000000000001 Or Iter = maxIter
                Usquare = cos2Alpha * ((smax ^ 2 - smix ^ 2) / smix ^ 2)
                k1 = (Sqrt(1 + Usquare) - 1) / (Sqrt(1 + Usquare) + 1)
                Aa = (1 + 0.25 * k1 ^ 2) / (1 - k1)
                Bb = k1 * (1 - (3 / 8) * k1 ^ 2)
                DeltaSigma = Bb * sinSigma * (cos2sigma + 0.25 * Bb * (cosSigma * (-1 + 2 * cos2sigma ^ 2) - (1 / 6) * Bb * cos2sigma * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cosSigma ^ 2)))
                dist = smix * Aa * (Sigma - DeltaSigma)
                'Alpha1 = Atan((Cos(U2) * Sin(lambda)) / (Cos(U1) * Sin(U2) - Sin(U1) * Cos(U2) * Cos(lambda)))
                'Alpha2 = Atan((Cos(U1) * Sin(lambda)) / (-Sin(U1) * Cos(U2) + Cos(U1) * Sin(U2) * Cos(lambda)))
                Dim Alp1 As Double = Atan2(Cos(U2) * Sin(lambda), (Cos(U1) * Sin(U2) - Sin(U1) * Cos(U2) * Cos(lambda)))
                Dim Alp2 As Double = Atan2(Cos(U1) * Sin(lambda), (-Sin(U1) * Cos(U2) + Cos(U1) * Sin(U2) * Cos(lambda)))
                alpha1 = Alp1
                alpha2 = Alp2
                If Double.IsNaN(dist) Then dist = 0
            End If
            Return dist
        End If
    End Function

    Public Function ToXml(Optional withTransformationValues As Boolean = True) As XElement
        'New(ID As Integer, FullName As String, ShortName As String, SemiMayorAxis As Double, InverseFlattening As Double)
        Dim x As New XmlTags

        Dim ellx As New XElement(x.Ellipsoid)
        ellx.Add(New XElement(x.EpsgId, EpsgId))
        ellx.Add(New XElement(x.FullName, FullName))
        ellx.Add(New XElement(x.ShortName, ShortName))
        ellx.Add(New XElement(x.SemiMayorAxis, SemiMayorAxis))
        ellx.Add(New XElement(x.InverseFlattening, InverseFlattening))
        If ToWgs84 IsNot Nothing Then
            If withTransformationValues Then
                ellx.Add(New XElement(x.ToWgs84, ToWgs84.ToXml))
            End If
        End If

        Return ellx

    End Function

    Public Shared Function FromXml(value As XElement) As Ellipsoid
        'New(ID As Integer, FullName As String, ShortName As String, SemiMayorAxis As Double, InverseFlattening As Double)
        Dim x As New XmlTags
        Dim xID As Integer
        Dim xFname, xSname As String
        Dim xSma, xInv As Double
        Dim xWgs As XElement
        Dim xEll As New Ellipsoid

        Try
            xFname = value.Element(x.FullName).Value
        Catch ex As Exception
            xFname = x.DefaultString
        End Try
        Try
            xSname = value.Element(x.ShortName).Value
        Catch ex As Exception
            xSname = x.DefaultString.ToUpper
        End Try
        Try
            xID = Integer.Parse(value.Element(x.EpsgId).Value, Globalization.CultureInfo.InvariantCulture)
        Catch ex As Exception
            xID = 0
        End Try
        Try
            xSma = Double.Parse(value.Element(x.SemiMayorAxis).Value, Globalization.CultureInfo.InvariantCulture)
        Catch ex As Exception
            xSma = 0.0
        End Try
        Try
            xInv = Double.Parse(value.Element(x.InverseFlattening).Value, Globalization.CultureInfo.InvariantCulture)
        Catch ex As Exception
            xInv = 0.0
        End Try

        xEll.FullName = xFname
        xEll.ShortName = xSname
        xEll.EpsgId = xID
        xEll.SemiMayorAxis = xSma
        xEll.InverseFlattening = xInv

        Try
            xWgs = value.Element(x.ToWgs84)
            If xWgs IsNot Nothing Then
                xEll.ToWgs84 = Transformations.FromXml(xWgs)
            Else
                xEll.ToWgs84 = New Geocentric3p(xEll, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)
            End If
        Catch ex As Exception
            xEll.ToWgs84 = New Geocentric3p(xEll, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)
        End Try

        Return xEll

    End Function

End Class