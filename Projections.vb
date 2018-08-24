Public MustInherit Class Projections
    Dim iFname, iSname As String
    Dim iEll As Ellipsoid

    ''' <summary>
    ''' The reference ellipsoid for the projection
    ''' </summary>
    ''' <returns>Return an Ellipsoid object.</returns>
    Property BaseEllipsoid As Ellipsoid
        Get
            Return iEll
        End Get
        Set(value As Ellipsoid)
            iEll = value
        End Set
    End Property

    ''' <summary>
    ''' Name of the projection operation
    ''' </summary>
    ''' <returns></returns>
    Property FullName As String
        Get
            Return iFname
        End Get
        Set(value As String)
            iFname = value
        End Set
    End Property

    ''' <summary>
    ''' Short version of the name (max 12 characters) for some particular use case.
    ''' </summary>
    ''' <returns></returns>
    Property ShortName As String
        Get
            Return iSname
        End Get
        Set(value As String)
            If value.Length > 12 Then
                iSname = value.Substring(0, 12)
            Else
                iSname = value
            End If
        End Set
    End Property

    'Public Overridable ReadOnly Property Type As Projections.Method
    '    Get
    '        Return Method.PseudoPlateCarree
    '    End Get
    'End Property

    ''' <summary>
    ''' Coordinate projection operations available
    ''' </summary>
    Public Enum Method
        LambertConicConformal2SP
        LambertConicConformal1SP
        LambertConicConformalWest
        LambertConicConformalBelgium
        LambertConicNearConformal
        Krovak
        KrovakNorth
        KrovakModified
        KrovakModifiedNorth
        MercatorVariantA
        MercatorVariantB
        MercatorVariantC
        MercatorSpherical
        MercatorPseudo
        CassiniSoldner
        CassiniSoldnerHyperbolic
        TransverseMercator
        TransverseMercatorUniversal
        TransverseMercatorZoned
        TransverseMercatorSouth
        ObliqueMercatorHotineA
        ObliqueMercatorHotineB
        ObliqueMercatorLaborde
        StereographicOblique
        StereographicPolarA
        StereographicPolarB
        StereographicPolarC

        'NewZealandGrid
        'TunisiaMining
        AmericanPolyconic

        LambertAzimutalEqualArea
        LambertAzimutalEqualAreaPolar

        'LambertAzimutalEqualAreaSpherical
        'LambertCylindricalEqualArea
        'LambertCylindricalEqualAreaSpherical
        AlbersEqualArea

        EquidistantCylindrical
        EquidistantCylindricalSpherical
        PseudoPlateCarree
        'Bonne
        'BonneSouth
        'AzimutalEquidistantModified
        'AzimutalEquidistantGuam

    End Enum

    Public Shared Function FromXml(value As XElement) As Projections
        Dim x As New XmlTags

        If value.Name = x.Projection Then
            Dim xfname, xsname, xtype As String
            Dim xLatO, xLat1, xLat2, xLonO, xEastO, xNorthO, xScaleO, xAzimut, xAngle, xLonI, xWidth, xFuse, xHemi As Double
            Dim xEll As XElement
            Dim eEll As Ellipsoid

            'Common values
            Try
                xfname = value.Element(x.FullName).Value
            Catch ex As Exception
                xfname = x.DefaultString
            End Try
            Try
                xsname = value.Element(x.ShortName).Value
            Catch ex As Exception
                xsname = x.DefaultString.ToUpper
            End Try
            Try
                xtype = value.Element(x.Type).Value
            Catch ex As Exception
                xtype = Method.MercatorPseudo.ToString
            End Try

            Try
                xLatO = Double.Parse(value.Element(x.OriginLatitude).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xLatO = 0.0
            End Try
            Try
                xLonO = Double.Parse(value.Element(x.OriginLongitude).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xLonO = 0.0
            End Try
            Try
                xEastO = Double.Parse(value.Element(x.OriginEasting).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xEastO = 0.0
            End Try
            Try
                xNorthO = Double.Parse(value.Element(x.OriginNorthing).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xNorthO = 0.0
            End Try
            Try
                xScaleO = Double.Parse(value.Element(x.OriginScale).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xScaleO = 0.0
            End Try
            Try
                xLat1 = Double.Parse(value.Element(x.FirstLatitude).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xLat1 = 0.0
            End Try
            Try
                xLat2 = Double.Parse(value.Element(x.SecondLatitude).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xLat2 = 0.0
            End Try
            Try
                xAzimut = Double.Parse(value.Element(x.AzimutLine).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xAzimut = 0.0
            End Try
            Try
                xAngle = Double.Parse(value.Element(x.AngleGrid).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xAngle = 0.0
            End Try
            Try
                xLonI = Double.Parse(value.Element(x.InitialLongitude).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xLonI = 0.0
            End Try
            Try
                xWidth = Double.Parse(value.Element(x.ZoneWidth).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xWidth = 0.0
            End Try
            Try
                xFuse = Double.Parse(value.Element(x.UTMFuse).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xFuse = 0.0
            End Try
            Try
                xHemi = Double.Parse(value.Element(x.UTMHemisphere).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xHemi = 0.0
            End Try
            Try
                xEll = value.Element(x.Ellipsoid)
                eEll = Ellipsoid.FromXml(xEll)
            Catch ex As Exception
                eEll = New Ellipsoid
            End Try

            Select Case xtype
                Case Method.AlbersEqualArea.ToString
                    Return New AlbersConicEqualArea(eEll, xfname, xsname, xLatO, xLonO, xLat1, xLat2, xEastO, xNorthO)

                Case Method.AmericanPolyconic.ToString
                    Return New AmericanPolyconic(eEll, xfname, xsname, xLatO, xLonO, xEastO, xNorthO)

                'Case Method.AzimutalEquidistantGuam.ToString
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)

                'Case Method.AzimutalEquidistantModified.ToString
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)

                'Case Method.Bonne.ToString
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)

                'Case Method.BonneSouth.ToString
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)

                Case Method.CassiniSoldner.ToString
                    Return New CassiniSoldner(eEll, xfname, xsname, xLatO, xLonO, xEastO, xNorthO)

                Case Method.CassiniSoldnerHyperbolic.ToString
                    Return New CassiniSoldnerHyperbolic(eEll, xfname, xsname, xLatO, xLonO, xEastO, xNorthO)

                Case Method.EquidistantCylindrical.ToString
                    Return New EquidistantCylindrical(eEll, xfname, xsname, xLat1, xLonO, xEastO, xNorthO)

                Case Method.EquidistantCylindricalSpherical.ToString
                    Return New EquidistantCylindricalSpherical(eEll, xfname, xsname, xLat1, xLonO, xEastO, xNorthO)

                Case Method.Krovak.ToString
                    Return New Krovak(eEll, xfname, xsname, xLatO, xLonO, xAzimut, xLat1, xScaleO, xEastO, xNorthO)

                Case Method.KrovakModified.ToString
                    Return New KrovakModified(eEll, xfname, xsname, xLatO, xLonO, xAzimut, xLat1, xScaleO, xEastO, xNorthO)

                Case Method.KrovakNorth.ToString
                    Return New KrovakNorth(eEll, xfname, xsname, xLatO, xLonO, xAzimut, xLat1, xScaleO, xEastO, xNorthO)

                Case Method.KrovakModifiedNorth.ToString
                    Return New KrovakModifiedNorth(eEll, xfname, xsname, xLatO, xLonO, xAzimut, xLat1, xScaleO, xEastO, xNorthO)

                Case Method.LambertAzimutalEqualArea.ToString
                    Return New LambertAzimutalEqualAreaOblique(eEll, xfname, xsname, xLatO, xLonO, xEastO, xNorthO)

                Case Method.LambertAzimutalEqualAreaPolar.ToString
                    Return New LambertAzimutalEqualAreaPolar(eEll, xfname, xsname, xLatO, xLonO, xEastO, xNorthO)

                'Case Method.LambertAzimutalEqualAreaSpherical.ToString
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjLambertAzimutalEqualAreaOblique(eEll, xfname, xsname, xLatO, xLonO, xEastO, xNorthO)

                Case Method.LambertConicConformal1SP.ToString
                    Return New LambertConic1SP(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.LambertConicConformal2SP.ToString
                    Return New LambertConic2SP(eEll, xfname, xsname, xLatO, xLonO, xLat1, xLat2, xEastO, xNorthO)

                Case Method.LambertConicConformalBelgium.ToString
                    Return New LambertConicBelgium(eEll, xfname, xsname, xLatO, xLonO, xLat1, xLat2, xEastO, xNorthO)

                Case Method.LambertConicConformalWest.ToString
                    Return New LambertConicWest(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.LambertConicNearConformal.ToString
                    Return New LambertConicNear(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.MercatorPseudo.ToString
                    Return New MercatorPseudo(eEll, xfname, xsname, xLonO, xEastO, xNorthO)

                Case Method.MercatorSpherical.ToString
                    Return New MercatorSpherical(eEll, xfname, xsname, xLonO, xEastO, xNorthO)

                Case Method.MercatorVariantA.ToString
                    Return New MercatorA(eEll, xfname, xsname, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.MercatorVariantB.ToString
                    Return New MercatorB(eEll, xfname, xsname, xLonO, xLat1, xEastO, xNorthO)

                Case Method.MercatorVariantC.ToString
                    Return New MercatorC(eEll, xfname, xsname, xLonO, xLat1, xLatO, xEastO, xNorthO)

                Case Method.ObliqueMercatorHotineA.ToString
                    Return New ObliqueMercatorHotineA(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xAzimut, xAngle, xEastO, xNorthO)

                Case Method.ObliqueMercatorHotineB.ToString
                    Return New ObliqueMercatorHotineB(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xAzimut, xAngle, xEastO, xNorthO)

                Case Method.ObliqueMercatorLaborde.ToString
                    Return New ObliqueMercatorLaborde(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xAzimut, xEastO, xNorthO)

                Case Method.PseudoPlateCarree.ToString
                    Return New PseudoPlateCarree(eEll, xfname, xsname)

                Case Method.StereographicOblique.ToString
                    Return New StereographicOblique(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.StereographicPolarA.ToString
                    Return New StereographicPolarA(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.StereographicPolarB.ToString
                    Return New StereographicPolarB(eEll, xfname, xsname, xLat1, xLonO, xEastO, xNorthO)

                Case Method.StereographicPolarC.ToString
                    Return New StereographicPolarC(eEll, xfname, xsname, xLat1, xLonO, xEastO, xNorthO)

                Case Method.TransverseMercator.ToString
                    Return New TransverseMercator(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.TransverseMercatorSouth.ToString
                    Return New TransverseMercatorSouth(eEll, xfname, xsname, xLatO, xLonO, xScaleO, xEastO, xNorthO)

                Case Method.TransverseMercatorZoned.ToString
                    Return New TransverseMercatorZoned(eEll, xfname, xsname, xLonI, xWidth, xLatO, xScaleO, xEastO, xNorthO)

                Case Method.TransverseMercatorUniversal.ToString
                    Dim tmpFuse As Byte = CByte(xFuse)
                    Dim tmpHemi As Boolean = (xHemi > 0)
                    Return New TransverseMercatorUniversal(eEll, xfname, xsname, tmpFuse, tmpHemi)

                Case Else
                    Return New MercatorPseudo(New Ellipsoid, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)

            End Select
        Else
            Return New MercatorPseudo(New Ellipsoid, x.DefaultString, x.DefaultString.ToUpper, 0.0, 0.0, 0.0)

        End If

    End Function

    ''' <summary>
    ''' Convert projected East/North coordinates to Lat/Long coordinates.
    ''' </summary>
    ''' <param name="point">Projected point coordinates to be converted.</param>
    ''' <returns>Point coordinates on the ellipsoid.</returns>
    Public MustOverride Function ToGeographic(point As jakStd20_MathExt.Point3D) As jakStd20_MathExt.Point3D

    ''' <summary>
    ''' Convert Lat/Long coordinates to projected East/North coordinates.
    ''' </summary>
    ''' <param name="point">Geographic point coordinates to convert.</param>
    ''' <returns>Projected point coordinates.</returns>
    Public MustOverride Function FromGeographic(point As jakStd20_MathExt.Point3D) As jakStd20_MathExt.Point3D

    ''' <summary>
    ''' Projection method used.
    ''' </summary>
    ''' <returns></returns>
    Public MustOverride ReadOnly Property Type As Method

    Public MustOverride Function ToXml() As XElement

    Public MustOverride Function ToP190Header() As List(Of String)

    Public MustOverride Function GetParams() As List(Of ParamNameValue)

    Public Shared Function GetParamCount(projMethod As Method) As Integer
        Select Case projMethod
            Case Method.TransverseMercatorUniversal
                Return 2

            Case Method.MercatorPseudo, Method.MercatorSpherical
                Return 3

            Case Method.MercatorVariantA, Method.MercatorVariantB, Method.CassiniSoldner, Method.CassiniSoldnerHyperbolic
                Return 4
            Case Method.StereographicPolarB, Method.StereographicPolarC, Method.AmericanPolyconic
                Return 4
            Case Method.LambertAzimutalEqualArea, Method.LambertAzimutalEqualAreaPolar
                Return 4
            Case Method.EquidistantCylindrical, Method.EquidistantCylindricalSpherical
                Return 4

            Case Method.LambertConicConformal1SP, Method.LambertConicConformalWest, Method.LambertConicNearConformal
                Return 5
            Case Method.MercatorVariantC, Method.TransverseMercator, Method.TransverseMercatorSouth
                Return 5
            Case Method.StereographicOblique, Method.StereographicPolarA
                Return 5

            Case Method.LambertConicConformal2SP, Method.LambertConicConformalBelgium, Method.TransverseMercatorZoned
                Return 6
            Case Method.ObliqueMercatorLaborde, Method.AlbersEqualArea
                Return 6

            Case Method.Krovak, Method.KrovakModified, Method.KrovakModifiedNorth, Method.KrovakNorth
                Return 7
            Case Method.ObliqueMercatorHotineA, Method.ObliqueMercatorHotineB
                Return 7

            Case Else
                Return 0

        End Select
    End Function

    Public Shared Function FromParams(projType As Method, projEllipsoid As Ellipsoid, projFullName As String, projShortName As String, params As List(Of ParamNameValue)) As Projections
        Dim tmpParams As New List(Of Double)
        For p = 0 To params.Count - 1
            tmpParams.Add(params(p).Value)
        Next
        Return FromParams(projType, projEllipsoid, projFullName, projShortName, tmpParams)
    End Function

    Public Shared Function FromParams(projType As Method, projEllipsoid As Ellipsoid, projFullName As String, projShortName As String, params As List(Of Double)) As Projections

        'Add 7 zeroed value to Params to avoid exceptions
        params.Add(0.0)
        params.Add(0.0)
        params.Add(0.0)
        params.Add(0.0)
        params.Add(0.0)
        params.Add(0.0)
        params.Add(0.0)

        Select Case projType
            Case Method.LambertConicConformal2SP
                Return New LambertConic2SP(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5))

            Case Method.LambertConicConformal1SP
                Return New LambertConic1SP(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.LambertConicConformalWest
                Return New LambertConicWest(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.LambertConicConformalBelgium
                Return New LambertConicBelgium(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5))

            Case Method.LambertConicNearConformal
                Return New LambertConicNear(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.Krovak
                Return New Krovak(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5), params(6))

            Case Method.KrovakNorth
                Return New KrovakNorth(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5), params(6))

            Case Method.KrovakModified
                Return New KrovakModified(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5), params(6))

            Case Method.KrovakModifiedNorth
                Return New KrovakModifiedNorth(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5), params(6))

            Case Method.MercatorVariantA
                Return New MercatorA(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.MercatorVariantB
                Return New MercatorB(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.MercatorVariantC
                Return New MercatorC(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.MercatorSpherical
                Return New MercatorSpherical(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2))

            Case Method.MercatorPseudo
                Return New MercatorPseudo(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2))

            Case Method.CassiniSoldner
                Return New CassiniSoldner(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.CassiniSoldnerHyperbolic
                Return New CassiniSoldnerHyperbolic(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.TransverseMercator
                Return New TransverseMercator(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.TransverseMercatorZoned
                Return New TransverseMercatorZoned(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5))

            Case Method.TransverseMercatorSouth
                Return New TransverseMercatorSouth(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.TransverseMercatorUniversal
                Dim tmpFuse As Byte = CByte(params(0))
                Dim tmpHemi As Boolean = (params(1) > 0)
                Return New TransverseMercatorUniversal(projEllipsoid, projFullName, projShortName, tmpFuse, tmpHemi)

            Case Method.ObliqueMercatorHotineA
                Return New ObliqueMercatorHotineA(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5), params(6))

            Case Method.ObliqueMercatorHotineB
                Return New ObliqueMercatorHotineB(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5), params(6))

            Case Method.ObliqueMercatorLaborde
                Return New ObliqueMercatorLaborde(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5))

            Case Method.StereographicOblique
                Return New StereographicOblique(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.StereographicPolarA
                Return New StereographicPolarA(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4))

            Case Method.StereographicPolarB
                Return New StereographicPolarB(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.StereographicPolarC
                Return New StereographicPolarC(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

                'Case Method.NewZealandGrid
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

                'Case Method.TunisiaMiningGrid
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

            Case Method.AmericanPolyconic
                Return New AmericanPolyconic(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.LambertAzimutalEqualArea
                Return New LambertAzimutalEqualAreaOblique(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.LambertAzimutalEqualAreaPolar
                Return New LambertAzimutalEqualAreaPolar(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

                'Case Method.LambertAzimutalEqualAreaSpherical
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

                'Case Method.LambertCylindricalEqualArea
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

                'Case Method.LambertCylindricalEqualAreaSpherical
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

            Case Method.AlbersEqualArea
                Return New AlbersConicEqualArea(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3), params(4), params(5))

            Case Method.EquidistantCylindrical
                Return New EquidistantCylindrical(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.EquidistantCylindricalSpherical
                Return New EquidistantCylindricalSpherical(projEllipsoid, projFullName, projShortName, params(0), params(1), params(2), params(3))

            Case Method.PseudoPlateCarree
                Return New PseudoPlateCarree(projEllipsoid, projFullName, projShortName)

                'Case Method.Bonne
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

                'Case Method.BonneSouth
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

                'Case Method.AzimutalEquidistantModified
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

                'Case Method.AzimutalEquidistantGuam
                '    'TODO NOT IMPLEMENTED
                '    Return New ProjMercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

            Case Else
                Return New MercatorPseudo(New Ellipsoid, "Default", "DEFAULT", 0.0, 0.0, 0.0)

        End Select

    End Function

    Public Shared Function GetAvailableProjections() As List(Of String)
        Dim tmpRes As New List(Of String)

        tmpRes.Add("Lambert Conic Conformal 2SP") 'LambertConicConformal2SP
        tmpRes.Add("Lambert Conic Conformal 1SP") 'LambertConicConformal1SP
        tmpRes.Add("Lambert Conic Conformal West") 'LambertConicConformalWest
        tmpRes.Add("Lambert Conic Conformal Belgium") 'LambertConicConformalBelgium
        tmpRes.Add("Lambert Conic Near Conformal") 'LambertConicNearConformal
        tmpRes.Add("Krovak") 'Krovak
        tmpRes.Add("Krovak North") 'KrovakNorth
        tmpRes.Add("Krovak Modified") 'KrovakModified
        tmpRes.Add("Krovak Modified North") 'KrovakModifiedNorth
        tmpRes.Add("Mercator Variant A") 'MercatorVariantA
        tmpRes.Add("Mercator Variant B") 'MercatorVariantB
        tmpRes.Add("Mercator Variant C") 'MercatorVariantC
        tmpRes.Add("Mercator Spherical") 'MercatorSpherical
        tmpRes.Add("Mercator Pseudo") 'MercatorPseudo
        tmpRes.Add("Cassini Soldner") 'CassiniSoldner
        tmpRes.Add("Cassini Soldner Hyperbolic") 'CassiniSoldnerHyperbolic
        tmpRes.Add("Transverse Mercator") 'TransverseMercator
        tmpRes.Add("Universal Transverse Mercator") 'TransverseMercatorUniversal
        tmpRes.Add("Transverse Mercator Zoned") 'TransverseMercatorZoned
        tmpRes.Add("Transverse Mercator South") 'TransverseMercatorSouth
        tmpRes.Add("Oblique Mercator Hotine A") 'ObliqueMercatorHotineA
        tmpRes.Add("Oblique Mercator Hotine B") 'ObliqueMercatorHotineB
        tmpRes.Add("Oblique Mercator Laborde") 'ObliqueMercatorLaborde
        tmpRes.Add("Stereographic Oblique") 'StereographicOblique
        tmpRes.Add("Stereographic Polar A") 'StereographicPolarA
        tmpRes.Add("Stereographic Polar B") 'StereographicPolarB
        tmpRes.Add("Stereographic Polar C") 'StereographicPolarC
        'tmpRes.Add("New Zealand Grid") 'NewZealandGrid
        'tmpRes.Add("Tunisia Mining Grid") 'TunisiaMining
        tmpRes.Add("American Polyconic") 'AmericanPolyconic
        tmpRes.Add("Lambert Azimutal Equal Area") 'LambertAzimutalEqualArea
        tmpRes.Add("Lambert Azimutal Equal Area Polar") 'LambertAzimutalEqualAreaPolar
        'tmpRes.Add("Lambert Azimutal Equal Area Spherical") 'LambertAzimutalEqualAreaSpherical
        'tmpRes.Add("Lambert Cylindrical Equal Area") 'LambertCylindricalEqualArea
        'tmpRes.Add("Lambert Cylindrical Equal Area Spherical") 'LambertCylindricalEqualAreaSpherical
        tmpRes.Add("Albers Equal Area") 'AlbersEqualArea
        tmpRes.Add("Equidistant Cylindrical") 'EquidistantCylindrical
        tmpRes.Add("Equidistant Cylindrical Spherical") 'EquidistantCylindricalSpherical
        tmpRes.Add("Pseudo Plate Carree") 'PseudoPlateCarree
        'tmpRes.Add("Bonne") 'Bonne
        'tmpRes.Add("Bonne South") 'BonneSouth
        'tmpRes.Add("Azimutal Equidistant Modified") 'AzimutalEquidistantModified
        'tmpRes.Add("Azimutal Equidistant Guam") 'AzimutalEquidistantGuam

        Return tmpRes

    End Function

End Class