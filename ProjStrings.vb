Public Class ProjStrings
    Public ReadOnly H1800 As String = ("H1800 PROJECTION NAME".PadRight(32))           'Projection Code and Name
    Public ReadOnly H1900 As String = ("H1900 PROJECTION ZONE").PadRight(32)           'Projection Zone including Hemisphere
    Public ReadOnly H2100 As String = ("H2100 STANDARD PARALLELS").PadRight(32)        'Latitude of standard parallel(s), (d.m.s. N/S)
    Public ReadOnly H2200 As String = ("H2200 CENTRAL MERIDIAN").PadRight(32)          'Longitude of central meridian, (d.m.s. E/W)
    Public ReadOnly H2301 As String = ("H2301 GRID ORIGIN").PadRight(32)               'Grid origin (Latitude, Longitude), (d.m.s. N/E)
    Public ReadOnly H2302 As String = ("H2302 GRID COORDINATE AT ORIGIN").PadRight(32) 'Grid co-ordinates at grid origin (E,N)
    Public ReadOnly H2401 As String = ("H2401 SCALE FACTOR").PadRight(32)              'Scale factor
    Public ReadOnly H2402 As String = ("H2402 SCALE FACTOR ORIGIN").PadRight(32)       'Latitude/Longitude at which scale factor is defined 2*(d.m.s)
    Public ReadOnly H2506 As String = ("H2506 INITIAL LINE POINTS").PadRight(32)       'Latitude/Longitude of two points defining initial line of projection 4*(d.m.s)
    Public ReadOnly H2507 As String = ("H2507 LINE CIRCULAR BEARING").PadRight(32)     'Circular bearing of initial line of projection (d.m.s) or (grads)
    Public ReadOnly H2508 As String = ("H2508 LINE QUADRANT BEARING").PadRight(32)     'Quadrant bearing of initial line of projection (N/S d.m.s E/W) or (N/S grads E/W)
    Public ReadOnly H2509 As String = ("H2509 ANGLE SKEW-RECTIFIED GRID").PadRight(32) 'Angle from skew to rectified grid (d.m.s.) or (grads)

    Public ReadOnly H2600ProjType As String = ("H2600 PROJECTION TYPE").PadRight(32)   'Projection Type for the not defined projections (Code 999)

    Public ReadOnly HemiNorth As String = "NORTH HEMISPHERE"
    Public ReadOnly HemiSouth As String = "SOUTH HEMISPHERE"

    Public Function TypeString(value As Projections.Method) As String
        If value = Projections.Method.LambertConicConformal2SP Then Return ("Lambert Conic Conformal 2SP")  'LambertConicConformal2SP
        If value = Projections.Method.LambertConicConformal1SP Then Return ("Lambert Conic Conformal 1SP") 'LambertConicConformal1SP
        If value = Projections.Method.LambertConicConformalWest Then Return ("Lambert Conic Conformal West") 'LambertConicConformalWest
        If value = Projections.Method.LambertConicConformalBelgium Then Return ("Lambert Conic Conformal Belgium") 'LambertConicConformalBelgium
        If value = Projections.Method.LambertConicNearConformal Then Return ("Lambert Conic Near Conformal") 'LambertConicNearConformal
        If value = Projections.Method.Krovak Then Return ("Krovak") 'Krovak
        If value = Projections.Method.KrovakNorth Then Return ("Krovak North") 'KrovakNorth
        If value = Projections.Method.KrovakModified Then Return ("Krovak Modified") 'KrovakModified
        If value = Projections.Method.KrovakModifiedNorth Then Return ("Krovak Modified North") 'KrovakModifiedNorth
        If value = Projections.Method.MercatorVariantA Then Return ("Mercator Variant A") 'MercatorVariantA
        If value = Projections.Method.MercatorVariantB Then Return ("Mercator Variant B") 'MercatorVariantB
        If value = Projections.Method.MercatorVariantC Then Return ("Mercator Variant C") 'MercatorVariantC
        If value = Projections.Method.MercatorSpherical Then Return ("Mercator Spherical") 'MercatorSpherical
        If value = Projections.Method.MercatorPseudo Then Return ("Mercator Pseudo") 'MercatorPseudo
        If value = Projections.Method.CassiniSoldner Then Return ("Cassini Soldner") 'CassiniSoldner
        If value = Projections.Method.CassiniSoldnerHyperbolic Then Return ("Cassini Soldner Hyperbolic") 'CassiniSoldnerHyperbolic
        If value = Projections.Method.TransverseMercator Then Return ("Transverse Mercator") 'TransverseMercator
        If value = Projections.Method.TransverseMercatorUniversal Then Return ("Universal Transverse Mercator") 'TransverseMercatorUniversal
        If value = Projections.Method.TransverseMercatorZoned Then Return ("Transverse Mercator Zoned") 'TransverseMercatorZoned
        If value = Projections.Method.TransverseMercatorSouth Then Return ("Transverse Mercator South") 'TransverseMercatorSouth
        If value = Projections.Method.ObliqueMercatorHotineA Then Return ("Oblique Mercator Hotine A") 'ObliqueMercatorHotineA
        If value = Projections.Method.ObliqueMercatorHotineB Then Return ("Oblique Mercator Hotine B") 'ObliqueMercatorHotineB
        If value = Projections.Method.ObliqueMercatorLaborde Then Return ("Oblique Mercator Laborde") 'ObliqueMercatorLaborde
        If value = Projections.Method.StereographicOblique Then Return ("Stereographic Oblique") 'StereographicOblique
        If value = Projections.Method.StereographicPolarA Then Return ("Stereographic Polar A") 'StereographicPolarA
        If value = Projections.Method.StereographicPolarB Then Return ("Stereographic Polar B") 'StereographicPolarB
        If value = Projections.Method.StereographicPolarC Then Return ("Stereographic Polar C") 'StereographicPolarC
        'If value = Projections.Method.NewZealandGrid Then Return ("New Zealand Grid") 'NewZealandGrid
        'If value = Projections.Method.TunisiaMining Then Return ("Tunisia Mining Grid") 'TunisiaMining
        If value = Projections.Method.AmericanPolyconic Then Return ("American Polyconic") 'AmericanPolyconic
        If value = Projections.Method.LambertAzimutalEqualArea Then Return ("Lambert Azimutal Equal Area") 'LambertAzimutalEqualArea
        If value = Projections.Method.LambertAzimutalEqualAreaPolar Then Return ("Lambert Azimutal Equal Area Polar") 'LambertAzimutalEqualAreaPolar
        'If value = Projections.Method.LambertAzimutalEqualAreaSpherical Then Return ("Lambert Azimutal Equal Area Spherical") 'LambertAzimutalEqualAreaSpherical
        'If value = Projections.Method.LambertCylindricalEqualArea Then Return ("Lambert Cylindrical Equal Area") 'LambertCylindricalEqualArea
        'If value = Projections.Method.LambertCylindricalEqualAreaSpherical Then Return ("Lambert Cylindrical Equal Area Spherical") 'LambertCylindricalEqualAreaSpherical
        If value = Projections.Method.AlbersEqualArea Then Return ("Albers Equal Area") 'AlbersEqualArea
        If value = Projections.Method.EquidistantCylindrical Then Return ("Equidistant Cylindrical") 'EquidistantCylindrical
        If value = Projections.Method.EquidistantCylindricalSpherical Then Return ("Equidistant Cylindrical Spherical") 'EquidistantCylindricalSpherical
        If value = Projections.Method.PseudoPlateCarree Then Return ("Pseudo Plate Carree") 'PseudoPlateCarree
        'If value = Projections.Method.Bonne Then Return ("Bonne") 'Bonne
        'If value = Projections.Method.BonneSouth Then Return ("Bonne South") 'BonneSouth
        'If value = Projections.Method.AzimutalEquidistantModified Then Return ("Azimutal Equidistant Modified") 'AzimutalEquidistantModified
        'If value = Projections.Method.AzimutalEquidistantGuam Then Return ("Azimutal Equidistant Guam") 'AzimutalEquidistantGuam

        Return "Not Available!"
    End Function
End Class
