Public Class EpsgEllipsoids
    Private iEpsg As New List(Of Ellipsoid)

    Public ReadOnly Property KnownEllipsoids As List(Of Ellipsoid)
        Get
            Return iEpsg
        End Get
    End Property

    Public Sub New()

        iEpsg.Add(New Ellipsoid(7001, "Airy 1830", "AIRY1830", 6377563.396, 299.3249646))
        iEpsg.Add(New Ellipsoid(7002, "Airy Modified 1849", "AIRYMOD1849", 6377340.189, 299.3249646))
        iEpsg.Add(New Ellipsoid(7003, "Australian National Spheroid", "AUSTRALIAN", 6378160, 298.25))
        iEpsg.Add(New Ellipsoid(7041, "Average Terrestrial System 1977", "AVG1977", 6378135, 298.257))
        iEpsg.Add(New Ellipsoid(7004, "Bessel 1841", "BESSEL1841", 6377397.155, 299.1528128))
        iEpsg.Add(New Ellipsoid(7005, "Bessel Modified", "BESSELMOD", 6377492.018, 299.1528128))
        iEpsg.Add(New Ellipsoid(1024, "CGCS2000", "CGCS2000", 6378137, 298.257222101))
        iEpsg.Add(New Ellipsoid(7008, "Clarke 1866", "CLARKE1866", 6378206.4, 294.978698213898))
        iEpsg.Add(New Ellipsoid(7052, "Clarke 1866 Authalic Sphere", "CLARKE66SPH", 6370997, 0))
        iEpsg.Add(New Ellipsoid(7013, "Clarke 1880 (Arc)", "CLARKE80ARC", 6378249.145, 293.4663077))
        iEpsg.Add(New Ellipsoid(7010, "Clarke 1880 (Benoit)", "CLARKE80BEN", 6378300.789, 293.466315538981))
        iEpsg.Add(New Ellipsoid(7011, "Clarke 1880 (IGN)", "CLARKE80IGN", 6378249.2, 293.466021293627))
        iEpsg.Add(New Ellipsoid(7012, "Clarke 1880 (RGS)", "CLARKE80RGS", 6378249.145, 293.465))
        iEpsg.Add(New Ellipsoid(7014, "Clarke 1880 (SGA 1922)", "CLARKE80SGA", 6378249.2, 293.46598))
        iEpsg.Add(New Ellipsoid(7051, "Danish 1876", "DANISH1876", 6377019.27, 300))
        iEpsg.Add(New Ellipsoid(7018, "Everest 1830 Modified", "EVEREST1830", 6377304.063, 300.8017))
        iEpsg.Add(New Ellipsoid(7015, "Everest 1830 (1937 Adjustment)", "EVEREST1937", 6377276.345, 300.8017))
        iEpsg.Add(New Ellipsoid(7044, "Everest 1830 (1962 Definition)", "EVEREST1962", 6377301.243, 300.8017255))
        iEpsg.Add(New Ellipsoid(7016, "Everest 1830 (1967 Definition)", "EVEREST1967", 6377298.556, 300.8017))
        iEpsg.Add(New Ellipsoid(7056, "Everest 1830 (RSO 1969)", "EVERESTRSO69", 6377295.664, 300.8017))
        iEpsg.Add(New Ellipsoid(7045, "Everest 1830 (1975 Definition)", "EVEREST1975", 6377299.151, 300.8017255))
        iEpsg.Add(New Ellipsoid(7031, "GEM 10C", "GEM10C", 6378137, 298.257223563))
        iEpsg.Add(New Ellipsoid(7036, "GRS 1967", "GRS1967", 6378160, 298.247167427))
        iEpsg.Add(New Ellipsoid(7050, "GRS 1967 Modified", "GRS1967MOD", 6378160, 298.25))
        iEpsg.Add(New Ellipsoid(7019, "GRS 1980", "GRS1980", 6378137, 298.257222101))
        iEpsg.Add(New Ellipsoid(7048, "GRS 1980 Authalic Sphere", "GRS1980SPH", 6371007, 0))
        iEpsg.Add(New Ellipsoid(7020, "Helmert 1906", "HELMERT1906", 6378200, 298.3))
        iEpsg.Add(New Ellipsoid(7053, "Hough 1960", "HOUGH1960", 6378270, 297))
        iEpsg.Add(New Ellipsoid(7058, "Hughes 1980", "HUGHES1980", 6378273, 298.279411123064))
        iEpsg.Add(New Ellipsoid(7049, "IAG 1975", "IAG1975", 6378140, 298.257))
        iEpsg.Add(New Ellipsoid(7021, "Indonesian National Spheroid", "INDONESIAN", 6378160, 298.247))
        iEpsg.Add(New Ellipsoid(7022, "International 1924", "INTL1924", 6378388, 297))
        iEpsg.Add(New Ellipsoid(7057, "International 1924 Authalic Sphere", "INTL1924SPH", 6371228, 0))
        iEpsg.Add(New Ellipsoid(7024, "Krassowsky 1940", "KRASS1940", 6378245, 298.3))
        iEpsg.Add(New Ellipsoid(7025, "NWL 9D", "NWL9D", 6378145, 298.25))
        iEpsg.Add(New Ellipsoid(7032, "OSU86F", "OSU86F", 6378136.2, 298.257223563))
        iEpsg.Add(New Ellipsoid(7033, "OSU91A", "OSU91A", 6378136.3, 298.257223563))
        iEpsg.Add(New Ellipsoid(7054, "PZ-90", "PZ-90", 6378136, 298.257839303))
        iEpsg.Add(New Ellipsoid(7027, "Plessis 1817", "PLESSIS1817", 6376523, 308.64))
        iEpsg.Add(New Ellipsoid(7028, "Struve 1860", "STRUVE1860", 6378298.3, 294.73))
        iEpsg.Add(New Ellipsoid(7043, "WGS 72", "WGS72", 6378135, 298.26))
        iEpsg.Add(New Ellipsoid(7030, "WGS 84", "WGS84", 6378137, 298.257223563))
        iEpsg.Add(New Ellipsoid(7029, "War Office", "WAROFFICE", 6378300, 296))

    End Sub

    Public Function GetFromEpsgID(epsgId As Integer) As Ellipsoid
        Dim IsFound As Boolean = False
        Dim FoundId As Integer
        If epsgId >= 0 Then
            For e = 0 To iEpsg.Count - 1
                If iEpsg(e).EpsgId = epsgId Then
                    IsFound = True
                    FoundId = e
                End If
            Next
        End If
        If IsFound Then
            Return iEpsg(FoundId)
        Else
            Return New Ellipsoid
        End If
    End Function

    Public Function GetAllNames() As List(Of String)
        Dim tmpList As New List(Of String)
        For e = 0 To iEpsg.Count - 1
            tmpList.Add(iEpsg(e).FullName)
        Next
        Return tmpList

    End Function

    Public Function GetAllEpsgIdAndNames() As List(Of String)
        Dim tmpList As New List(Of String)
        For e = 0 To iEpsg.Count - 1
            tmpList.Add("[" & iEpsg(e).EpsgId & "] " & iEpsg(e).FullName)
        Next
        Return tmpList

    End Function

End Class