Partial Public Class Ellipsoid

    'The coefficients C4[l] in the Fourier expansion of I4
    Private Function C4coeff(order As Integer) As Double()
        Dim C4index As Integer

        If order <= 0 Or order > 8 Then
            C4index = 0
        Else
            C4index = CInt((order * (order + 1)) / 2)
        End If

        'Found the ellipsoid costants
        Dim f1, e2, ep2 As Double
        f1 = 1 - Me.Flattening
        e2 = Me.Flattening * (2 - Me.Flattening)
        ep2 = (e2 / f1 ^ 2)
        'n = Me.Flattening / (2 - Me.Flattening)

        Dim calcC4(C4index) As Double

        Select Case order
            Case 1
                calcC4(0) = 2 / (3)
                Return calcC4
            Case 2
                calcC4(0) = (10 - ep2) / 15
                calcC4(1) = -1 / (20)
                calcC4(2) = 1 / (180)
                Return calcC4
            Case 3
                calcC4(0) = (ep2 * (4 * ep2 - 7) + 70) / 105
                calcC4(1) = (4 * ep2 - 7) / 140
                calcC4(2) = 1 / (42)
                calcC4(3) = (7 - 4 * ep2) / 1260
                calcC4(4) = -1 / (252)
                calcC4(5) = 1 / (2100)
                Return calcC4
            Case 4
                calcC4(0) = (ep2 * ((12 - 8 * ep2) * ep2 - 21) + 210) / 315
                calcC4(1) = ((12 - 8 * ep2) * ep2 - 21) / 420
                calcC4(2) = (3 - 2 * ep2) / 126
                calcC4(3) = -1 / (72)
                calcC4(4) = (ep2 * (8 * ep2 - 12) + 21) / 3780
                calcC4(5) = (2 * ep2 - 3) / 756
                calcC4(6) = 1 / (360)
                calcC4(7) = (3 - 2 * ep2) / 6300
                calcC4(8) = -1 / (1800)
                calcC4(9) = 1 / (17640)
                Return calcC4
            Case 5
                calcC4(0) = (ep2 * (ep2 * (ep2 * (64 * ep2 - 88) + 132) - 231) + 2310) / 3465
                calcC4(1) = (ep2 * (ep2 * (64 * ep2 - 88) + 132) - 231) / 4620
                calcC4(2) = (ep2 * (16 * ep2 - 22) + 33) / 1386
                calcC4(3) = (8 * ep2 - 11) / 792
                calcC4(4) = 1 / (110)
                calcC4(5) = (ep2 * ((88 - 64 * ep2) * ep2 - 132) + 231) / 41580
                calcC4(6) = ((22 - 16 * ep2) * ep2 - 33) / 8316
                calcC4(7) = (11 - 8 * ep2) / 3960
                calcC4(8) = -1 / (495)
                calcC4(9) = (ep2 * (16 * ep2 - 22) + 33) / 69300
                calcC4(10) = (8 * ep2 - 11) / 19800
                calcC4(11) = 1 / (1925)
                calcC4(12) = (11 - 8 * ep2) / 194040
                calcC4(13) = -1 / (10780)
                calcC4(14) = 1 / (124740)
                Return calcC4
            Case 6
                calcC4(0) = (ep2 * (ep2 * (ep2 * ((832 - 640 * ep2) * ep2 - 1144) + 1716) - 3003) + 30030) / 45045
                calcC4(1) = (ep2 * (ep2 * ((832 - 640 * ep2) * ep2 - 1144) + 1716) - 3003) / 60060
                calcC4(2) = (ep2 * ((208 - 160 * ep2) * ep2 - 286) + 429) / 18018
                calcC4(3) = ((104 - 80 * ep2) * ep2 - 143) / 10296
                calcC4(4) = (13 - 10 * ep2) / 1430
                calcC4(5) = -1 / (156)
                calcC4(6) = (ep2 * (ep2 * (ep2 * (640 * ep2 - 832) + 1144) - 1716) + 3003) / 540540
                calcC4(7) = (ep2 * (ep2 * (160 * ep2 - 208) + 286) - 429) / 108108
                calcC4(8) = (ep2 * (80 * ep2 - 104) + 143) / 51480
                calcC4(9) = (10 * ep2 - 13) / 6435
                calcC4(10) = 5 / (3276)
                calcC4(11) = (ep2 * ((208 - 160 * ep2) * ep2 - 286) + 429) / 900900
                calcC4(12) = ((104 - 80 * ep2) * ep2 - 143) / 257400
                calcC4(13) = (13 - 10 * ep2) / 25025
                calcC4(14) = -1 / (2184)
                calcC4(15) = (ep2 * (80 * ep2 - 104) + 143) / 2522520
                calcC4(16) = (10 * ep2 - 13) / 140140
                calcC4(17) = 5 / (45864)
                calcC4(18) = (13 - 10 * ep2) / 1621620
                calcC4(19) = -1 / (58968)
                calcC4(20) = 1 / (792792)
                Return calcC4
            Case 7
                calcC4(0) = (ep2 * (ep2 * (ep2 * (ep2 * (ep2 * (512 * ep2 - 640) + 832) - 1144) + 1716) - 3003) + 30030) / 45045
                calcC4(1) = (ep2 * (ep2 * (ep2 * (ep2 * (512 * ep2 - 640) + 832) - 1144) + 1716) - 3003) / 60060
                calcC4(2) = (ep2 * (ep2 * (ep2 * (128 * ep2 - 160) + 208) - 286) + 429) / 18018
                calcC4(3) = (ep2 * (ep2 * (64 * ep2 - 80) + 104) - 143) / 10296
                calcC4(4) = (ep2 * (8 * ep2 - 10) + 13) / 1430
                calcC4(5) = (4 * ep2 - 5) / 780
                calcC4(6) = 1 / (210)
                calcC4(7) = (ep2 * (ep2 * (ep2 * ((640 - 512 * ep2) * ep2 - 832) + 1144) - 1716) + 3003) / 540540
                calcC4(8) = (ep2 * (ep2 * ((160 - 128 * ep2) * ep2 - 208) + 286) - 429) / 108108
                calcC4(9) = (ep2 * ((80 - 64 * ep2) * ep2 - 104) + 143) / 51480
                calcC4(10) = ((10 - 8 * ep2) * ep2 - 13) / 6435
                calcC4(11) = (5 - 4 * ep2) / 3276
                calcC4(12) = -1 / (840)
                calcC4(13) = (ep2 * (ep2 * (ep2 * (128 * ep2 - 160) + 208) - 286) + 429) / 900900
                calcC4(14) = (ep2 * (ep2 * (64 * ep2 - 80) + 104) - 143) / 257400
                calcC4(15) = (ep2 * (8 * ep2 - 10) + 13) / 25025
                calcC4(16) = (4 * ep2 - 5) / 10920
                calcC4(17) = 1 / (2520)
                calcC4(18) = (ep2 * ((80 - 64 * ep2) * ep2 - 104) + 143) / 2522520
                calcC4(19) = ((10 - 8 * ep2) * ep2 - 13) / 140140
                calcC4(20) = (5 - 4 * ep2) / 45864
                calcC4(21) = -1 / (8820)
                calcC4(22) = (ep2 * (8 * ep2 - 10) + 13) / 1621620
                calcC4(23) = (4 * ep2 - 5) / 294840
                calcC4(24) = 1 / (41580)
                calcC4(25) = (5 - 4 * ep2) / 3963960
                calcC4(26) = -1 / (304920)
                calcC4(27) = 1 / (4684680)
                Return calcC4
            Case 8
                calcC4(0) = (ep2 * (ep2 * (ep2 * (ep2 * (ep2 * ((8704 - 7168 * ep2) * ep2 - 10880) + 14144) - 19448) + 29172) - 51051) + 510510) / 765765
                calcC4(1) = (ep2 * (ep2 * (ep2 * (ep2 * ((8704 - 7168 * ep2) * ep2 - 10880) + 14144) - 19448) + 29172) - 51051) / 1021020
                calcC4(2) = (ep2 * (ep2 * (ep2 * ((2176 - 1792 * ep2) * ep2 - 2720) + 3536) - 4862) + 7293) / 306306
                calcC4(3) = (ep2 * (ep2 * ((1088 - 896 * ep2) * ep2 - 1360) + 1768) - 2431) / 175032
                calcC4(4) = (ep2 * ((136 - 112 * ep2) * ep2 - 170) + 221) / 24310
                calcC4(5) = ((68 - 56 * ep2) * ep2 - 85) / 13260
                calcC4(6) = (17 - 14 * ep2) / 3570
                calcC4(7) = -1 / (272)
                calcC4(8) = (ep2 * (ep2 * (ep2 * (ep2 * (ep2 * (7168 * ep2 - 8704) + 10880) - 14144) + 19448) - 29172) + 51051) / 9189180
                calcC4(9) = (ep2 * (ep2 * (ep2 * (ep2 * (1792 * ep2 - 2176) + 2720) - 3536) + 4862) - 7293) / 1837836
                calcC4(10) = (ep2 * (ep2 * (ep2 * (896 * ep2 - 1088) + 1360) - 1768) + 2431) / 875160
                calcC4(11) = (ep2 * (ep2 * (112 * ep2 - 136) + 170) - 221) / 109395
                calcC4(12) = (ep2 * (56 * ep2 - 68) + 85) / 55692
                calcC4(13) = (14 * ep2 - 17) / 14280
                calcC4(14) = 7 / (7344)
                calcC4(15) = (ep2 * (ep2 * (ep2 * ((2176 - 1792 * ep2) * ep2 - 2720) + 3536) - 4862) + 7293) / 15315300
                calcC4(16) = (ep2 * (ep2 * ((1088 - 896 * ep2) * ep2 - 1360) + 1768) - 2431) / 4375800
                calcC4(17) = (ep2 * ((136 - 112 * ep2) * ep2 - 170) + 221) / 425425
                calcC4(18) = ((68 - 56 * ep2) * ep2 - 85) / 185640
                calcC4(19) = (17 - 14 * ep2) / 42840
                calcC4(20) = -7 / (20400)
                calcC4(21) = (ep2 * (ep2 * (ep2 * (896 * ep2 - 1088) + 1360) - 1768) + 2431) / 42882840
                calcC4(22) = (ep2 * (ep2 * (112 * ep2 - 136) + 170) - 221) / 2382380
                calcC4(23) = (ep2 * (56 * ep2 - 68) + 85) / 779688
                calcC4(24) = (14 * ep2 - 17) / 149940
                calcC4(25) = 1 / (8976)
                calcC4(26) = (ep2 * ((136 - 112 * ep2) * ep2 - 170) + 221) / 27567540
                calcC4(27) = ((68 - 56 * ep2) * ep2 - 85) / 5012280
                calcC4(28) = (17 - 14 * ep2) / 706860
                calcC4(29) = -7 / (242352)
                calcC4(30) = (ep2 * (56 * ep2 - 68) + 85) / 67387320
                calcC4(31) = (14 * ep2 - 17) / 5183640
                calcC4(32) = 7 / (1283568)
                calcC4(33) = (17 - 14 * ep2) / 79639560
                calcC4(34) = -1 / (1516944)
                calcC4(35) = 1 / (26254800)
                Return calcC4
            Case Else
                Return calcC4
        End Select
    End Function

    'The coefficients C3[l] in the Fourier expansion of B3
    Private Function C3coeff(order As Integer) As Double()
        'Found the ellipsoid costants
        Dim f1, e2, n As Double
        f1 = 1 - Me.Flattening
        e2 = Me.Flattening * (2 - Me.Flattening)
        'ep2 = (e2 / f1 ^ 2)
        n = Me.Flattening / (2 - Me.Flattening)

        Dim C3index As Integer

        If order <= 0 Or order > 8 Then
            C3index = 0
        Else
            C3index = CInt((order * (order - 1)) / 2)
        End If

        Dim calcC3(C3index) As Double

        Select Case order
            Case 2
                calcC3(0) = 1 / (4)
                Return calcC3
            Case 3
                calcC3(0) = (1 - n) / 4
                calcC3(1) = 1 / (8)
                calcC3(2) = 1 / (16)
                Return calcC3
            Case 4
                calcC3(0) = (1 - n) / 4
                calcC3(1) = 1 / (8)
                calcC3(2) = 3 / (64)
                calcC3(3) = (2 - 3 * n) / 32
                calcC3(4) = 3 / (64)
                calcC3(5) = 5 / (192)
                Return calcC3
            Case 5
                calcC3(0) = (1 - n) / 4
                calcC3(1) = (1 - n * n) / 8
                calcC3(2) = (3 * n + 3) / 64
                calcC3(3) = 5 / (128)
                calcC3(4) = ((n - 3) * n + 2) / 32
                calcC3(5) = (3 - 2 * n) / 64
                calcC3(6) = 3 / (128)
                calcC3(7) = (5 - 9 * n) / 192
                calcC3(8) = 3 / (128)
                calcC3(9) = 7 / (512)
                Return calcC3
            Case 6
                calcC3(0) = (1 - n) / 4
                calcC3(1) = (1 - n * n) / 8
                calcC3(2) = ((3 - n) * n + 3) / 64
                calcC3(3) = (2 * n + 5) / 128
                calcC3(4) = 3 / (128)
                calcC3(5) = ((n - 3) * n + 2) / 32
                calcC3(6) = ((-3 * n - 2) * n + 3) / 64
                calcC3(7) = (n + 3) / 128
                calcC3(8) = 5 / (256)
                calcC3(9) = (n * (5 * n - 9) + 5) / 192
                calcC3(10) = (9 - 10 * n) / 384
                calcC3(11) = 7 / (512)
                calcC3(12) = (7 - 14 * n) / 512
                calcC3(13) = 7 / (512)
                calcC3(14) = 21 / (2560)
                Return calcC3
            Case 7
                calcC3(0) = (1 - n) / 4
                calcC3(1) = (1 - n * n) / 8
                calcC3(2) = (n * ((-5 * n - 1) * n + 3) + 3) / 64
                calcC3(3) = (n * (2 * n + 2) + 5) / 128
                calcC3(4) = (11 * n + 12) / 512
                calcC3(5) = 21 / (1024)
                calcC3(6) = ((n - 3) * n + 2) / 32
                calcC3(7) = (n * (n * (2 * n - 3) - 2) + 3) / 64
                calcC3(8) = ((2 - 9 * n) * n + 6) / 256
                calcC3(9) = (n + 5) / 256
                calcC3(10) = 27 / (2048)
                calcC3(11) = (n * ((5 - n) * n - 9) + 5) / 192
                calcC3(12) = ((-6 * n - 10) * n + 9) / 384
                calcC3(13) = (21 - 4 * n) / 1536
                calcC3(14) = 3 / (256)
                calcC3(15) = (n * (10 * n - 14) + 7) / 512
                calcC3(16) = (7 - 10 * n) / 512
                calcC3(17) = 9 / (1024)
                calcC3(18) = (21 - 45 * n) / 2560
                calcC3(19) = 9 / (1024)
                calcC3(20) = 11 / (2048)
                Return calcC3
            Case 8
                calcC3(0) = (1 - n) / 4
                calcC3(1) = (1 - n * n) / 8
                calcC3(2) = (n * ((-5 * n - 1) * n + 3) + 3) / 64
                calcC3(3) = (n * ((2 - 2 * n) * n + 2) + 5) / 128
                calcC3(4) = (n * (3 * n + 11) + 12) / 512
                calcC3(5) = (10 * n + 21) / 1024
                calcC3(6) = 243 / (16384)
                calcC3(7) = ((n - 3) * n + 2) / 32
                calcC3(8) = (n * (n * (2 * n - 3) - 2) + 3) / 64
                calcC3(9) = (n * ((-6 * n - 9) * n + 2) + 6) / 256
                calcC3(10) = ((1 - 2 * n) * n + 5) / 256
                calcC3(11) = (69 * n + 108) / 8192
                calcC3(12) = 187 / (16384)
                calcC3(13) = (n * ((5 - n) * n - 9) + 5) / 192
                calcC3(14) = (n * (n * (10 * n - 6) - 10) + 9) / 384
                calcC3(15) = ((-77 * n - 8) * n + 42) / 3072
                calcC3(16) = (12 - n) / 1024
                calcC3(17) = 139 / (16384)
                calcC3(18) = (n * ((20 - 7 * n) * n - 28) + 14) / 1024
                calcC3(19) = ((-7 * n - 40) * n + 28) / 2048
                calcC3(20) = (72 - 43 * n) / 8192
                calcC3(21) = 127 / (16384)
                calcC3(22) = (n * (75 * n - 90) + 42) / 5120
                calcC3(23) = (9 - 15 * n) / 1024
                calcC3(24) = 99 / (16384)
                calcC3(25) = (44 - 99 * n) / 8192
                calcC3(26) = 99 / (16384)
                calcC3(27) = 429 / (114688)
                Return calcC3
            Case Else
                Return calcC3
        End Select
    End Function

    'The scale factor A3 = mean value of I3
    Private Function A3coeff(order As Integer) As Double()
        'Found the ellipsoid costants
        Dim f1, e2, n As Double
        f1 = 1 - Me.Flattening
        e2 = Me.Flattening * (2 - Me.Flattening)
        'ep2 = (e2 / f1 ^ 2)
        n = Me.Flattening / (2 - Me.Flattening)

        Dim A3index As Integer

        If order <= 0 Or order > 8 Then
            A3index = 0
        Else
            A3index = CInt((order * (order - 1)) / 2)
        End If

        Dim calcA3(A3index) As Double


        '  // The scale factor A3 = mean value of I3
        Select Case order
            Case 1
                calcA3(0) = 1
                Return calcA3
            Case 2
                calcA3(0) = 1
                calcA3(1) = -1 / 2
                Return calcA3
            Case 3
                calcA3(0) = 1
                calcA3(1) = (n - 1) / 2
                calcA3(2) = -1 / 4
                Return calcA3
            Case 4
                calcA3(0) = 1
                calcA3(1) = (n - 1) / 2
                calcA3(2) = (-n - 2) / 8
                calcA3(3) = -1 / 16
                Return calcA3
            Case 5
                calcA3(0) = 1
                calcA3(1) = (n - 1) / 2
                calcA3(2) = (n * (3 * n - 1) - 2) / 8
                calcA3(3) = (-3 * n - 1) / 16
                calcA3(4) = -3 / 64
                Return calcA3
            Case 6
                calcA3(0) = 1
                calcA3(1) = (n - 1) / 2
                calcA3(2) = (n * (3 * n - 1) - 2) / 8
                calcA3(3) = ((-n - 3) * n - 1) / 16
                calcA3(4) = (-2 * n - 3) / 64
                calcA3(5) = -3 / 128
                Return calcA3
            Case 7
                calcA3(0) = 1
                calcA3(1) = (n - 1) / 2
                calcA3(2) = (n * (3 * n - 1) - 2) / 8
                calcA3(3) = (n * (n * (5 * n - 1) - 3) - 1) / 16
                calcA3(4) = ((-10 * n - 2) * n - 3) / 64
                calcA3(5) = (-5 * n - 3) / 128
                calcA3(6) = -5 / 256
                Return calcA3
            Case 8
                calcA3(0) = 1
                calcA3(1) = (n - 1) / 2
                calcA3(2) = (n * (3 * n - 1) - 2) / 8
                calcA3(3) = (n * (n * (5 * n - 1) - 3) - 1) / 16
                calcA3(4) = (n * ((-5 * n - 20) * n - 4) - 6) / 128
                calcA3(5) = ((-5 * n - 10) * n - 6) / 256
                calcA3(6) = (-15 * n - 20) / 1024
                calcA3(7) = -25 / 2048
                Return calcA3
            Case Else
                Return calcA3
        End Select
    End Function

    'Evaluation sum(_A3c[k] * eps^k, k, 0, nA3x_-1) by Horner's method
    Private Function A3f(ByVal eps As Double, order As Integer) As Double
        Dim a3coeffs() As Double = A3coeff(order)
        Dim v As Double = 0
        Dim i As Integer = order
        While i > 0
            i -= 1
            v = eps * v + a3coeffs(i)
        End While
        Return v
    End Function

    'Evaluation C3 coeffs by Horner's method
    Private Function C3f(ByVal eps As Double, order As Integer) As Double()
        Dim c3coeffs() As Double = C3coeff(order)
        Dim j As Integer = c3coeffs.GetUpperBound(0)
        Dim k As Integer = order - 1
        Dim c(order) As Double
        While k > 0
            Dim t As Double = 0
            Dim i As Integer = order - k
            While i > 0
                j -= 1
                t = eps * t + c3coeffs(j)
                i -= 1
            End While
            c(k) = t
            k = k - 1
        End While
        Dim mult As Double = 1
        k = 1
        While k < order
            mult *= eps
            c(k) *= mult
            k = k + 1
        End While
        Return c
    End Function

    'Evaluation C4 coeffs by Horner's method
    Private Function C4f(ByVal k2 As Double, order As Integer) As Double()
        Dim c4coeffs() As Double = C4coeff(order)
        Dim j As Integer = c4coeffs.GetUpperBound(0)
        Dim k As Integer = order
        Dim c(order) As Double
        While k > 0
            Dim t As Double = 0
            Dim i As Integer = order - k + 1
            While i > 0
                j -= 1
                t = k2 * t + c4coeffs(j)
                i -= 1
            End While
            k = k - 1
            c(k) = t
        End While

        Dim mult As Double = 1
        k = 1
        While k < order
            mult *= k2
            c(k) *= mult
            k = k + 1
        End While
        Return c
    End Function

    'The scale factor A1-1 = mean value of I1-1
    Private Function A1m1f(ByVal eps As Double, order As Integer) As Double
        Dim eps2 As Double = eps ^ 2
        Dim t As Double
        Select Case CInt(order / 2)
            Case 0
                t = 0
            Case 1
                t = eps2 / 4
            Case 2
                t = eps2 * (eps2 + 16) / 64
            Case 3
                t = eps2 * (eps2 * (eps2 + 4) + 64) / 256
            Case 4
                t = eps2 * (eps2 * (eps2 * (25 * eps2 + 64) + 256) + 4096) / 16384
            Case Else
                t = 0
        End Select
        Return (t + eps) / (1 - eps)
    End Function

    'The coefficients C1[l] in the Fourier expansion of B1
    Private Function C1f(ByVal eps As Double, order As Integer) As Double()
        Dim eps2 As Double = eps ^ 2
        Dim d As Double = eps
        Dim c(order) As Double
        Select Case order
            Case 1
                c(1) = -d / 2
            Case 2
                c(1) = -d / 2
                d *= eps
                c(2) = -d / 16
            Case 3
                c(1) = d * (3 * eps2 - 8) / 16
                d *= eps
                c(2) = -d / 16
                d *= eps
                c(3) = -d / 48
            Case 4
                c(1) = d * (3 * eps2 - 8) / 16
                d *= eps
                c(2) = d * (eps2 - 2) / 32
                d *= eps
                c(3) = -d / 48
                d *= eps
                c(4) = -5 * d / 512
            Case 5
                c(1) = d * ((6 - eps2) * eps2 - 16) / 32
                d *= eps
                c(2) = d * (eps2 - 2) / 32
                d *= eps
                c(3) = d * (9 * eps2 - 16) / 768
                d *= eps
                c(4) = -5 * d / 512
                d *= eps
                c(5) = -7 * d / 1280
            Case 6
                c(1) = d * ((6 - eps2) * eps2 - 16) / 32
                d *= eps
                c(2) = d * ((64 - 9 * eps2) * eps2 - 128) / 2048
                d *= eps
                c(3) = d * (9 * eps2 - 16) / 768
                d *= eps
                c(4) = d * (3 * eps2 - 5) / 512
                d *= eps
                c(5) = -7 * d / 1280
                d *= eps
                c(6) = -7 * d / 2048
            Case 7
                c(1) = d * (eps2 * (eps2 * (19 * eps2 - 64) + 384) - 1024) / 2048
                d *= eps
                c(2) = d * ((64 - 9 * eps2) * eps2 - 128) / 2048
                d *= eps
                c(3) = d * ((72 - 9 * eps2) * eps2 - 128) / 6144
                d *= eps
                c(4) = d * (3 * eps2 - 5) / 512
                d *= eps
                c(5) = d * (35 * eps2 - 56) / 10240
                d *= eps
                c(6) = -7 * d / 2048
                d *= eps
                c(7) = -33 * d / 14336
            Case 8
                c(1) = d * (eps2 * (eps2 * (19 * eps2 - 64) + 384) - 1024) / 2048
                d *= eps
                c(2) = d * (eps2 * (eps2 * (7 * eps2 - 18) + 128) - 256) / 4096
                d *= eps
                c(3) = d * ((72 - 9 * eps2) * eps2 - 128) / 6144
                d *= eps
                c(4) = d * ((96 - 11 * eps2) * eps2 - 160) / 16384
                d *= eps
                c(5) = d * (35 * eps2 - 56) / 10240
                d *= eps
                c(6) = d * (9 * eps2 - 14) / 4096
                d *= eps
                c(7) = -33 * d / 14336
                d *= eps
                c(8) = -429 * d / 262144
        End Select
        Return c
    End Function

    'The coefficients C1p[l] in the Fourier expansion of B1p
    Private Function C1pf(ByVal eps As Double, order As Integer) As Double()
        Dim eps2 As Double = eps ^ 2
        Dim d As Double = eps
        Dim c(order) As Double
        Select Case order
            Case 1
                c(1) = d / 2
            Case 2
                c(1) = d / 2
                d *= eps
                c(2) = 5 * d / 16
            Case 3
                c(1) = d * (16 - 9 * eps2) / 32
                d *= eps
                c(2) = 5 * d / 16
                d *= eps
                c(3) = 29 * d / 96
            Case 4
                c(1) = d * (16 - 9 * eps2) / 32
                d *= eps
                c(2) = d * (30 - 37 * eps2) / 96
                d *= eps
                c(3) = 29 * d / 96
                d *= eps
                c(4) = 539 * d / 1536
            Case 5
                c(1) = d * (eps2 * (205 * eps2 - 432) + 768) / 1536
                d *= eps
                c(2) = d * (30 - 37 * eps2) / 96
                d *= eps
                c(3) = d * (116 - 225 * eps2) / 384
                d *= eps
                c(4) = 539 * d / 1536
                d *= eps
                c(5) = 3467 * d / 7680
            Case 6
                c(1) = d * (eps2 * (205 * eps2 - 432) + 768) / 1536
                d *= eps
                c(2) = d * (eps2 * (4005 * eps2 - 4736) + 3840) / 12288
                d *= eps
                c(3) = d * (116 - 225 * eps2) / 384
                d *= eps
                c(4) = d * (2695 - 7173 * eps2) / 7680
                d *= eps
                c(5) = 3467 * d / 7680
                d *= eps
                c(6) = 38081 * d / 61440
            Case 7
                c(1) = d * (eps2 * ((9840 - 4879 * eps2) * eps2 - 20736) + 36864) / 73728
                d *= eps
                c(2) = d * (eps2 * (4005 * eps2 - 4736) + 3840) / 12288
                d *= eps
                c(3) = d * (eps2 * (8703 * eps2 - 7200) + 3712) / 12288
                d *= eps
                c(4) = d * (2695 - 7173 * eps2) / 7680
                d *= eps
                c(5) = d * (41604 - 141115 * eps2) / 92160
                d *= eps
                c(6) = 38081 * d / 61440
                d *= eps
                c(7) = 459485 * d / 516096
            Case 8
                c(1) = d * (eps2 * ((9840 - 4879 * eps2) * eps2 - 20736) + 36864) / 73728
                d *= eps
                c(2) = d * (eps2 * ((120150 - 86171 * eps2) * eps2 - 142080) + 115200) / 368640
                d *= eps
                c(3) = d * (eps2 * (8703 * eps2 - 7200) + 3712) / 12288
                d *= eps
                c(4) = d * (eps2 * (1082857 * eps2 - 688608) + 258720) / 737280
                d *= eps
                c(5) = d * (41604 - 141115 * eps2) / 92160
                d *= eps
                c(6) = d * (533134 - 2200311 * eps2) / 860160
                d *= eps
                c(7) = 459485 * d / 516096
                d *= eps
                c(8) = 109167851 * d / 82575360
        End Select
        Return c
    End Function

    'The scale factor A2-1 = mean value of I2-1
    Private Function A2m1f(ByVal eps As Double, order As Integer) As Double
        Dim eps2 As Double = eps ^ 2
        Dim t As Double
        Select Case CInt(order / 2)
            Case 0
                t = 0
            Case 1
                t = eps2 / 4
            Case 2
                t = eps2 * (9 * eps2 + 16) / 64
            Case 3
                t = eps2 * (eps2 * (25 * eps2 + 36) + 64) / 256
            Case 4
                t = eps2 * (eps2 * (eps2 * (1225 * eps2 + 1600) + 2304) + 4096) / 16384
            Case Else
                t = 0
        End Select
        Return t * (1 - eps) - eps
    End Function

    'The coefficients C2[l] in the Fourier expansion of B2
    Private Function C2f(ByVal eps As Double, order As Integer) As Double()
        Dim eps2 As Double = eps ^ 2
        Dim d As Double = eps
        Dim c(order) As Double
        Select Case order
            Case 1
                c(1) = d / 2
            Case 2
                c(1) = d / 2
                d *= eps
                c(2) = 3 * d / 16
            Case 3
                c(1) = d * (eps2 + 8) / 16
                d *= eps
                c(2) = 3 * d / 16
                d *= eps
                c(3) = 5 * d / 48
            Case 4
                c(1) = d * (eps2 + 8) / 16
                d *= eps
                c(2) = d * (eps2 + 6) / 32
                d *= eps
                c(3) = 5 * d / 48
                d *= eps
                c(4) = 35 * d / 512
            Case 5
                c(1) = d * (eps2 * (eps2 + 2) + 16) / 32
                d *= eps
                c(2) = d * (eps2 + 6) / 32
                d *= eps
                c(3) = d * (15 * eps2 + 80) / 768
                d *= eps
                c(4) = 35 * d / 512
                d *= eps
                c(5) = 63 * d / 1280
            Case 6
                c(1) = d * (eps2 * (eps2 + 2) + 16) / 32
                d *= eps
                c(2) = d * (eps2 * (35 * eps2 + 64) + 384) / 2048
                d *= eps
                c(3) = d * (15 * eps2 + 80) / 768
                d *= eps
                c(4) = d * (7 * eps2 + 35) / 512
                d *= eps
                c(5) = 63 * d / 1280
                d *= eps
                c(6) = 77 * d / 2048
            Case 7
                c(1) = d * (eps2 * (eps2 * (41 * eps2 + 64) + 128) + 1024) / 2048
                d *= eps
                c(2) = d * (eps2 * (35 * eps2 + 64) + 384) / 2048
                d *= eps
                c(3) = d * (eps2 * (69 * eps2 + 120) + 640) / 6144
                d *= eps
                c(4) = d * (7 * eps2 + 35) / 512
                d *= eps
                c(5) = d * (105 * eps2 + 504) / 10240
                d *= eps
                c(6) = 77 * d / 2048
                d *= eps
                c(7) = 429 * d / 14336
            Case 8
                c(1) = d * (eps2 * (eps2 * (41 * eps2 + 64) + 128) + 1024) / 2048
                d *= eps
                c(2) = d * (eps2 * (eps2 * (47 * eps2 + 70) + 128) + 768) / 4096
                d *= eps
                c(3) = d * (eps2 * (69 * eps2 + 120) + 640) / 6144
                d *= eps
                c(4) = d * (eps2 * (133 * eps2 + 224) + 1120) / 16384
                d *= eps
                c(5) = d * (105 * eps2 + 504) / 10240
                d *= eps
                c(6) = d * (33 * eps2 + 154) / 4096
                d *= eps
                c(7) = 429 * d / 14336
                d *= eps
                c(8) = 6435 * d / 262144
        End Select
        Return c
    End Function

End Class
