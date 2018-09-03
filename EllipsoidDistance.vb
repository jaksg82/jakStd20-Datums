Imports System.Math
Imports MathExt

Partial Public Class Ellipsoid

    Dim tiny As Double = 1.4916668528863944E-154 'Sqrt(Double.Parse("2.22507e-308"))
    Dim tol0 As Double = 0.000000000000000222045 'Double.Parse("2.22045e-016")
    Dim tol1 As Double = 200 * tol0
    Dim tol2 As Double = Sqrt(tol0)
    Dim xthresh As Double = 1000 * tol2
    Dim GeodOrd As Integer = 6

    Private Function SinCosNorm(ByRef SinX As Double, ByRef CosX As Double) As Boolean
        Dim r As Double = Hypot(SinX, CosX)
        SinX /= r
        CosX /= r
        Return True
    End Function

    Private Function SinCosSeries(ByVal SinP As Boolean, ByRef SinX As Double, ByRef CosX As Double, ByRef c() As Double, ByVal n As Integer) As Double
        Dim countC As Integer
        'Point to one beyond last element
        If SinP Then
            countC = n + 1
        Else
            countC = n
        End If
        Dim ar As Double = 2 * (CosX - SinX) * (CosX + SinX) ' 2 * cos(2 * x)
        'Accumulators for sum
        Dim y1 As Double = 0
        Dim y0 As Double
        If n = 1 Then
            y0 = countC - 1
        Else
            y0 = 0
        End If
        'Now n is even
        n = CInt(n / 2)
        While n > 0
            'Unroll loop x 2, so accumulators return to their original role
            countC -= 1
            y1 = ar * y0 - y1 + c(countC)
            countC -= 1
            y0 = ar * y1 - y0 + c(countC)
            n -= 1
        End While
        Return If(SinP, 2 * SinX * CosX * y0, CosX * (y0 - y1))
    End Function

    Private Function Lengths(ByVal eps As Double, ByVal sig12 As Double, ByVal ssig1 As Double, ByVal csig1 As Double,
                     ByVal ssig2 As Double, ByVal csig2 As Double, ByVal cbet1 As Double, ByVal cbet2 As Double,
                     ByRef s12b As Double, ByRef m12a As Double, ByRef m0 As Double, ByRef scalep As Boolean,
                     ByRef M12 As Double, ByRef M21 As Double, ByVal C1a() As Double, ByVal C2a() As Double) As Boolean
        Dim e2 As Double = Eccentricity ^ 2
        Dim f1 As Double = 1 - Flattening
        C1a = C1f(eps, GeodOrd)
        C2a = C2f(eps, GeodOrd)
        Dim A1m1 As Double = A1m1f(eps, GeodOrd)
        Dim AB1 As Double = (1 + A1m1) * (SinCosSeries(True, ssig2, csig2, C1a, GeodOrd) - SinCosSeries(True, ssig1, csig1, C1a, GeodOrd))
        Dim A2m1 As Double = A2m1f(eps, GeodOrd)
        Dim AB2 As Double = (1 + A2m1) * (SinCosSeries(True, ssig2, csig2, C2a, GeodOrd) - SinCosSeries(True, ssig1, csig1, C2a, GeodOrd))
        Dim cbet1sq As Double = cbet1 ^ 2
        Dim cbet2sq As Double = cbet2 ^ 2
        Dim w1 As Double = Sqrt(1 - e2 * cbet1sq)
        Dim w2 As Double = Sqrt(1 - e2 * cbet2sq)
        'Make sure it's OK to have repeated dummy arguments
        Dim m0x As Double = A1m1 - A2m1
        Dim J12 As Double = m0x * sig12 + (AB1 - AB2)
        m0 = m0x
        'Missing a factor of a
        m12a = (w2 * (csig1 * ssig2) - w1 * (ssig1 * csig2)) - f1 * csig1 * csig2 * J12
        'Missing a factor of b
        s12b = (1 + A1m1) * sig12 + AB1
        If scalep Then
            Dim csig12 As Double = csig1 * csig2 + ssig1 * ssig2
            J12 *= f1
            M12 = csig12 + (e2 * (cbet1sq - cbet2sq) * ssig2 / (w1 + w2) - csig2 * J12) * ssig1 / w1
            M21 = csig12 - (e2 * (cbet1sq - cbet2sq) * ssig1 / (w1 + w2) - csig1 * J12) * ssig2 / w2
        End If
        Return True
    End Function

    Private Function Astroid(ByVal X As Double, ByVal Y As Double) As Double
        Dim k As Double
        Dim p As Double = X ^ 2
        Dim q As Double = Y ^ 2
        Dim r As Double = (p + q - 1) / 6
        If Not (q = 0 And r <= 0) Then
            'Avoid possible division by zero when r = 0 by multiplying equations for s and t by r^3 and r, resp.
            Dim S As Double = p * q / 4
            Dim r2 As Double = r ^ 2
            Dim r3 As Double = r * r2
            'The discrimant of the quadratic equation for T3.  This is zero on the evolute curve p^(1/3)+q^(1/3) = 1
            Dim disc As Double = S * (S + 2 * r3)
            Dim u As Double = r
            If disc >= 0 Then
                Dim T3 As Double = S + r3
                'Pick the sign on the sqrt to maximize abs(T3).  This minimizes loss of precision due to cancellation.
                'The result is unchanged because of the way the T is used in definition of u.
                If T3 < 0 Then
                    T3 = T3 - Sqrt(disc)
                Else
                    T3 = T3 + Sqrt(disc)
                End If
                'N.B. NthRoot always returns the real root.  NthRoot(-8) = -2.
                Dim T As Double = NthRoot(T3, 3)
                'T can be zero; but then r2 / T -> 0.
                u += T + (If(Not (T = 0), r2 / T, 0))
            Else
                'T is complex, but the way u is defined the result is real.
                Dim ang As Double = Atan2(Sqrt(-disc), -(S + r3))
                'There are three possible cube roots.  We choose the root which avoids cancellation.
                'Note that disc < 0 implies that r < 0.
                u += 2 * r * Cos(ang / 3)
            End If
            Dim v As Double = Sqrt(u ^ 2 + q) 'Guaranteed positive
            'Avoid loss of accuracy when u < 0.
            Dim uv As Double = If(u < 0, q / (v - u), u + v) 'u+v, guaranteed positive
            Dim w As Double = (uv - q) / (2 * v)
            'Rearrange expression for k to avoid loss of accuracy due to subtraction.  Division by 0 not possible because uv > 0, w >= 0.
            k = uv / (Sqrt(uv + w ^ 2) + w) 'guaranteed positive
        Else
            'y = 0 with |x| <= 1.  Handle this case directly.
            'for y small, positive root is k = abs(y)/sqrt(1-x^2)
            k = 0
        End If
        Return k
    End Function

    'Return a starting point for Newton's method in salp1 and calp1 (function value is -1).
    'If Newton's method doesn't need to be used, return also salp2 and calp2 and function value is sig12.
    Private Function InverseStart(ByVal sbet1 As Double, ByVal cbet1 As Double, ByVal sbet2 As Double,
                          ByVal cbet2 As Double, ByVal lam12 As Double, ByRef salp1 As Double,
                          ByRef calp1 As Double, ByRef salp2 As Double, ByRef calp2 As Double,
                          ByVal C1a() As Double, ByVal C2a() As Double) As Double
        Dim e2 As Double = Eccentricity ^ 2
        Dim f1 As Double = 1 - Flattening
        Dim ep2 As Double = SecondEccentricity ^ 2
        Dim n As Double = Flattening / (2 - Flattening)
        Dim etol2 As Double = 10 * tol2 / Max(0.1, Sqrt(Abs(e2)))
        Dim sig12 As Double = -1 'return value
        Dim sbet12 As Double = sbet2 * cbet1 - cbet2 * sbet1
        Dim cbet12 As Double = cbet2 * cbet1 + sbet2 * sbet1
        Dim sbet12a As Double = sbet2 * cbet1 + cbet2 * sbet1
        Dim shortline As Boolean = cbet12 >= 0 And sbet12 < 0.5 And lam12 <= PI / 6
        Dim omg12 As Double = If(Not (shortline), lam12, lam12 / Sqrt(1 - e2 * ((cbet1 + cbet2) / 2) ^ 2))
        Dim somg12 As Double = Sin(omg12)
        Dim comg12 As Double = Cos(omg12)
        salp1 = cbet2 * somg12
        calp1 = If(comg12 >= 0, sbet12 + cbet2 * sbet1 * somg12 ^ 2 / (1 + comg12), sbet12a - cbet2 * sbet1 * somg12 ^ 2 / (1 - comg12))
        Dim ssig12 As Double = Hypot(salp1, calp1)
        Dim csig12 As Double = sbet1 * sbet2 + cbet1 * cbet2 * comg12
        If shortline And ssig12 < etol2 Then
            'Really short lines
            salp2 = cbet1 * somg12
            calp2 = sbet12 - cbet1 * sbet2 * somg12 ^ 2 / (1 + comg12)
            SinCosNorm(salp2, calp2)
            'Set return value
            sig12 = Atan2(ssig12, csig12)
        ElseIf csig12 >= 0 Or ssig12 >= 3 * Abs(Flattening) * PI * cbet1 ^ 2 Then
            'Nothing to do, zeroth order spherical approximation is OK
        Else
            'Scale lam12 and bet2 to x, y coordinate system where antipodal point
            'is at origin and singular point is at y = 0, x = -1.
            Dim x, y, lamscale, betscale As Double
            If Flattening >= 0 Then
                'In fact f == 0 does not get here
                ' x = dlong, y = dlat
                Dim k2 As Double = sbet1 ^ 2 * ep2
                Dim eps As Double = k2 / (2 * (1 + Sqrt(1 + k2)) + k2)
                lamscale = Flattening * cbet1 * A3f(eps, GeodOrd) * PI
                betscale = lamscale * cbet1
                x = (lam12 - PI) / lamscale
                y = sbet12a / betscale
            Else ' f < 0
                ' x = dlat, y = dlong
                Dim cbet12a As Double = cbet2 * cbet1 - sbet2 * sbet1
                Dim bet12a As Double = Atan2(sbet12a, cbet12a)
                Dim m12a, m0, dummy As Double
                Lengths(n, PI + bet12a, sbet1, -cbet1, sbet2, cbet2, cbet1, cbet2, dummy, m12a, m0, False, dummy, dummy, C1a, C2a)
                x = -1 + m12a / (f1 * cbet1 * cbet2 * m0 * PI)
                betscale = If(x < -0.01, sbet12a / x, -Flattening * cbet1 ^ 2 * PI)
                lamscale = betscale / cbet1
                y = (lam12 - PI) / lamscale
            End If
            If y > -tol1 And x > -1 - xthresh Then
                'strip near cut
                If Flattening >= 0 Then
                    salp1 = Min(1, -x)
                    calp1 = -Sqrt(1 - salp1 ^ 2)
                Else
                    calp1 = Max(If(x > -tol1, 0, -1), x)
                    salp1 = Sqrt(1 - calp1 ^ 2)
                End If
            Else
                ' Estimate alp1, by solving the astroid problem.
                '
                ' Could estimate alpha1 = theta + pi/2, directly, i.e.,
                '   calp1 = y/k; salp1 = -x/(1+k);  for _f >= 0
                '   calp1 = x/(1+k); salp1 = -y/k;  for _f < 0 (need to check)
                '
                ' However, it's better to estimate omg12 from astroid and use
                ' spherical formula to compute alp1.  This reduces the mean number of
                ' Newton iterations for astroid cases from 2.24 (min 0, max 6) to 2.12
                ' (min 0 max 5).  The changes in the number of iterations are as
                ' follows:
                '
                ' change percent
                '    1       5
                '    0      78
                '   -1      16
                '   -2       0.6
                '   -3       0.04
                '   -4       0.002
                '
                ' The histogram of iterations is (m = number of iterations estimating
                ' alp1 directly, n = number of iterations estimating via omg12, total
                ' number of trials = 148605):
                '
                '  iter    m      n
                '    0   148    186
                '    1 13046  13845
                '    2 93315 102225
                '    3 36189  32341
                '    4  5396      7
                '    5   455      1
                '    6    56      0
                '
                ' Because omg12 is near pi, estimate work with omg12a = pi - omg12
                Dim k As Double = Astroid(x, y)
                Dim omg12a As Double = lamscale * If(Flattening >= 0, -x * k / (1 + k), -y * (1 + k) / k)
                somg12 = Sin(omg12a)
                comg12 = -Cos(omg12a)
                'Update spherical estimate of alp1 using omg12 instead of lam12
                salp1 = cbet2 * somg12
                calp1 = sbet12a - cbet2 * sbet1 * somg12 ^ 2 / (1 - comg12)
            End If
        End If
        SinCosNorm(salp1, calp1)
        Return sig12
    End Function

    Private Function Lambda12(ByVal sbet1 As Double, ByVal cbet1 As Double, ByVal sbet2 As Double, ByVal cbet2 As Double,
                      ByVal salp1 As Double, ByVal calp1 As Double, ByRef salp2 As Double, ByRef calp2 As Double,
                      ByRef sig12 As Double, ByRef ssig1 As Double, ByRef csig1 As Double, ByRef ssig2 As Double,
                      ByRef csig2 As Double, ByRef eps As Double, ByRef domg12 As Double, ByVal diffp As Boolean,
                      ByRef dlam12 As Double, ByVal C1a() As Double, ByVal C2a() As Double, ByVal C3a() As Double) As Double
        If sbet1 = 0 And calp1 = 0 Then calp1 = -tiny
        'Break degeneracy of equatorial line.  This case has already been handled.

        ' sin(alp1) * cos(bet1) = sin(alp0)
        Dim salp0 As Double = salp1 * cbet1
        Dim calp0 As Double = Hypot(calp1, salp1 * sbet1)  ' calp0 > 0
        Dim e2 As Double = Eccentricity ^ 2
        Dim ep2 As Double = SecondEccentricity ^ 2

        Dim somg1, comg1, somg2, comg2, omg12, lam12 As Double
        ssig1 = sbet1
        somg1 = salp0 * sbet1
        comg1 = calp1 * cbet1
        csig1 = comg1
        SinCosNorm(ssig1, csig1)
        ' Enforce symmetries in the case abs(bet2) = -bet1.  Need to be careful
        ' about this case, since this can yield singularities in the Newton iteration.
        ' sin(alp2) * cos(bet2) = sin(alp0)
        salp2 = If(Not (cbet2 = cbet1), salp0 / cbet2, salp1)
        ' calp2 = sqrt(1 - sq(salp2))
        '       = sqrt(sq(calp0) - sq(sbet2)) / cbet2
        ' and subst for calp0 and rearrange to give (choose positive sqrt
        ' to give alp2 in [0, pi/2]).
        'calp2 = If(cbet2 <> cbet1 OrElse Abs(sbet2) <> -sbet1, Sqrt((calp1 * cbet1) ^ 2 + (If(cbet1 < -sbet1, (cbet2 - cbet1) * (cbet1 + cbet2), (sbet1 - sbet2) * (sbet1 + sbet2)))) / cbet2, Abs(calp1))
        If cbet2 <> cbet1 Or Abs(sbet2) <> -sbet1 Then
            If cbet1 < -sbet1 Then
                calp2 = Sqrt((calp1 * cbet1) ^ 2 + ((cbet2 - cbet1) * (cbet1 + cbet2))) / cbet2
            Else
                calp2 = Sqrt((calp1 * cbet1) ^ 2 + ((sbet1 - sbet2) * (sbet1 + sbet2))) / cbet2
            End If
        Else
            calp2 = Abs(calp1)
        End If
        ssig2 = sbet2
        somg2 = salp0 * sbet2
        comg2 = calp2 * cbet2
        csig2 = comg2
        SinCosNorm(ssig2, csig2)
        ' sig12 = sig2 - sig1, limit to [0, pi]
        sig12 = Atan2(Max(csig1 * ssig2 - ssig1 * csig2, 0.0), csig1 * csig2 + ssig1 * ssig2)
        ' omg12 = omg2 - omg1, limit to [0, pi]
        omg12 = Atan2(Max(comg1 * somg2 - somg1 * comg2, 0.0), comg1 * comg2 + somg1 * somg2)
        Dim B312, h0 As Double
        Dim k2 As Double = calp0 ^ 2 * ep2
        eps = k2 / (2 * (1 + Sqrt(1 + k2)) + k2)
        C3a = C3f(eps, GeodOrd)
        B312 = (SinCosSeries(True, ssig2, csig2, C3a, GeodOrd - 1) - SinCosSeries(True, ssig1, csig1, C3a, GeodOrd - 1))
        h0 = -Flattening * A3f(eps, GeodOrd)
        domg12 = salp0 * h0 * (sig12 + B312)
        lam12 = omg12 + domg12

        If diffp Then
            If calp2 = 0 Then
                dlam12 = -2 * Sqrt(1 - e2 * cbet1 ^ 2) / sbet1
            Else
                Dim dummy As Double
                Lengths(eps, sig12, ssig1, csig1, ssig2, csig2, cbet1, cbet2, dummy, dlam12, dummy, False, dummy, dummy, C1a, C2a)
                dlam12 /= calp2 * cbet2
            End If
        End If
        Return lam12
    End Function

    'Public Function GenInverse(ByVal Lat1 As Double, ByVal Lon1 As Double, ByVal Lat2 As Double, ByVal Lon2 As Double,
    '                           ByVal outmask As UInteger, ByRef s12 As Double, ByRef azi1 As Double, ByRef azi2 As Double, ByRef m12 As Double,
    '                           ByRef Mm12 As Double, ByRef Mm21 As Double, ByRef Ss12 As Double) As Double
    '    'Found the ellipsoid costants
    '    a = Me.SemiMayorAxis
    '    f = Me.Flattening
    '    f1 = 1 - f
    '    e2 = f * (2 - f)
    '    ep2 = (e2 / f1 ^ 2)
    '    n = f / (2 - f)
    '    b = a * f1
    '    If e2 = 0 Then
    '        c2 = a ^ 2 + b ^ 2
    '    Else
    '        If e2 > 0 Then
    '            c2 = a ^ 2 + b ^ 2 * JakMathLib.ATanH(Sqrt(e2))
    '        Else
    '            c2 = a ^ 2 + b ^ 2 * Atan(Sqrt(-e2))
    '        End If
    '    End If
    '    etol2 = 10 * tol2 / Max(0.1, Sqrt(Abs(e2)))
    '    A3coeff()
    '    C3coeff()
    '    C4coeff()
    '    If outmask = 0 Then outmask = &H7F80UI
    '    Dim GEODESICSCALE As UInteger = 1UI
    '    Dim DISTANCE As UInteger = 1UI
    '    Dim REDUCEDLENGTH As UInteger = 1UI
    '    Dim AREA As UInteger = 1UI
    '    Dim AZIMUTH As UInteger = 1UI
    '    Dim Lon12 As Double
    '    Dim LonSign, LatSign As Integer

    '    Lon1 = JakMathLib.RadDeg(JakMathLib.AngleFit1Pi(JakMathLib.DegRad(Lon1)))
    '    Lon12 = JakMathLib.RadDeg(JakMathLib.AngleFit1Pi(JakMathLib.AngleFit1Pi(JakMathLib.DegRad(Lon2)) - JakMathLib.DegRad(Lon1)))
    '    'Lon12 = AngRound(Lon12)
    '    'Make the longitude difference positive
    '    If Lon12 < 0 Then
    '        Lon12 = -Lon12
    '        LonSign = -1
    '    Else
    '        LonSign = 1
    '    End If
    '    'If really close to the equator, treat as on equator.
    '    'Lat1 = AngRound(Lat1)
    '    'Lat2 = AngRound(Lat2)
    '    'Swap points so that point with higher abs latitude is point 1
    '    Dim swapp As Integer = If(Abs(Lat1) >= Abs(Lat2), 1, -1)
    '    If swapp < 0 Then
    '        LonSign *= -1
    '        Swap(Lat1, Lat2)
    '    End If
    '    'Make Lat1 <= 0
    '    If Lat1 < 0 Then
    '        LatSign = 1
    '    Else
    '        LatSign = -1
    '    End If
    '    Lat1 = Lat1 * LatSign
    '    Lat2 = Lat2 * LatSign
    '    'Now we have
    '    '    0 <= Lon12 <= 180
    '    '    -90 <= Lat1 <= 0
    '    '    Lat1 <= Lat2 <= -lat1
    '    Dim phi, sbet1, cbet1, sbet2, cbet2, s12x, m12x As Double
    '    phi = Lat1 * (PI / 180)
    '    sbet1 = f1 * Sin(phi)
    '    'Ensure cbet1 = +epsilon at poles
    '    If Lat1 = -90 Then
    '        cbet1 = tiny
    '    Else
    '        cbet1 = Cos(phi)
    '    End If
    '    phi = Lat2 * (PI / 180)
    '    SinCosNorm(sbet1, cbet1)
    '    sbet2 = f1 * Sin(phi)
    '    'Ensure cbet2 = +epsilon at poles
    '    If Abs(Lat2) = 90 Then
    '        cbet2 = tiny
    '    Else
    '        cbet2 = Cos(phi)
    '    End If
    '    SinCosNorm(sbet2, cbet2)
    '    ' If cbet1 < -sbet1, then cbet2 - cbet1 is a sensitive measure of the
    '    ' |bet1| - |bet2|.  Alternatively (cbet1 >= -sbet1), abs(sbet2) + sbet1 is
    '    ' a better measure.  This logic is used in assigning calp2 in Lambda12.
    '    ' Sometimes these quantities vanish and in that case we force bet2 = +/-
    '    ' bet1 exactly.  An example where is is necessary is the inverse problem
    '    ' 48.522876735459 0 -48.52287673545898293 179.599720456223079643
    '    ' which failed with Visual Studio 10 (Release and Debug)
    '    If cbet1 < -sbet1 Then
    '        If cbet2 = cbet1 Then sbet2 = If(sbet2 < 0, sbet1, -sbet1)
    '    Else
    '        If Abs(sbet2) = -sbet1 Then cbet2 = cbet1
    '    End If

    '    Dim lam12 As Double = Lon12 * (PI / 180)
    '    Dim slam12 As Double = If(Lon12 = 180, 0, Sin(lam12))
    '    Dim clam12 As Double = Cos(lam12)  'lon12 = 90 isn't interesting
    '    Dim a12, sig12, calp1, salp1, calp2, salp2 As Double
    '    Dim C1a(nC1 + 1), C2a(nC2 + 1), C3a(nC3) As Double
    '    Dim meridian As Boolean = (Lat1 = -90) Or (slam12 = 0)
    '    If meridian Then
    '        'End points are on a single full meridian, so the geodesic might lie on a meridian.
    '        'Head to the target longitude
    '        calp1 = clam12
    '        salp1 = slam12
    '        'At the target we're heading north
    '        calp2 = 1
    '        salp2 = 0
    '        'tan(bet) = tan(sig) * cos(alp)
    '        Dim ssig1 As Double = sbet1
    '        Dim csig1 As Double = calp1 * cbet1
    '        Dim ssig2 As Double = sbet2
    '        Dim csig2 As Double = calp2 * cbet2
    '        'sig12 = sig2 - sig1
    '        sig12 = Atan2(Max(csig1 * ssig2 - ssig1 * csig2, 0), csig1 * csig2 + ssig1 * ssig2)
    '        Dim dummy As Double
    '        Lengths(n, sig12, ssig1, csig1, ssig2, csig2, cbet1, cbet2, s12x, m12x, dummy, True, Mm12, Mm21, C1a, C2a)
    '        ' Add the check for sig12 since zero length geodesics might yield m12 < 0.  Test case was
    '        '
    '        '    echo 20.001 0 20.001 0 | Geod -i
    '        '
    '        ' In fact, we will have sig12 > pi/2 for meridional geodesic which is not a shortest path.
    '        If sig12 < 1 Or m12x >= 0 Then
    '            m12x *= a
    '            s12x *= b
    '            a12 = sig12 / (PI / 180)
    '        Else
    '            ' m12 < 0, i.e., prolate and too close to anti-podal
    '            meridian = False
    '        End If
    '    End If
    '    Dim omg12 As Double
    '    If Not (meridian) And sbet1 = 0 And (f <= 0 Or lam12 <= PI - f * PI) Then
    '        ' Mimic the way Lambda12 works with calp1 = 0
    '        ' Geodesic runs along equator
    '        calp1 = 0
    '        calp2 = 0
    '        salp1 = 1
    '        salp2 = 1
    '        s12x = a * lam12
    '        m12x = b * Sin(lam12 / f1)
    '        If outmask <> 0UI Then
    '            Mm12 = Cos(lam12 / f1)
    '            Mm21 = Mm12
    '        End If
    '        a12 = Lon12 / f1
    '        sig12 = lam12 / f1
    '        omg12 = sig12
    '    ElseIf Not (meridian) Then
    '        'Now point1 and point2 belong within a hemisphere bounded by a
    '        'meridian and geodesic is neither meridional or equatorial.
    '        'Figure a starting point for Newton's method
    '        sig12 = InverseStart(sbet1, cbet1, sbet2, cbet2, lam12, salp1, calp1, salp2, calp2, C1a, C2a)
    '        If sig12 >= 0 Then
    '            ' Short lines (InverseStart sets salp2, calp2)
    '            Dim wm As Double = Sqrt(1 - e2 * ((cbet1 + cbet2) / 2) ^ 2)
    '            s12x = sig12 * a * wm
    '            m12x = wm ^ 2 * a / f1 * Sin(sig12 * f1 / wm)
    '            If outmask <> 0UI Then
    '                Mm12 = Cos(sig12 * f1 / wm)
    '                Mm21 = Mm12
    '            End If
    '            a12 = sig12 / (PI / 180)
    '            omg12 = lam12 / wm
    '        Else
    '            ' Newton's method
    '            Dim ssig1, csig1, ssig2, csig2, eps, ov As Double
    '            ov = 0
    '            Dim numit As UInteger = 0
    '            Dim trip As UInteger = 0
    '            While numit < MaxIt
    '                numit += 1UI
    '                Dim dv, v As Double
    '                v = Lambda12(sbet1, cbet1, sbet2, cbet2, salp1, calp1, salp2, calp2, sig12, ssig1,
    '                             csig1, ssig2, csig2, eps, omg12, trip < 1, dv, C1a, C2a, C3a) - lam12
    '                If Not (Abs(v) > tiny) Or Not (trip < 1) Then
    '                    If Not (Abs(v) <= Max(tol1, ov)) Then
    '                        numit = MaxIt
    '                    End If
    '                    Exit While
    '                End If
    '                Dim dalp1 As Double = -v / dv
    '                Dim sdalp1 As Double = Sin(dalp1)
    '                Dim cdalp1 As Double = Cos(dalp1)
    '                Dim nsalp1 As Double = salp1 * cdalp1 + calp1 * sdalp1
    '                calp1 = calp1 * cdalp1 - salp1 * sdalp1
    '                salp1 = Max(0.0, nsalp1)
    '                SinCosNorm(salp1, calp1)
    '                ' In some regimes we don't get quadratic convergence because slope
    '                ' -> 0.  So use convergence conditions based on epsilon instead of
    '                ' sqrt(epsilon).  The first criterion is a test on abs(v) against
    '                ' 100 * epsilon.  The second takes credit for an anticipated
    '                ' reduction in abs(v) by v/ov (due to the latest update in alp1) and
    '                ' checks this against epsilon.
    '                If Not (Abs(v) >= tol1 And v ^ 2 >= ov * tol0) Then
    '                    trip += 1UI
    '                End If
    '                ov = Abs(v)
    '            End While

    '            If numit >= MaxIt Then
    '                'Signal failure.
    '                s12 = Double.NaN
    '                azi1 = Double.NaN
    '                azi2 = Double.NaN
    '                m12 = Double.NaN
    '                Mm12 = Double.NaN
    '                Mm21 = Double.NaN
    '                Ss12 = Double.NaN
    '                Return Double.NaN
    '            End If

    '            Dim dummy As Double
    '            Lengths(eps, sig12, ssig1, csig1, ssig2, csig2, cbet1, cbet2, s12x, m12x, dummy, True, Mm12, Mm21, C1a, C2a)
    '            m12x *= a
    '            s12x *= b
    '            a12 = sig12 / (PI / 180)
    '            omg12 = lam12 - omg12
    '        End If
    '    End If
    '    'If (outmask And DISTANCE) <> 0UI Then s12 = 0 + s12x ' Convert -0 to 0
    '    'If (outmask And REDUCEDLENGTH) <> 0UI Then m12 = 0 + m12x ' Convert -0 to 0
    '    'If (outmask And AREA) <> 0UI Then
    '    If outmask <> 0UI Then
    '        m12 = Abs(m12x)
    '        s12 = Abs(s12x)
    '        'From Lambda12: sin(alp1) * cos(bet1) = sin(alp0)
    '        Dim salp0 As Double = salp1 * cbet1
    '        Dim calp0 As Double = JakMathLib.Hypot(calp1, salp1 * sbet1)
    '        Dim alp12 As Double
    '        If calp0 <> 0 And salp0 <> 0 Then
    '            'From Lambda12: tan(bet) = tan(sig) * cos(alp)
    '            Dim ssig1 As Double = sbet1
    '            Dim csig1 As Double = calp1 * cbet1
    '            Dim ssig2 As Double = sbet2
    '            Dim csig2 As Double = calp2 * cbet2
    '            Dim k2 As Double = calp0 ^ 2 * ep2
    '            'Multiplier = a^2 * e^2 * cos(alpha0) * sin(alpha0)
    '            Dim A4 As Double = a ^ 2 * calp0 * salp0 * e2
    '            SinCosNorm(ssig1, csig1)
    '            SinCosNorm(ssig2, csig2)
    '            Dim C4a(nC4) As Double
    '            C4f(k2, C4a)
    '            Dim B41 As Double = SinCosSeries(False, ssig1, csig1, C4a, nC4)
    '            Dim B42 As Double = SinCosSeries(False, ssig2, csig2, C4a, nC4)
    '            Ss12 = A4 * (B42 - B41)
    '        Else
    '            'Avoid problems with indeterminate sig1, sig2 on equator
    '            Ss12 = 0
    '        End If

    '        If Not (meridian) And omg12 < 0.75 * PI And sbet2 - sbet1 < 1.75 Then
    '            'Long difference too big  &  Lat difference too big
    '            ' Use tan(Gamma/2) = tan(omg12/2)
    '            ' * (tan(bet1/2)+tan(bet2/2))/(1+tan(bet1/2)*tan(bet2/2))
    '            ' with tan(x/2) = sin(x)/(1+cos(x))
    '            Dim somg12 As Double = Sin(omg12)
    '            Dim domg12 As Double = 1 + Cos(omg12)
    '            Dim dbet1 As Double = 1 + cbet1
    '            Dim dbet2 As Double = 1 + cbet2
    '            alp12 = 2 * Atan2(somg12 * (sbet1 * dbet2 + sbet2 * dbet1), domg12 * (sbet1 * sbet2 + dbet1 * dbet2))
    '        Else
    '            ' alp12 = alp2 - alp1, used in atan2 so no need to normalize
    '            Dim salp12 As Double = salp2 * calp1 - calp2 * salp1
    '            Dim calp12 As Double = calp2 * calp1 + salp2 * salp1
    '            ' The right thing appears to happen if alp1 = +/-180 and alp2 = 0, viz salp12 = -0 and alp12 = -180.
    '            ' However this depends on the sign being attached to 0 correctly.
    '            ' The following ensures the correct behavior.
    '            If salp12 = 0 And calp12 < 0 Then
    '                salp12 = tiny * calp1
    '                calp12 = -1
    '            End If
    '            alp12 = Atan2(salp12, calp12)
    '        End If
    '        Ss12 += c2 * alp12
    '        Ss12 *= swapp * LonSign * LatSign
    '        ' Convert -0 to 0
    '        Ss12 = Abs(Ss12)
    '    End If

    '    ' Convert calp, salp to azimuth accounting for lonsign, swapp, latsign.
    '    If swapp < 0 Then
    '        Swap(salp1, salp2)
    '        Swap(calp1, calp2)
    '        If outmask <> 0UI Then Swap(Mm12, Mm21)
    '    End If

    '    salp1 *= swapp * LonSign
    '    calp1 *= swapp * LatSign
    '    salp2 *= swapp * LonSign
    '    calp2 *= swapp * LatSign

    '    'If (outmask And AZIMUTH) <> 0 Then
    '    If outmask <> 0 Then
    '        ' minus signs give range [-180, 180). 0- converts -0 to +0.
    '        azi1 = 0 - Atan2(-salp1, calp1) / (PI / 180)
    '        azi2 = 0 - Atan2(-salp2, calp2) / (PI / 180)
    '    End If

    '    ' Returned value in [0, 180]
    '    Return a12
    'End Function

    Public Function DistanceKarney(point1 As Point3D, ByVal point2 As Point3D, Optional ByRef alpha1 As Double = 0.0, Optional ByRef alpha2 As Double = 0.0) As Double
        Dim MaxIt As UInteger = 500

        'Lat1 As Double, ByVal Lon1 As Double, ByVal Lat2 As Double, ByVal Lon2 As Double,
        'ByVal outmask As UInteger, ByRef s12 As Double, ByRef azi1 As Double, ByRef azi2 As Double, ByRef m12 As Double,
        'ByRef Mm12 As Double, ByRef Mm21 As Double, ByRef Ss12 As Double
        '00604      * @param[in] lat1 latitude of point 1 (degrees).
        '00605      * @param[in] lon1 longitude of point 1 (degrees).
        '00606      * @param[in] lat2 latitude of point 2 (degrees).
        '00607      * @param[in] lon2 longitude of point 2 (degrees).
        '00608      * @param[out] s12 distance between point 1 And point 2 (meters).
        '00609      * @param[out] azi1 azimuth at point 1 (degrees).
        '00610      * @param[out] azi2 (forward) azimuth at point 2 (degrees).
        '00611      * @param[out] m12 reduced length of geodesic (meters).
        '00612      * @param[out] M12 geodesic scale of point 2 relative to point 1 (dimensionless).
        '00614      * @param[out] M21 geodesic scale of point 1 relative to point 2 (dimensionless).
        '00616      * @param[out] S12 area under the geodesic (meters<sup>2</sup>).
        '00617      * @return \e a12 arc length of between point 1 And point 2 (degrees).

        'Dim outmask As UInteger 'Output mask result
        Dim Lat1 As Double = RadDeg(point1.Y)
        Dim Lat2 As Double = RadDeg(point2.Y)

        Dim s12 As Double 'distance between point 1 And point 2 (meters)
        Dim azi1 As Double 'azimuth at point 1 (degrees)
        Dim azi2 As Double '(forward) azimuth at point 2 (degrees).
        Dim m12 As Double 'reduced length of geodesic (meters).
        Dim Mm12 As Double 'geodesic scale of point 2 relative to point 1 (dimensionless).
        Dim Mm21 As Double 'geodesic scale of point 1 relative to point 2 (dimensionless).
        Dim Ss12 As Double 'area under the geodesic (meters^2).

        'Found the ellipsoid costants
        Dim f1 As Double = 1 - Flattening
        Dim e2 As Double = Flattening * (2 - Flattening)
        Dim ep2 As Double = (e2 / f1 ^ 2)
        Dim n As Double = Flattening / (2 - Flattening)
        Dim c2 As Double
        'Dim etol2 As Double = 10 * tol2 / Max(0.1, Sqrt(Abs(e2)))
        If e2 = 0 Then
            c2 = SemiMayorAxis ^ 2 + SemiMinorAxis ^ 2
        Else
            If e2 > 0 Then
                c2 = SemiMayorAxis ^ 2 + SemiMinorAxis ^ 2 * ATanH(Sqrt(e2))
            Else
                c2 = SemiMayorAxis ^ 2 + SemiMinorAxis ^ 2 * Atan(Sqrt(-e2))
            End If
        End If

        'A3coeff()
        'C3coeff()
        'C4coeff()
        'If outmask = 0 Then outmask = &H7F80UI
        'Dim GEODESICSCALE As UInteger = 1UI
        'Dim DISTANCE As UInteger = 1UI
        'Dim REDUCEDLENGTH As UInteger = 1UI
        'Dim AREA As UInteger = 1UI
        'Dim AZIMUTH As UInteger = 1UI
        Dim Lon12 As Double
        Dim LonSign, LatSign As Integer

        'Lon1 = JakMathLib.RadDeg(JakMathLib.AngleFit1Pi(JakMathLib.DegRad(Lon1)))
        'Lon12 = JakMathLib.RadDeg(JakMathLib.AngleFit1Pi(JakMathLib.AngleFit1Pi(JakMathLib.DegRad(Lon2)) - JakMathLib.DegRad(Lon1)))
        Lon12 = RadDeg(AngleFit1Pi(point2.X - point1.X))

        'Make the longitude difference positive
        If Lon12 < 0 Then
            Lon12 = -Lon12
            LonSign = -1
        Else
            LonSign = 1
        End If
        'If really close to the equator, treat as on equator.
        'Lat1 = AngRound(Lat1)
        'Lat2 = AngRound(Lat2)
        'Swap points so that point with higher abs latitude is point 1
        Dim swapp As Integer = If(Abs(Lat1) >= Abs(Lat2), 1, -1)
        If swapp < 0 Then
            LonSign *= -1
            Swap(Lat1, Lat2)
        End If
        'Make Lat1 <= 0
        If Lat1 < 0 Then
            LatSign = 1
        Else
            LatSign = -1
        End If
        Lat1 = Lat1 * LatSign
        Lat2 = Lat2 * LatSign
        'Now we have
        '    0 <= Lon12 <= 180
        '    -90 <= Lat1 <= 0
        '    Lat1 <= Lat2 <= -lat1
        Dim phi, sbet1, cbet1, sbet2, cbet2, s12x, m12x As Double
        phi = Lat1 * (PI / 180)
        sbet1 = f1 * Sin(phi)
        'Ensure cbet1 = +epsilon at poles
        If Lat1 = -90 Then
            cbet1 = tiny
        Else
            cbet1 = Cos(phi)
        End If
        phi = Lat2 * (PI / 180)
        SinCosNorm(sbet1, cbet1)
        sbet2 = f1 * Sin(phi)
        'Ensure cbet2 = +epsilon at poles
        If Abs(Lat2) = 90 Then
            cbet2 = tiny
        Else
            cbet2 = Cos(phi)
        End If
        SinCosNorm(sbet2, cbet2)
        ' If cbet1 < -sbet1, then cbet2 - cbet1 is a sensitive measure of the
        ' |bet1| - |bet2|.  Alternatively (cbet1 >= -sbet1), abs(sbet2) + sbet1 is
        ' a better measure.  This logic is used in assigning calp2 in Lambda12.
        ' Sometimes these quantities vanish and in that case we force bet2 = +/-
        ' bet1 exactly.  An example where is is necessary is the inverse problem
        ' 48.522876735459 0 -48.52287673545898293 179.599720456223079643
        ' which failed with Visual Studio 10 (Release and Debug)
        If cbet1 < -sbet1 Then
            If cbet2 = cbet1 Then sbet2 = If(sbet2 < 0, sbet1, -sbet1)
        Else
            If Abs(sbet2) = -sbet1 Then cbet2 = cbet1
        End If

        Dim lam12 As Double = Lon12 * (PI / 180)
        Dim slam12 As Double = If(Lon12 = 180, 0, Sin(lam12))
        Dim clam12 As Double = Cos(lam12)  'lon12 = 90 isn't interesting
        Dim a12, sig12, calp1, salp1, calp2, salp2 As Double
        Dim C1a(GeodOrd + 1), C2a(GeodOrd + 1), C3a(GeodOrd) As Double
        Dim meridian As Boolean = (Lat1 = -90) Or (slam12 = 0)
        If meridian Then
            'End points are on a single full meridian, so the geodesic might lie on a meridian.
            'Head to the target longitude
            calp1 = clam12
            salp1 = slam12
            'At the target we're heading north
            calp2 = 1
            salp2 = 0
            'tan(bet) = tan(sig) * cos(alp)
            Dim ssig1 As Double = sbet1
            Dim csig1 As Double = calp1 * cbet1
            Dim ssig2 As Double = sbet2
            Dim csig2 As Double = calp2 * cbet2
            'sig12 = sig2 - sig1
            sig12 = Atan2(Max(csig1 * ssig2 - ssig1 * csig2, 0), csig1 * csig2 + ssig1 * ssig2)
            Dim dummy As Double
            Lengths(n, sig12, ssig1, csig1, ssig2, csig2, cbet1, cbet2, s12x, m12x, dummy, True, Mm12, Mm21, C1a, C2a)
            ' Add the check for sig12 since zero length geodesics might yield m12 < 0.  Test case was
            '
            '    echo 20.001 0 20.001 0 | Geod -i
            '
            ' In fact, we will have sig12 > pi/2 for meridional geodesic which is not a shortest path.
            If sig12 < 1 Or m12x >= 0 Then
                m12x *= SemiMayorAxis
                s12x *= SemiMinorAxis
                a12 = sig12 / (PI / 180)
            Else
                ' m12 < 0, i.e., prolate and too close to anti-podal
                meridian = False
            End If
        End If
        Dim omg12 As Double
        If Not (meridian) And sbet1 = 0 And (Flattening <= 0 Or lam12 <= PI - Flattening * PI) Then
            ' Mimic the way Lambda12 works with calp1 = 0
            ' Geodesic runs along equator
            calp1 = 0
            calp2 = 0
            salp1 = 1
            salp2 = 1
            s12x = SemiMayorAxis * lam12
            m12x = SemiMinorAxis * Sin(lam12 / f1)
            Mm12 = Cos(lam12 / f1)
            Mm21 = Mm12
            a12 = Lon12 / f1
            sig12 = lam12 / f1
            omg12 = sig12
        ElseIf Not (meridian) Then
            'Now point1 and point2 belong within a hemisphere bounded by a
            'meridian and geodesic is neither meridional or equatorial.
            'Figure a starting point for Newton's method
            sig12 = InverseStart(sbet1, cbet1, sbet2, cbet2, lam12, salp1, calp1, salp2, calp2, C1a, C2a)
            If sig12 >= 0 Then
                ' Short lines (InverseStart sets salp2, calp2)
                Dim wm As Double = Sqrt(1 - e2 * ((cbet1 + cbet2) / 2) ^ 2)
                s12x = sig12 * SemiMayorAxis * wm
                m12x = wm ^ 2 * SemiMayorAxis / f1 * Sin(sig12 * f1 / wm)
                Mm12 = Cos(sig12 * f1 / wm)
                Mm21 = Mm12
                a12 = sig12 / (PI / 180)
                omg12 = lam12 / wm
            Else
                ' Newton's method
                Dim ssig1, csig1, ssig2, csig2, eps, ov As Double
                ov = 0
                Dim numit As UInteger = 0
                Dim trip As UInteger = 0
                While numit < MaxIt
                    numit += 1UI
                    Dim dv, v As Double
                    v = Lambda12(sbet1, cbet1, sbet2, cbet2, salp1, calp1, salp2, calp2, sig12, ssig1,
                                 csig1, ssig2, csig2, eps, omg12, trip < 1, dv, C1a, C2a, C3a) - lam12
                    If Not (Abs(v) > tiny) Or Not (trip < 1) Then
                        If Not (Abs(v) <= Max(tol1, ov)) Then
                            numit = MaxIt
                        End If
                        Exit While
                    End If
                    Dim dalp1 As Double = -v / dv
                    Dim sdalp1 As Double = Sin(dalp1)
                    Dim cdalp1 As Double = Cos(dalp1)
                    Dim nsalp1 As Double = salp1 * cdalp1 + calp1 * sdalp1
                    calp1 = calp1 * cdalp1 - salp1 * sdalp1
                    salp1 = Max(0.0, nsalp1)
                    SinCosNorm(salp1, calp1)
                    ' In some regimes we don't get quadratic convergence because slope
                    ' -> 0.  So use convergence conditions based on epsilon instead of
                    ' sqrt(epsilon).  The first criterion is a test on abs(v) against
                    ' 100 * epsilon.  The second takes credit for an anticipated
                    ' reduction in abs(v) by v/ov (due to the latest update in alp1) and
                    ' checks this against epsilon.
                    If Not (Abs(v) >= tol1 And v ^ 2 >= ov * tol0) Then
                        trip += 1UI
                    End If
                    ov = Abs(v)
                End While

                If numit >= MaxIt Then
                    'Signal failure.
                    s12 = Double.NaN
                    azi1 = Double.NaN
                    azi2 = Double.NaN
                    m12 = Double.NaN
                    Mm12 = Double.NaN
                    Mm21 = Double.NaN
                    Ss12 = Double.NaN
                    Return Double.NaN
                End If

                Dim dummy As Double
                Lengths(eps, sig12, ssig1, csig1, ssig2, csig2, cbet1, cbet2, s12x, m12x, dummy, True, Mm12, Mm21, C1a, C2a)
                m12x *= SemiMayorAxis
                s12x *= SemiMinorAxis
                a12 = sig12 / (PI / 180)
                omg12 = lam12 - omg12
            End If
        End If

        m12 = Abs(m12x)
        s12 = Abs(s12x)
        'From Lambda12: sin(alp1) * cos(bet1) = sin(alp0)
        Dim salp0 As Double = salp1 * cbet1
        Dim calp0 As Double = Hypot(calp1, salp1 * sbet1)
        Dim alp12 As Double
        If calp0 <> 0 And salp0 <> 0 Then
            'From Lambda12: tan(bet) = tan(sig) * cos(alp)
            Dim ssig1 As Double = sbet1
            Dim csig1 As Double = calp1 * cbet1
            Dim ssig2 As Double = sbet2
            Dim csig2 As Double = calp2 * cbet2
            Dim k2 As Double = calp0 ^ 2 * ep2
            'Multiplier = a^2 * e^2 * cos(alpha0) * sin(alpha0)
            Dim A4 As Double = SemiMayorAxis ^ 2 * calp0 * salp0 * e2
            SinCosNorm(ssig1, csig1)
            SinCosNorm(ssig2, csig2)
            Dim C4a(GeodOrd) As Double
            C4a = C4f(k2, GeodOrd)
            Dim B41 As Double = SinCosSeries(False, ssig1, csig1, C4a, GeodOrd)
            Dim B42 As Double = SinCosSeries(False, ssig2, csig2, C4a, GeodOrd)
            Ss12 = A4 * (B42 - B41)
        Else
            'Avoid problems with indeterminate sig1, sig2 on equator
            Ss12 = 0
        End If

        If Not (meridian) And omg12 < 0.75 * PI And sbet2 - sbet1 < 1.75 Then
            'Long difference too big  &  Lat difference too big
            ' Use tan(Gamma/2) = tan(omg12/2)
            ' * (tan(bet1/2)+tan(bet2/2))/(1+tan(bet1/2)*tan(bet2/2))
            ' with tan(x/2) = sin(x)/(1+cos(x))
            Dim somg12 As Double = Sin(omg12)
            Dim domg12 As Double = 1 + Cos(omg12)
            Dim dbet1 As Double = 1 + cbet1
            Dim dbet2 As Double = 1 + cbet2
            alp12 = 2 * Atan2(somg12 * (sbet1 * dbet2 + sbet2 * dbet1), domg12 * (sbet1 * sbet2 + dbet1 * dbet2))
        Else
            ' alp12 = alp2 - alp1, used in atan2 so no need to normalize
            Dim salp12 As Double = salp2 * calp1 - calp2 * salp1
            Dim calp12 As Double = calp2 * calp1 + salp2 * salp1
            ' The right thing appears to happen if alp1 = +/-180 and alp2 = 0, viz salp12 = -0 and alp12 = -180.
            ' However this depends on the sign being attached to 0 correctly.
            ' The following ensures the correct behavior.
            If salp12 = 0 And calp12 < 0 Then
                salp12 = tiny * calp1
                calp12 = -1
            End If
            alp12 = Atan2(salp12, calp12)
        End If
        Ss12 += c2 * alp12
        Ss12 *= swapp * LonSign * LatSign
        ' Convert -0 to 0
        Ss12 = Abs(Ss12)

        ' Convert calp, salp to azimuth accounting for lonsign, swapp, latsign.
        If swapp < 0 Then
            Swap(salp1, salp2)
            Swap(calp1, calp2)
            Swap(Mm12, Mm21)
        End If

        salp1 *= swapp * LonSign
        calp1 *= swapp * LatSign
        salp2 *= swapp * LonSign
        calp2 *= swapp * LatSign

        'If (outmask And AZIMUTH) <> 0 Then
        ' minus signs give range [-180, 180). 0- converts -0 to +0.
        azi1 = 0 - Atan2(-salp1, calp1) / (PI / 180)
        azi2 = 0 - Atan2(-salp2, calp2) / (PI / 180)

        ' Returned values
        alpha1 = DegRad(azi1)
        alpha2 = DegRad(azi2)
        Return s12
    End Function

    Private Sub Swap(ByRef Value1 As Double, ByRef Value2 As Double)
        Dim Value3 As Double = Value1
        Value1 = Value2
        Value2 = Value3
    End Sub

End Class