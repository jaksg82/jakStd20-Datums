﻿Imports MathExt

Public Class Helmert7p
    Inherits Transformations

    Private iSrcEll, iTgtEll As Ellipsoid
    Private idx, idy, idz, irx, iry, irz, ippm As Double
    Private iRotConv As RotationConventions

    Overrides ReadOnly Property Type As Methods
        Get
            Return Methods.Helmert7Parameter
        End Get
    End Property

    Public Overrides Property RotationConvention As RotationConventions
        Get
            Return iRotConv
        End Get
        Set(value As RotationConventions)
            iRotConv = value
        End Set
    End Property

    Public Sub New(sourceEllipsoid As Ellipsoid, longName As String, compactName As String, deltaX As Double, deltaY As Double, deltaZ As Double,
                   rotationX As Double, rotationY As Double, rotationZ As Double, scaleDifference As Double,
                   Optional convention As RotationConventions = RotationConventions.PositionVector)
        FullName = longName
        ShortName = compactName
        iSrcEll = sourceEllipsoid
        iTgtEll = New Ellipsoid(7030)
        idx = deltaX
        idy = deltaY
        idz = deltaZ
        irx = rotationX
        iry = rotationY
        irz = rotationZ
        ippm = scaleDifference
        iRotConv = convention
    End Sub

    Public Sub New(sourceEllipsoid As Ellipsoid, targetEllipsoid As Ellipsoid, longName As String, compactName As String, deltaX As Double, deltaY As Double, deltaZ As Double,
                   rotationX As Double, rotationY As Double, rotationZ As Double, scaleDifference As Double,
                   Optional convention As RotationConventions = RotationConventions.PositionVector)
        FullName = longName
        ShortName = compactName
        iSrcEll = sourceEllipsoid
        iTgtEll = targetEllipsoid
        idx = deltaX
        idy = deltaY
        idz = deltaZ
        irx = rotationX
        iry = rotationY
        irz = rotationZ
        ippm = scaleDifference
        iRotConv = convention
    End Sub

    Public Overrides Function ToWGS84(point As Point3D) As Point3D
        Dim TmpResult, TmpGC1, TmpGC2 As New Point3D

        Try
            'Source Ellipsoid costants
            Dim axis, ecc, axis2 As Double
            axis = iSrcEll.SemiMayorAxis
            ecc = iSrcEll.Eccentricity
            axis2 = iSrcEll.SemiMinorAxis

            'From LL to XYZ
            Dim V1 As Double
            V1 = iSrcEll.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            TmpGC1.X = (V1 + point.Z) * Math.Cos(point.Y) * Math.Cos(point.X)
            TmpGC1.Y = (V1 + point.Z) * Math.Cos(point.Y) * Math.Sin(point.X)
            TmpGC1.Z = ((1 - ecc ^ 2) * V1 + point.Z) * Math.Sin(point.Y)

            'From XYZ to XYZ
            Dim M, Xr, Yr, Zr As Double
            Dim r1, r2, r3, r4, r5, r6, r7, r8, r9 As Double
            r1 = Math.Cos(iry) * Math.Cos(irz)
            r2 = Math.Cos(irx) * Math.Sin(irz) + Math.Sin(irx) * Math.Sin(iry) * Math.Cos(irz)
            r3 = Math.Sin(irx) * Math.Sin(irz) - Math.Cos(irx) * Math.Sin(iry) * Math.Cos(irz)
            r4 = -Math.Cos(iry) * Math.Sin(irz)
            r5 = Math.Cos(irx) * Math.Cos(irz) - Math.Sin(irx) * Math.Sin(iry) * Math.Sin(irz)
            r6 = Math.Sin(irx) * Math.Cos(irz) + Math.Cos(irx) * Math.Sin(iry) * Math.Sin(irz)
            r7 = Math.Sin(iry)
            r8 = -Math.Sin(irx) * Math.Cos(iry)
            r9 = Math.Cos(irx) * Math.Cos(iry)

            M = 1 + (ippm * 10 ^ -6)
            Xr = TmpGC1.X
            Yr = TmpGC1.Y
            Zr = TmpGC1.Z
            TmpGC2.X = M * ((r1 * Xr) + (r2 * Yr) + (r3 * Zr)) + idx
            TmpGC2.Y = M * ((r4 * Xr) + (r5 * Yr) + (r6 * Zr)) + idy
            TmpGC2.Z = M * ((r7 * Xr) + (r8 * Yr) + (r9 * Zr)) + idz

            'Target Ellipsoid costants
            axis = iTgtEll.SemiMayorAxis
            ecc = iTgtEll.Eccentricity
            axis2 = iTgtEll.SemiMinorAxis

            'From XYZ to LL
            Dim eps, p, q, V2 As Double
            eps = ecc ^ 2 / (1 - ecc ^ 2)
            p = (TmpGC2.X ^ 2 + TmpGC2.Y ^ 2) ^ 0.5
            q = Math.Atan((TmpGC2.Z * axis) / (p * axis2))

            TmpResult.X = Math.Atan(TmpGC2.Y / TmpGC2.X)
            TmpResult.Y = Math.Atan((TmpGC2.Z + eps * axis2 * Math.Sin(q) ^ 3) / (p - ecc ^ 2 * axis * Math.Cos(q) ^ 3))
            'Calculate the elevetion
            V2 = iTgtEll.GetRadiuosOfCurvatureInThePrimeVertical(TmpResult.Y)
            TmpResult.Z = (p / Math.Cos(TmpResult.Y)) - V2

            Return TmpResult
        Catch ex As Exception
            Return New Point3D
        End Try

    End Function

    Public Overrides Function FromWGS84(point As Point3D) As Point3D
        Dim TmpResult, TmpGC1, TmpGC2 As New Point3D

        Try
            'Target Ellipsoid costants
            Dim axis, ecc, axis2 As Double
            axis = iTgtEll.SemiMayorAxis
            ecc = iTgtEll.Eccentricity
            axis2 = iTgtEll.SemiMinorAxis

            'From LL to XYZ
            Dim V1 As Double
            V1 = iSrcEll.GetRadiuosOfCurvatureInThePrimeVertical(point.Y)
            TmpGC1.X = (V1 + point.Z) * Math.Cos(point.Y) * Math.Cos(point.X)
            TmpGC1.Y = (V1 + point.Z) * Math.Cos(point.Y) * Math.Sin(point.X)
            TmpGC1.Z = ((1 - ecc ^ 2) * V1 + point.Z) * Math.Sin(point.Y)

            'From XYZ to XYZ
            Dim M, Xr, Yr, Zr As Double
            Dim r1, r2, r3, r4, r5, r6, r7, r8, r9 As Double
            r1 = Math.Cos(iry) * Math.Cos(irz)
            r2 = Math.Cos(irx) * Math.Sin(irz) + Math.Sin(irx) * Math.Sin(iry) * Math.Cos(irz)
            r3 = Math.Sin(irx) * Math.Sin(irz) - Math.Cos(irx) * Math.Sin(iry) * Math.Cos(irz)
            r4 = -Math.Cos(iry) * Math.Sin(irz)
            r5 = Math.Cos(irx) * Math.Cos(irz) - Math.Sin(irx) * Math.Sin(iry) * Math.Sin(irz)
            r6 = Math.Sin(irx) * Math.Cos(irz) + Math.Cos(irx) * Math.Sin(iry) * Math.Sin(irz)
            r7 = Math.Sin(iry)
            r8 = -Math.Sin(irx) * Math.Cos(iry)
            r9 = Math.Cos(irx) * Math.Cos(iry)

            M = 1 + (-ippm * 10 ^ -6)
            Xr = TmpGC1.X
            Yr = TmpGC1.Y
            Zr = TmpGC1.Z
            TmpGC2.X = M * ((-r1 * Xr) + (-r2 * Yr) + (-r3 * Zr)) - idx
            TmpGC2.Y = M * ((-r4 * Xr) + (-r5 * Yr) + (-r6 * Zr)) - idy
            TmpGC2.Z = M * ((-r7 * Xr) + (-r8 * Yr) + (-r9 * Zr)) - idz

            'Source Ellipsoid costants
            axis = iSrcEll.SemiMayorAxis
            ecc = iSrcEll.Eccentricity
            axis2 = iSrcEll.SemiMinorAxis

            'From XYZ to LL
            Dim eps, p, q, V2 As Double
            eps = ecc ^ 2 / (1 - ecc ^ 2)
            p = (TmpGC2.X ^ 2 + TmpGC2.Y ^ 2) ^ 0.5
            q = Math.Atan((TmpGC2.Z * axis) / (p * axis2))

            TmpResult.X = Math.Atan(TmpGC2.Y / TmpGC2.X)
            TmpResult.Y = Math.Atan((TmpGC2.Z + eps * axis2 * Math.Sin(q) ^ 3) / (p - ecc ^ 2 * axis * Math.Cos(q) ^ 3))
            'Calculate the elevetion
            V2 = iTgtEll.GetRadiuosOfCurvatureInThePrimeVertical(TmpResult.Y)
            TmpResult.Z = (p / Math.Cos(TmpResult.Y)) - V2

            Return TmpResult
        Catch ex As Exception
            Return New Point3D
        End Try

    End Function

    Public Overrides Function ToXml() As XElement
        'New(Source As Ellipsoid, FullName As String, ShortName As String, DeltaX As Double, DeltaY As Double, DeltaZ As Double)
        Dim x As New XmlTags
        Dim ellx As New XElement(x.Transformation)
        ellx.Add(New XElement(x.Type, Type.ToString))
        ellx.Add(New XElement(x.FullName, FullName))
        ellx.Add(New XElement(x.ShortName, ShortName))
        ellx.Add(New XElement(x.DeltaX, idx))
        ellx.Add(New XElement(x.DeltaY, idy))
        ellx.Add(New XElement(x.DeltaZ, idz))
        ellx.Add(New XElement(x.RotationConvention, iRotConv))
        ellx.Add(New XElement(x.RotationX, irx))
        ellx.Add(New XElement(x.RotationY, iry))
        ellx.Add(New XElement(x.RotationZ, irz))
        ellx.Add(New XElement(x.ScaleDifference, ippm))
        ellx.Add(New XElement(x.SourceEllipsoid, iSrcEll.ToXml(False)))
        ellx.Add(New XElement(x.TargetEllipsoid, iTgtEll.ToXml(False)))

        Return ellx

    End Function

    Public Overrides Function GetParams() As List(Of ParamNameValue)
        Dim tmpList As New List(Of ParamNameValue) From {
            New ParamNameValue("Delta X", idx, ParamType.Generic, True),
            New ParamNameValue("Delta Y", idy, ParamType.Generic, True),
            New ParamNameValue("Delta Z", idz, ParamType.Generic, True),
            New ParamNameValue("Rotation X", irx, ParamType.Angle, True),
            New ParamNameValue("Rotation Y", iry, ParamType.Angle, True),
            New ParamNameValue("Rotation Z", irz, ParamType.Angle, True),
            New ParamNameValue("Rotation convention", iRotConv, ParamType.Generic, True),
            New ParamNameValue("Scale difference", ippm, ParamType.ScaleFactor, True)
        }

        Return tmpList

    End Function

End Class