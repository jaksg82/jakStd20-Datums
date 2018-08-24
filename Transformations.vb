Public MustInherit Class Transformations
    Private iFname, iSname As String
    Private iSrcEll, iTgtEll As Ellipsoid

    ''' <summary>
    ''' The reference ellipsoid for the transformation
    ''' </summary>
    ''' <returns>Return an Ellipsoid object.</returns>
    Property SourceEllipsoid As Ellipsoid
        Get
            Return iSrcEll
        End Get
        Set(value As Ellipsoid)
            iSrcEll = value
        End Set
    End Property

    ''' <summary>
    ''' The reference ellipsoid for the transformation
    ''' </summary>
    ''' <returns>Return an Ellipsoid object.</returns>
    Property TargetEllipsoid As Ellipsoid
        Get
            Return iTgtEll
        End Get
        Set(value As Ellipsoid)
            iTgtEll = value
        End Set
    End Property

    ''' <summary>
    ''' Name of the transformation operation
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

    Overridable ReadOnly Property Type As Methods
        Get
            Return Methods.None
        End Get
    End Property

    Public MustOverride Property RotationConvention As RotationConventions

    Public Enum Methods
        None
        Geocentric3Parameter
        Helmert7Parameter
        MolodenskyBadekas10Parameter
        AbridgedMolodensky
    End Enum

    Public Enum RotationConventions
        PositionVector
        CoordinateFrame
    End Enum

    Public Function GetAllMethodNames() As List(Of String)
        Dim tmpList As New List(Of String)

        tmpList.Add("No transformation")
        tmpList.Add("3-parameter geocentric translation")
        tmpList.Add("Helmert 7-parameter transformation")
        tmpList.Add("Molodensky-Badekas 10-parameter transformation")
        tmpList.Add("Abridged Molodensky transformation")

        Return tmpList

    End Function

    Public Shared Function FromXml(value As XElement) As Transformations
        Dim x As New XmlTags

        If value.Name = x.Transformation Then
            Dim xfname, xsname, xtype, xRot As String
            Dim xdx, xdy, xdz, xrx, xry, xrz, xpx, xpy, xpz, xppm As Double
            Dim xSrc, xTgt As XElement
            Dim eSrc, eTgt As Ellipsoid
            Dim eRot As RotationConventions

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
                xtype = Methods.Geocentric3Parameter.ToString
            End Try
            Try
                xRot = value.Element(x.RotationConvention).Value
            Catch ex As Exception
                xRot = RotationConventions.PositionVector.ToString
            End Try
            Try
                xdx = Double.Parse(value.Element(x.DeltaX).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xdx = 0.0
            End Try
            Try
                xdy = Double.Parse(value.Element(x.DeltaY).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xdy = 0.0
            End Try
            Try
                xdz = Double.Parse(value.Element(x.DeltaZ).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xdz = 0.0
            End Try
            Try
                xSrc = value.Element(x.SourceEllipsoid)
                eSrc = Ellipsoid.FromXml(xSrc)
            Catch ex As Exception
                eSrc = New Ellipsoid
            End Try
            Try
                xTgt = value.Element(x.TargetEllipsoid)
                eTgt = Ellipsoid.FromXml(xTgt)
            Catch ex As Exception
                eTgt = New Ellipsoid
            End Try

            'Extra values for 7 parameters
            Try
                xrx = Double.Parse(value.Element(x.RotationX).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xrx = 0.0
            End Try
            Try
                xry = Double.Parse(value.Element(x.RotationY).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xry = 0.0
            End Try
            Try
                xrz = Double.Parse(value.Element(x.RotationZ).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xrz = 0.0
            End Try
            Try
                xppm = Double.Parse(value.Element(x.ScaleDifference).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xppm = 0.0
            End Try

            'Extra values for 10 parameters
            Try
                xpx = Double.Parse(value.Element(x.PivotX).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xpx = 0.0
            End Try
            Try
                xpy = Double.Parse(value.Element(x.PivotY).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xpy = 0.0
            End Try
            Try
                xpz = Double.Parse(value.Element(x.PivotZ).Value, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                xpz = 0.0
            End Try

            Select Case xRot
                Case RotationConventions.PositionVector.ToString
                    eRot = RotationConventions.PositionVector
                Case RotationConventions.CoordinateFrame.ToString
                    eRot = RotationConventions.CoordinateFrame
                Case Else
                    eRot = RotationConventions.PositionVector
            End Select

            Select Case xtype
                Case Methods.AbridgedMolodensky.ToString
                    Return New AbridgedMolodensky(eSrc, eTgt, xfname, xsname, xdx, xdy, xdz)

                Case Methods.Geocentric3Parameter.ToString
                    Return New Geocentric3p(eSrc, eTgt, xfname, xsname, xdx, xdy, xdz)

                Case Methods.Helmert7Parameter.ToString
                    Return New Helmert7p(eSrc, eTgt, xfname, xsname, xdx, xdy, xdz, xrx, xry, xrz, xppm, eRot)

                Case Methods.MolodenskyBadekas10Parameter.ToString
                    Return New MolodenskyBadekas10p(eSrc, eTgt, xfname, xsname, xdx, xdy, xdz, xrx, xry, xrz, xppm, xpx, xpy, xpz, eRot)

                Case Else
                    Return New Geocentric3p()

            End Select
        Else
            Return New Geocentric3p()

        End If

    End Function

    Public Shared Function FromParams(method As Methods, sourceEllipsoid As Ellipsoid, longName As String, compactName As String, params As List(Of ParamNameValue)) As Transformations
        Dim eTgt As New Ellipsoid
        For p = 0 To 10
            params.Add(New ParamNameValue("Fill", 0.0, ParamType.Generic, True))
        Next

        Select Case method
            Case Methods.AbridgedMolodensky
                Return New AbridgedMolodensky(sourceEllipsoid, eTgt, longName, compactName, params(0).Value, params(1).Value, params(2).Value)

            Case Methods.Geocentric3Parameter
                Return New Geocentric3p(sourceEllipsoid, eTgt, longName, compactName, params(0).Value, params(1).Value, params(2).Value)

            Case Methods.Helmert7Parameter
                Return New Helmert7p(sourceEllipsoid, eTgt, longName, compactName, params(0).Value, params(1).Value, params(2).Value,
                                           params(3).Value, params(4).Value, params(5).Value, params(7).Value, CType(params(6).Value, RotationConventions))

            Case Methods.MolodenskyBadekas10Parameter
                Return New MolodenskyBadekas10p(sourceEllipsoid, eTgt, longName, compactName, params(0).Value, params(1).Value, params(2).Value,
                                                      params(3).Value, params(4).Value, params(5).Value, params(7).Value,
                                                      params(8).Value, params(9).Value, params(10).Value, CType(params(6).Value, RotationConventions))

            Case Else
                Return New Geocentric3p()

        End Select

    End Function

    Public MustOverride Function ToXml() As XElement

    Public MustOverride Function GetParams() As List(Of ParamNameValue)

    Public MustOverride Function ToWGS84(point As jakStd20_MathExt.Point3D) As jakStd20_MathExt.Point3D

    Public MustOverride Function FromWGS84(point As jakStd20_MathExt.Point3D) As jakStd20_MathExt.Point3D

End Class