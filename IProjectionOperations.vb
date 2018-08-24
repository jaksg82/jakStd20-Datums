Public Interface IProjectionOperations

    ''' <summary>
    ''' Convert projected East/North coordinates to Lat/Long coordinates.
    ''' </summary>
    ''' <param name="point">Projected point coordinates to be converted.</param>
    ''' <returns>Point coordinates on the ellipsoid.</returns>
    Function ToGeographic(point As jakStd20_MathExt.Point3D) As jakStd20_MathExt.Point3D

    ''' <summary>
    ''' Convert Lat/Long coordinates to projected East/North coordinates.
    ''' </summary>
    ''' <param name="point">Geographic point coordinates to convert.</param>
    ''' <returns>Projected point coordinates.</returns>
    Function FromGeographic(point As jakStd20_MathExt.Point3D) As jakStd20_MathExt.Point3D

    ''' <summary>
    ''' Projection method used.
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property Type As Projections.Method

    'Function ToWellKnownText() As String

    'Function ToXml() As XElement

End Interface