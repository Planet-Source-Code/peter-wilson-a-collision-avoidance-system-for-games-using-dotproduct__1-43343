Attribute VB_Name = "mInternalData"
Option Explicit


Public Function CreateRandomShapeAsteroid(Radius As Single) As mdr2DObject

    ' Draws a deformed circle, by adjusting the radius at random intervals around the circumference.
    
    Dim sngAngle As Single
    Dim sngAngleIncrement As Single
    Dim sngMaxRadiusVariation As Single
    Dim sngNewRadius As Single
    Dim intMinSegmentAngle As Integer
    Dim intMaxSegmentAngle As Integer
    Dim sngRadiusVariation As Single
    Dim sngWorldX As Single
    Dim sngWorldY As Single
    Dim sngRadians As Single
    Dim intVertexCount As Integer
    Dim intFaceCount As Integer
    Dim intN As Integer
    
    Dim varVertices As Variant
    
    With CreateRandomShapeAsteroid
    
        ' Set these Min/Max properties to make a random looking asteroid.
        ' Basically, this is a deformed circle.
        ' ===============================================================
        sngMaxRadiusVariation = Radius * 0.2 ' ie. 20% of Radius
        intMinSegmentAngle = 5
        intMaxSegmentAngle = 45
        
        ReDim .Vertex(0)
        intVertexCount = -1
        sngAngle = 0
        Do
        
            sngNewRadius = GetRNDNumberBetween(Radius - sngMaxRadiusVariation, Radius + sngMaxRadiusVariation)
            
            sngRadians = ConvertDeg2Rad(sngAngle)
            sngWorldX = Cos(sngRadians) * sngNewRadius
            sngWorldY = Sin(sngRadians) * sngNewRadius
            
            ' Create new Vertex
            intVertexCount = intVertexCount + 1
            ReDim Preserve .Vertex(intVertexCount)
            .Vertex(intVertexCount).x = sngWorldX
            .Vertex(intVertexCount).y = sngWorldY
            .Vertex(intVertexCount).w = 1
            
            sngAngleIncrement = GetRNDNumberBetween(intMinSegmentAngle, intMaxSegmentAngle)
            sngAngle = sngAngle + sngAngleIncrement
        
        Loop Until sngAngle >= 360
        ReDim .TVertex(intVertexCount)
        
        ' Create the Asteroid's edges (ie. it's outer perimeter)
        ' ie. Face(0) = Array(0,1,2,...,n-1,n)
        ' =====================================================
        ReDim varVertices(intVertexCount + 1)
        ReDim .Face(0)
        For intN = 0 To intVertexCount
            varVertices(intN) = intN
        Next intN
        varVertices(intN) = 0
        .Face(0) = varVertices
        
        ' Create a Single Dot in the middle of the Asteroid and also create a face for it
        ' having only a single vertex.  This isn't really a face, more of a place-holder so
        ' I don't have to re-write my drawing routine.
        ' =================================================================================
        intVertexCount = UBound(.Vertex)
        ReDim Preserve .Vertex(intVertexCount + 1)
        ReDim Preserve .TVertex(intVertexCount + 1)
        .Vertex(intVertexCount + 1).x = 0
        .Vertex(intVertexCount + 1).y = 0
        .Vertex(intVertexCount + 1).w = 1
        
        intFaceCount = UBound(.Face)
        ReDim Preserve .Face(intFaceCount + 1)
        .Face(intFaceCount + 1) = Array(intVertexCount + 1)
    
    End With

End Function

