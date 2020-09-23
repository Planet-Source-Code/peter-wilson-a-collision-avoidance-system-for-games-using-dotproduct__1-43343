Attribute VB_Name = "mDataStructures"
Option Explicit


Public Type mdrVector3
    x As Single
    y As Single
    w As Single
End Type


Public Type mdr2DObject
    Caption As String   ' (Optional)
    
    Enabled As Boolean  ' Normallt TRUE, if FALSE then no calculations take place.
    Visible As Boolean  ' Normally TRUE. ie. Ships can still be included in calculations even if they are invisible.
    ParticleLifeRemaining As Single ' A Particle Object is only Enabled & Visible for a short time.
    
    ' Geometery
    Vertex() As mdrVector3  ' Original Vertices
    TVertex() As mdrVector3 ' Transformed Vertices
    Face() As Variant   ' Array of indicies. ie. Face(0) = Array(0,1,2,3... n-1, n)
    
    WorldPos As mdrVector3 ' Position of the object in World Coordinate system.
    
    Vector As mdrVector3    ' Direction/Speed Vector.
    TVector As mdrVector3   ' Transformed Vector
    
    SpinVector As Single ' Usually between -4 and 4
    RotationAboutZ As Single ' Rotation in degrees.
    
    Red As Integer ' 0-255
    Green As Integer ' 0-255
    Blue As Integer ' 0-255
End Type


Public Type mdrMATRIX3x3
    rc11 As Single: rc12 As Single: rc13 As Single
    rc21 As Single: rc22 As Single: rc23 As Single
    rc31 As Single: rc32 As Single: rc33 As Single
End Type


