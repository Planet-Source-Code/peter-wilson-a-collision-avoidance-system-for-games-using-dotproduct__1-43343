Attribute VB_Name = "m2DTransforms"
Option Explicit

Public Function MatrixIdentity() As mdrMATRIX3x3
    
    With MatrixIdentity
    
        .rc11 = 1: .rc12 = 0: .rc13 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1
    
    End With
    
End Function

Public Function MatrixMultiplyVector(MatrixIn As mdrMATRIX3x3, VectorIn As mdrVector3) As mdrVector3
        
    With MatrixMultiplyVector
    
        .x = (MatrixIn.rc11 * VectorIn.x) + (MatrixIn.rc12 * VectorIn.y) + (MatrixIn.rc13 * VectorIn.w)
        .y = (MatrixIn.rc21 * VectorIn.x) + (MatrixIn.rc22 * VectorIn.y) + (MatrixIn.rc23 * VectorIn.w)
        
        ' Although W is included in the above calculations,
        ' we can set w=1 for the resulting vector.
        .w = 1
        
    End With
    
End Function

Public Function Vect3Subtract(V1 As mdrVector3, V2 As mdrVector3) As mdrVector3

    ' Subtracts V2 from V1
    
    Vect3Subtract.x = V1.x - V2.x
    Vect3Subtract.y = V1.y - V2.y
    
    ' We can safely ignore the W component.
    Vect3Subtract.w = 1
    
End Function
Public Function Vect3Addition(V1 As mdrVector3, V2 As mdrVector3) As mdrVector3

    ' Adds V1 and V2 together.
    
    Vect3Addition.x = V1.x + V2.x
    Vect3Addition.y = V1.y + V2.y
    
    ' We can safely ignore the W component.
    Vect3Addition.w = 1
    
End Function

Public Function MatrixMultiply(m1 As mdrMATRIX3x3, m2 As mdrMATRIX3x3) As mdrMATRIX3x3
    
    Dim m1b As mdrMATRIX3x3
    Dim m2b As mdrMATRIX3x3
    m1b = m1 ' This assignment is neccessary because I can't pass a custom data-type
    m2b = m2 ' ByVal. By creating addition copies of my values I can work around this limitation.
    
    MatrixMultiply = MatrixIdentity
    
    With MatrixMultiply
    
        .rc11 = (m1b.rc11 * m2b.rc11) + (m1b.rc21 * m2b.rc12) + (m1b.rc31 * m2b.rc13)
        .rc12 = (m1b.rc12 * m2b.rc11) + (m1b.rc22 * m2b.rc12) + (m1b.rc32 * m2b.rc13)
        .rc13 = (m1b.rc13 * m2b.rc11) + (m1b.rc23 * m2b.rc12) + (m1b.rc33 * m2b.rc13)
        
        .rc21 = (m1b.rc11 * m2b.rc21) + (m1b.rc21 * m2b.rc22) + (m1b.rc31 * m2b.rc23)
        .rc22 = (m1b.rc12 * m2b.rc21) + (m1b.rc22 * m2b.rc22) + (m1b.rc32 * m2b.rc23)
        .rc23 = (m1b.rc13 * m2b.rc21) + (m1b.rc23 * m2b.rc22) + (m1b.rc33 * m2b.rc23)
        
        .rc31 = (m1b.rc11 * m2b.rc31) + (m1b.rc21 * m2b.rc32) + (m1b.rc31 * m2b.rc33)
        .rc32 = (m1b.rc12 * m2b.rc31) + (m1b.rc22 * m2b.rc32) + (m1b.rc32 * m2b.rc33)
        .rc33 = (m1b.rc13 * m2b.rc31) + (m1b.rc23 * m2b.rc32) + (m1b.rc33 * m2b.rc33)
    
    End With
    
End Function

Public Function Vec3Length(V1 As mdrVector3) As Single

    ' Returns the length of a 3-D vector.
    ' The length of a vector is from the origin (0,0) to x,y
    ' We work this out using Pythagoras theorem:  c^2 = a^2 + b^2
    
    Vec3Length = Sqr((V1.x ^ 2) + (V1.y ^ 2))
    
    ' We can safely ignore the W component.
    
End Function
Public Function Vec3MultiplyByScalar(V1 As mdrVector3, Scalar As Single) As mdrVector3
    
    Vec3MultiplyByScalar.x = V1.x * Scalar
    Vec3MultiplyByScalar.y = V1.y * Scalar
    
    ' We can safely ignore the W component.
    Vec3MultiplyByScalar.w = 1
    
End Function
Public Function Vec3Normalize(V1 As mdrVector3) As mdrVector3

    ' Returns the normalized version of a 3D vector.
    '
    ' When you divide a vector by it's own length (from origin 0,0 to x,y)
    ' you'll get a vector who's length is exactly 1.0
    
    Dim sngLength As Single
    
    sngLength = Vec3Length(V1)
    
    If sngLength = 0 Then sngLength = 1
    
    Vec3Normalize.x = V1.x / sngLength
    Vec3Normalize.y = V1.y / sngLength
    
    ' We can safely ignore the W component.
    Vec3Normalize.w = 1
    
End Function
Public Function DotProduct(VectorU As mdrVector3, VectorV As mdrVector3) As Single

    ' Determines the dot-product of two vectors.
    
    DotProduct = (VectorU.x * VectorV.x) + (VectorU.y * VectorV.y)

    ' We can safely ignore the W component.
    
End Function
Public Function MatrixTranslation(OffsetX As Single, OffsetY As Single) As mdrMATRIX3x3
    
    MatrixTranslation = MatrixIdentity
    
    MatrixTranslation.rc13 = OffsetX
    MatrixTranslation.rc23 = OffsetY
    
End Function


Public Function MatrixScaling(ScaleX As Single, ScaleY As Single) As mdrMATRIX3x3
    
    MatrixScaling = MatrixIdentity
    
    MatrixScaling.rc11 = ScaleX
    MatrixScaling.rc22 = ScaleY
    
End Function

Public Function MatrixRotationZ(Radians As Single) As mdrMATRIX3x3

    ' In this VB application:
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points *into* the monitor.
    
    Dim sngCosine As Single
    Dim sngSine As Single
    
    sngCosine = Cos(Radians)
    sngSine = Sin(Radians)
    
    MatrixRotationZ = MatrixIdentity

    ' Positive rotations in a left-handed system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' Z-Axis rotation.
    ' A positive rotation of 90° transforms the X axis into the Y axis
    ' =================================================================
    With MatrixRotationZ
        
        .rc11 = sngCosine
        .rc12 = -sngSine
        .rc21 = sngSine
        .rc22 = sngCosine

    End With
    
End Function


Public Function MatrixViewMapping(Xmin As Single, Xmax As Single, Ymin As Single, Ymax As Single, Umin As Single, Umax As Single, Vmin As Single, Vmax As Single) As mdrMATRIX3x3
        
    Dim matTranslateA As mdrMATRIX3x3
    Dim matScale As mdrMATRIX3x3
    Dim matTranslateB As mdrMATRIX3x3
    Dim sngScaleUX As Single
    Dim sngScaleVY As Single
    
    matTranslateA = MatrixTranslation(-Xmin, -Ymin)
    
    sngScaleUX = (Umax - Umin) / (Xmax - Xmin)
    sngScaleVY = (Vmax - Vmin) / (Ymax - Ymin)
    
    matScale = MatrixScaling(sngScaleUX, sngScaleVY)
    
    matTranslateB = MatrixTranslation(Umin, Vmin)
    
    MatrixViewMapping = MatrixIdentity
    MatrixViewMapping = MatrixMultiply(MatrixViewMapping, matTranslateA)
    MatrixViewMapping = MatrixMultiply(MatrixViewMapping, matScale)
    MatrixViewMapping = MatrixMultiply(MatrixViewMapping, matTranslateB)
    
End Function

