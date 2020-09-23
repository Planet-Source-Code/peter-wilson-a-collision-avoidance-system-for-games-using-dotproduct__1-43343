Attribute VB_Name = "mGameLoop"
Option Explicit

Public g_strGameState As String
Private m_intLevel As Integer


Public m_Enemies() As mdr2DObject

Private m_MaxAsteroids As Integer
Private m_Asteroids() As mdr2DObject

' World Window Limits ie. This is the game's world coordinates (which could be very large)
Private m_Xmin As Single
Private m_Xmax As Single
Private m_Ymin As Single
Private m_Ymax As Single

' ViewPort Limits ie. Usually the limits of a VB form, or picturebox (which could be very small)
Private m_Umin As Single
Private m_Umax As Single
Private m_Vmin As Single
Private m_Vmax As Single

' Module Level Matrices (that don't change much)
Private m_matScale As mdrMATRIX3x3
Private m_matViewMapping As mdrMATRIX3x3

Public Function Create_Asteroids(ByVal Qty As Integer, MinSize As Integer, MaxSize As Integer, WorldX As Single, WorldY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single) As Integer

    ' "Attempts to create" the specified number of Asteroids,
    ' and returns the number of Asteroids "actually created".
    Create_Asteroids = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Asteroids(intN).Enabled = False Then
            ' This Asteroid is no longer used, so we can use this one --> m_Asteroids(intN)
            
            If Qty > 0 Then
                
                ' Create a random sized asteroid within the min/max parameters specified.
                sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                m_Asteroids(intN) = CreateRandomShapeAsteroid(sngRadius)
                
                ' Fill-in some properties.
                With m_Asteroids(intN)
                    .Enabled = True
                    .Visible = True
                    .Caption = "Asteroid"
                    .ParticleLifeRemaining = LifeTime
                    
                    .WorldPos.x = WorldX
                    .WorldPos.y = WorldY
                    .WorldPos.w = 1
                    
                    ' Initial Vector
                    .Vector.x = GetRNDNumberBetween(-500, 500)
                    .Vector.y = GetRNDNumberBetween(-500, 500)
                    .Vector.w = 1
                    
                    .SpinVector = GetRNDNumberBetween(-5, 5)
                    .RotationAboutZ = 0
                    
                    .Red = Red: .Green = Green: .Blue = Blue
                    
                End With
                
                Qty = Qty - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxAsteroids) Or (Qty = 0)
        
End Function

Public Sub Main()

    ' ==========================================================================
    ' This routine get's called by a Timer Event regardless of what's happening.
    ' (Although you can have multiple Timer Controls, it tends to make programs
    '  disorganised and less predictable. By using only a single Timer control,
    '  I have very strict control over what occurs and when. This routine is
    '  actually a mini-"state machine"... well actually most computer programs
    '  are, but I digress... look them up, learn them, they are cool.)
    '  Also See: www.midar.com.au/vblessons/
    ' ==========================================================================
    
    Call ProcessKeyboardInput
    
    Select Case g_strGameState
        Case ""
            ' =======================
            ' Initialize game/program
            ' =======================
            Randomize
            m_intLevel = 0
            frmCanvas.Show
            g_strGameState = "LevelComplete"
            
        Case "PlayingLevel"
            Call PlayGame
            
        Case "LevelComplete"
            m_intLevel = m_intLevel + 1
            Call LoadLevel(m_intLevel)
            g_strGameState = "PlayingLevel"
                        
        Case "Quit"
            frmCanvas.Timer_DoAnimation.Enabled = False
            Unload frmCanvas
            
    End Select
    
End Sub
Private Sub LoadLevel(ByVal Level As Integer)

    Dim intN As Integer
    Dim sngRadius As Single
        
    ' ================
    ' Create Asteroids
    ' ================
    ' One Large Asteroid can be split into two medium asteroids,
    ' and then each of these medium ones, can be split again into smaller ones.
    m_MaxAsteroids = Level * 4
    If m_MaxAsteroids <> 0 Then
        ReDim m_Asteroids(m_MaxAsteroids - 1)
        Call Create_Asteroids(Level, 1000, 1000, 0, 0, 0, 192, 192, 0)
    End If
    
    
    ' =====================================================================
    ' Create Enemies
    ' This should be space ships, but I've just made them asteroids for now
    ' =====================================================================
    ReDim m_Enemies(Int(Level / 2))
    For intN = 0 To Int(Level / 2)
        sngRadius = GetRNDNumberBetween(200, 800)
        m_Enemies(intN) = CreateRandomShapeAsteroid(sngRadius)
        
        With m_Enemies(intN)
            .Caption = "Enemy" & intN
            .Enabled = True
            .WorldPos.x = GetRNDNumberBetween(-20000, 20000)
            .WorldPos.y = GetRNDNumberBetween(-20000, 20000)
            .WorldPos.w = 1
            .Vector.x = 0
            .Vector.y = 0
            .SpinVector = GetRNDNumberBetween(-5, 5)
            .Red = 255: .Green = 0: .Blue = 0
        End With
    Next intN
    
End Sub

Public Sub PlayGame()

    ' Set global scale factor.
    m_matScale = MatrixScaling(5, 5)
        
    ' Place asteroid objects in the world, rotate, scale and map them correctly.
    Call Calculate_Asteroids
        
    ' Place enemy objects in the world, rotate, scale and map them correctly.
    Call Calculate_Enemies_Part1
    
    ' Draw everything.
    Call Refresh_GameScreen
    
    ' Calculate enemy danger level (and draw results)
    ' This is the part that uses the DotProduct.
    Call Calculate_EnemyAI
    
    
    
    ' =============================================
    ' Just some quickhelp code (delete if you want)
    ' =============================================
    Static s_lngCounter As Long
    s_lngCounter = s_lngCounter + 1
    If s_lngCounter = (2 ^ 31) - 1 Then s_lngCounter = 0
    Select Case s_lngCounter
        Case 100
            MsgBox "Use the arrow keys to move one of the little red things around", vbInformation
        Case 300
            MsgBox "Press the Space Bar to change levels", vbInformation
    End Select
    ' ==========================================
    ' End of QuickHelp Code (delete if you want)
    ' ==========================================
    
End Sub

Public Sub Calculate_Asteroids()

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    On Error GoTo errTrap
    
    
    For intN = LBound(m_Asteroids) To UBound(m_Asteroids)
        With m_Asteroids(intN)
            If .Enabled = True Then
            
                ' Translate
                ' =========
                .WorldPos.x = .WorldPos.x + .Vector.x
                .WorldPos.y = .WorldPos.y + .Vector.y
                
                If .WorldPos.x > m_Xmax Then .WorldPos.x = .WorldPos.x - (m_Xmax - m_Xmin)
                If .WorldPos.x < m_Xmin Then .WorldPos.x = .WorldPos.x + (m_Xmax - m_Xmin)
                If .WorldPos.y > m_Ymax Then .WorldPos.y = .WorldPos.y - (m_Ymax - m_Ymin)
                If .WorldPos.y < m_Ymin Then .WorldPos.y = .WorldPos.y + (m_Ymax - m_Ymin)
                
                matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
                
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                
                ' Multiply matrices in the correct order.
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, m_matScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, m_matViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
                ' Conditionally Compiled (see Project Properties)
                #If gcShowVectors = -1 Then
                    ' Transform the Direction/Speed Vector to screen space
                    ' Do this step, only if you wish to display this vector on the screen.
                    ' Displaying the vector on screen, is only useful for debugging/instructional purposes.
                    ' Remember, DO NOT rotate the Direction/Speed vector (try it, and see what happens)
                    matResult = MatrixIdentity
                    matResult = MatrixMultiply(matResult, m_matScale)
                    matResult = MatrixMultiply(matResult, matTranslate)
                    matResult = MatrixMultiply(matResult, m_matViewMapping)
                    .TVector = MatrixMultiplyVector(matResult, .Vector)
                #End If
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub
Public Sub Calculate_Enemies_Part1()

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    For intN = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intN)
                        
            ' Translate
            ' =========
            .WorldPos.x = .WorldPos.x + .Vector.x
            .WorldPos.y = .WorldPos.y + .Vector.y
            
            If .WorldPos.x > m_Xmax Then .WorldPos.x = .WorldPos.x - (m_Xmax - m_Xmin)
            If .WorldPos.x < m_Xmin Then .WorldPos.x = .WorldPos.x + (m_Xmax - m_Xmin)
            If .WorldPos.y > m_Ymax Then .WorldPos.y = .WorldPos.y - (m_Ymax - m_Ymin)
            If .WorldPos.y < m_Ymin Then .WorldPos.y = .WorldPos.y + (m_Ymax - m_Ymin)
            
            matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
            
            .RotationAboutZ = .RotationAboutZ + .SpinVector
            matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))

            
            ' Multiply matrices in the correct order.
            matResult = MatrixIdentity
            matResult = MatrixMultiply(matResult, m_matScale)
            matResult = MatrixMultiply(matResult, matRotationAboutZ)
            matResult = MatrixMultiply(matResult, matTranslate)
            matResult = MatrixMultiply(matResult, m_matViewMapping)
            
            For intJ = LBound(.Vertex) To UBound(.Vertex)
                .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
            Next intJ
            
        End With
    Next intN

End Sub

Public Sub Calculate_EnemyAI()

    Dim intEnemy As Integer
    Dim intAsteroid As Integer
    Dim VectDifference As mdrVector3
    Dim VectU As mdrVector3
    Dim VectV As mdrVector3
    Dim sngDotProduct As Single
    
    ' Debug Display Graphics Only
    #If gcShowVectors = -1 Then
        Dim vectAsteroid As mdrVector3
        Dim vectDisplay As mdrVector3
        Dim strMsg As String
    #End If
    
    On Error GoTo errTrap
    
    ' Loop through all active Enemy objects ...
    For intEnemy = LBound(m_Enemies) To UBound(m_Enemies)
        With m_Enemies(intEnemy)
            If .Enabled = True Then
                
                ' ...looking for Asteroids that are enabled.
                For intAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                    If m_Asteroids(intAsteroid).Enabled = True Then
                    
                        ' Calculate the X,Y difference between the an Enemy and an Asteroid.
                        VectDifference = Vect3Subtract(m_Asteroids(intAsteroid).WorldPos, .WorldPos)
                        
                        ' Normalize the vectors.
                        VectU = Vec3Normalize(m_Asteroids(intAsteroid).Vector)
                        VectV = Vec3Normalize(VectDifference)
                        
                        ' Get the DotProduct between the two vectors.
                        ' This tells us the cosine of the angle between them.
                        sngDotProduct = DotProduct(VectU, VectV)
                        
                        #If gcShowVectors = -1 Then
                            strMsg = ""
                            
                            ' Create some end-points that we can draw too.
                            vectDisplay = Vec3Normalize(VectDifference)
                            vectDisplay = Vec3MultiplyByScalar(vectDisplay, 5000) ' <<< Try changing this 5000 amount to see what happens.
                            vectDisplay = Vect3Addition(vectDisplay, .WorldPos)
                            vectDisplay = MatrixMultiplyVector(m_matViewMapping, vectDisplay)
                            
                            vectAsteroid = m_Asteroids(intAsteroid).WorldPos
                            vectAsteroid = MatrixMultiplyVector(m_matViewMapping, vectAsteroid)
                            
                            frmCanvas.FontSize = 7
                            frmCanvas.DrawWidth = 1
                            frmCanvas.DrawStyle = vbDot
                            frmCanvas.Font = "Small Fonts"
                            
                            strMsg = ""
                            If sngDotProduct > 0.7 Then
                                strMsg = "(very safe - moving away)"
                                frmCanvas.ForeColor = RGB(0, 255, 0)
                            ElseIf sngDotProduct > 0 Then
                                strMsg = "(safe - moving away)"
                                frmCanvas.ForeColor = RGB(255, 255, 0)
                            ElseIf sngDotProduct < -0.98 Then
                                strMsg = "(Extreme Danger!!!)"
                                frmCanvas.DrawStyle = vbSolid
                                frmCanvas.DrawWidth = 3
                                frmCanvas.ForeColor = RGB(255, 0, 0)
                            ElseIf sngDotProduct < -0.95 Then
                                strMsg = "(Danger)"
                                frmCanvas.DrawStyle = vbSolid
                                frmCanvas.DrawWidth = 2
                                frmCanvas.ForeColor = RGB(255, 0, 0)
                            ElseIf sngDotProduct < -0.9 Then
                                strMsg = "(threat)"
                                frmCanvas.DrawStyle = vbSolid
                                frmCanvas.ForeColor = RGB(255, 64, 0)
                            ElseIf sngDotProduct < -0.8 Then
                                strMsg = "(possible threat)"
                                frmCanvas.ForeColor = RGB(255, 127, 0)
                            ElseIf sngDotProduct < 0 Then
                                strMsg = "(avoid if required)"
                                frmCanvas.ForeColor = RGB(255, 255, 0)
                            End If
                            
                            frmCanvas.Line (vectAsteroid.x, vectAsteroid.y)-(vectDisplay.x, vectDisplay.y)
                            frmCanvas.Print Format(sngDotProduct, "0.00") & " ";
                            frmCanvas.Print strMsg
                            
                        #End If
                    End If ' Is Asteroid Enabled?
                Next intAsteroid
            End If ' Is Enemy Enabled?
        End With
    Next intEnemy
    
    Exit Sub
errTrap:
    
End Sub

Public Sub Draw_Faces(CurrentObject() As mdr2DObject)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    frmCanvas.DrawStyle = vbSolid
    frmCanvas.DrawMode = vbCopyPen
    frmCanvas.DrawWidth = 2
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                frmCanvas.ForeColor = RGB(.Red, .Green, .Blue)
                
                For intFaceIndex = LBound(.Face) To UBound(.Face)
                    
                    For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                    
                        intVertexIndex = .Face(intFaceIndex)(intK)
                        xPos = .TVertex(intVertexIndex).x
                        yPos = .TVertex(intVertexIndex).y
                        
                        If LBound(.Face(intFaceIndex)) = UBound(.Face(intFaceIndex)) Then
                            If .Caption = "Asteroid" Then
                                #If gcShowVectors = -1 Then
                                    frmCanvas.DrawWidth = 3
                                    frmCanvas.ForeColor = RGB(255, 255, 0)
                                    frmCanvas.PSet (xPos, yPos)
                                    frmCanvas.DrawWidth = 1
                                    frmCanvas.ForeColor = RGB(.Red, .Green, .Blue)
                                    frmCanvas.Line (xPos, yPos)-(.TVector.x, .TVector.y)
                                #End If
                            End If
                        Else
                        
                            ' Normal Face; move to first point, then draw to the others.
                            ' ==========================================================
                            If intK = LBound(.Face(intFaceIndex)) Then
                                ' Move to first point
                                frmCanvas.Line (xPos, yPos)-(xPos, yPos)
                                
                            Else
                                ' Draw to point
                                frmCanvas.Line -(xPos, yPos)
                            End If
                            
                        End If
                        
                    Next intK
                Next intFaceIndex
                
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub
Public Sub Init_ViewMapping()

    ' Set the size of the World's window.
    m_Xmin = -32768
    m_Xmax = 32767
    
    m_Ymin = -32768
    m_Ymax = 32767
    
    ' Invert Y so that the following is true.
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points *into* the monitor

    
    ' Set the size of the ViewPort's windows.
    m_Umin = 0
    m_Umax = frmCanvas.Width
    m_Vmin = frmCanvas.Height
    m_Vmax = 0
        
    m_matViewMapping = MatrixViewMapping(m_Xmin, m_Xmax, m_Ymin, m_Ymax, m_Umin, m_Umax, m_Vmin, m_Vmax)

End Sub


Private Sub Refresh_GameScreen()

    frmCanvas.Cls
    Call Draw_Faces(m_Asteroids)
    Call Draw_Faces(m_Enemies)
    
End Sub

