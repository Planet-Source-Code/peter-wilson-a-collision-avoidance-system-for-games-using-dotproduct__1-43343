Attribute VB_Name = "mKeyboard"
Option Explicit

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub ProcessKeyboardInput()

    Static s_strPreviousValue As String
    Static s_blnKeyDeBounce As Boolean
    
    Dim lngKeyState As Long
    
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then g_strGameState = "Quit"
    
    lngKeyState = GetKeyState(vbKeyLeft)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldPos.x = m_Enemies(0).WorldPos.x - 1000
    
    lngKeyState = GetKeyState(vbKeyRight)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldPos.x = m_Enemies(0).WorldPos.x + 1000
    
    lngKeyState = GetKeyState(vbKeyUp)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldPos.y = m_Enemies(0).WorldPos.y + 1000
    
    lngKeyState = GetKeyState(vbKeyDown)
    If (lngKeyState And &H8000) Then m_Enemies(0).WorldPos.y = m_Enemies(0).WorldPos.y - 1000
    
    
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        If s_blnKeyDeBounce = False Then
            s_blnKeyDeBounce = True
            g_strGameState = "LevelComplete"
        End If
    Else
        s_blnKeyDeBounce = False
    End If
    
End Sub

