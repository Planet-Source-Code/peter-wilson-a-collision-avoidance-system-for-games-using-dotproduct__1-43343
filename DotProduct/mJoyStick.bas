Attribute VB_Name = "mJoyStick"
Option Explicit

' Flags
Private Const JOYSTICKID1 = 0
Private Const JOYSTICKID2 = 1
Private Const JOY_RETURNBUTTONS = &H80&
Private Const JOY_RETURNCENTERED = &H400&
Private Const JOY_RETURNPOV = &H40&
Private Const JOY_RETURNR = &H8&
Private Const JOY_RETURNU = &H10
Private Const JOY_RETURNV = &H20
Private Const JOY_RETURNX = &H1&
Private Const JOY_RETURNY = &H2&
Private Const JOY_RETURNZ = &H4&
Private Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)

' The JOYINFOEX structure contains extended information about the joystick position, point-of-view position, and button state.
Private Type JOYINFOEX
   dwSize As Long
   dwFlags As Long
   dwXpos As Long
   dwYpos As Long
   dwZpos As Long
   dwRpos As Long
   dwUpos As Long
   dwVpos As Long
   dwButtons As Long
   dwButtonNumber As Long
   dwPOV As Long
   dwReserved1 As Long
   dwReserved2 As Long
End Type

' Return-Values for joyGetPosEx
Private Const JOYERR_NOERROR = 0
Private Const MMSYSERR_NODRIVER = 6      ' No device driver present.
Private Const MMSYSERR_INVALPARAM = 11   ' Invalid parameter passed.
Private Const MMSYSERR_BADDEVICEID = 2   ' Device ID out of range.
Private Const JOYERR_UNPLUGGED = 167     ' Joystick is unplugged.

' ===========
' joyGetPosEx
' ===========
' This function provides access to extended devices such as rudder pedals, point-of-view hats, devices with a large number of buttons, and coordinate systems using up to six axes.
' (For joystick devices that use three axes or fewer and have fewer than four buttons, use the joyGetPos function.)
'
'   m_hJoyStickID   : Identifier of the joystick (JOYSTICKID1 or JOYSTICKID2) to be queried.
'   m_hJoyInfoEx    : Pointer to a JOYINFOEX structure that contains extended position information and button status of the joystick. You must set the dwSize and dwFlags members or joyGetPosEx will fail. The information returned from joyGetPosEx depends on the flags you specify in dwFlags.
Private Declare Function joyGetPosEx Lib "winmm.dll" (ByVal m_hJoyStickID As Long, m_hJoyInfoEx As JOYINFOEX) As Long


Private m_hJoyStickID As Long   ' Set to 0(zero) for Default Joystick.
Private m_hJoyInfoEx As JOYINFOEX

Public Sub GetJoyStickResults(Roll As Single, Pitch As Single, Yaw As Single, Throttle As Single, Buttons As Long)

    Dim lngReturnValue As Long
    
    ' Set to 0(zero) for Default Joystick.
    m_hJoyStickID = 0
        
    ' Reset Pointer
    m_hJoyInfoEx.dwFlags = JOY_RETURNALL
    m_hJoyInfoEx.dwSize = Len(m_hJoyInfoEx)
    
    ' Fetch JoyStick Data via API.
    lngReturnValue = joyGetPosEx(m_hJoyStickID, m_hJoyInfoEx)
    
    ' Returns JOYERR_NOERROR if successful or one of the following error values:
    Select Case lngReturnValue
        Case JOYERR_NOERROR
            ' No Error - everything ok.
            Roll = CSng(m_hJoyInfoEx.dwXpos) - 32768
            Pitch = CSng(m_hJoyInfoEx.dwYpos) - 32768
            Yaw = CSng(m_hJoyInfoEx.dwRpos) - 32768
            Throttle = CSng(m_hJoyInfoEx.dwZpos) - 32768
            Buttons = m_hJoyInfoEx.dwButtons
            
        Case MMSYSERR_NODRIVER
            ' No device driver present.
        
        Case MMSYSERR_INVALPARAM
            ' Invalid parameter passed.
        
        Case MMSYSERR_BADDEVICEID
            ' Device ID out of range.
            
        Case JOYERR_UNPLUGGED
            ' Joystick is unplugged.
            MsgBox "Please plug in your JoyStick and get a firm grip!", vbExclamation, "JoyStick Unplugged"
            
    End Select
    
    
End Sub

