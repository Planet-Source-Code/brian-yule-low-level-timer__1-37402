Attribute VB_Name = "Module1"
Public Declare Function SetTimer Lib "User32.dll" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long, ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "User32.dll" (ByVal hWnd As Long, _
    ByVal nIDEEvent As Long) As Long
    
Public gcolTimers As New Collection

Public Enum ErrorCodes
    UnanticipatedProgramError = 26000
    InvalidInterval
    SetTimerError
    NoCurrentTimer
    KillTimerError
End Enum

Public Function TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal idEvent As Long, ByVal dwTime As Long) As Long
    On Error GoTo Error_TimerProc
    Dim m_objTimer As APITimer
    Set m_objTimer = gcolTimers.Item("T" & CStr(idEvent))
    If Not m_objTimer Is Nothing Then m_objTimer.APITimerIntervalReached
Exit Function
Error_TimerProc:
    If Not m_objTimer Is Nothing Then Set m_objTimer = Nothing
End Function

