VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APITimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event IntervalReached()

Private m_blnEnabled As Boolean
Private m_lngInterval As Long
Private m_lngTimerID As Long

Public Sub Enabled(val As Boolean, TimeInterval As Long)
    Dim lngTemp As Long
    If val = False Then
        KillCurrentTimer
    Else
        If TimeInterval > 0 Then
            If TimerID <> 0 Then KillCurrentTimer
            SetCurrentTimer TimeInterval
        Else
            Err.Raise ErrorCodes.InvalidInterval, _
                "APITimer.Enabled", "Interval value must be positive."
        End If
    End If
End Sub

Private Sub SetCurrentTimer(TimeInt As Long)
    Dim lngTemp As Long
    lngTemp = SetTimer(0&, 0&, TimeInt, AddressOf TimerProc)
    If lngTemp <> 0 Then
        gcolTimers.Add Me, "T" & CStr(lngTemp)
        TimerID = lngTemp
        Interval = TimeInt
    Else
        Err.Raise ErrorCodes.SetTimerError, "APITimer.SetCurrentTimer", _
            "Unable to set the timer."
    End If
End Sub
Public Property Get Interval() As Long
    Interval = m_lngInterval
End Property

Private Property Let Interval(val As Long)
    m_lngInterval = val
End Property
Public Property Get TimerID() As Long
    TimerID = m_lngTimerID
End Property
Private Property Let TimerID(val As Long)
    m_lngTimerID = val
End Property
Friend Sub APITimerIntervalReached()
    RaiseEvent IntervalReached
End Sub
Private Sub KillCurrentTimer()
    Dim lngTemp As Long
    If TimerID <> 0 Then
        lngTemp = KillTimer(0&, TimerID)
        If lngTemp <> 0 Then
            gcolTimers.Remove "T" & CStr(TimerID)
            TimerID = 0
        Else
            Err.Raise ErrorCodes.KillTimerError, _
                "APITimer.KillCurrentTimer", "Unable to kill the timer"
        End If
    End If
End Sub
