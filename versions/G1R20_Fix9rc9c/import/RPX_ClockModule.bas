Attribute VB_Name = "RPX_ClockModule"
Option Explicit
' One seconds clock for every elapsed measurement and deadline.
' Timer fallback needs at least one observation per day to count every wrap.
#If VBA7 Then
#If Win64 Then
Private Declare PtrSafe Function RPX_GetTickCount64 Lib "kernel32" Alias "GetTickCount64" () As LongLong
#End If
#End If
Private clockShift As Double, clockTimerDays As Double, clockTimerLast As Double
Private clockTimerReady As Boolean
Private clockTestWrapAt As Double, clockTestRawShift As Double, clockTestStarted As Double
Private clockTestWraps As Long, clockTestRawJump As Double, clockTestCorrectedStep As Double

Public Function RPX_Clock() As Double
    Dim rawSeconds As Double
#If VBA7 Then
#If Win64 Then
    If clockTestWrapAt <= 0# Then
        RPX_Clock = CDbl(RPX_GetTickCount64()) / 1000# + clockShift
        Exit Function
    End If
#End If
#End If
    rawSeconds = CDbl(Timer)
    If clockTestWrapAt > 0# Then
        ' Rotate seconds-of-day so its next midnight is wrap_at seconds away.
        ' This forces the same fallback correction used for a real midnight.
        rawSeconds = rawSeconds + clockTestRawShift
        rawSeconds = rawSeconds - 86400# * Int(rawSeconds / 86400#)
    End If
    If clockTimerReady Then
        If rawSeconds < clockTimerLast Then
            clockTimerDays = clockTimerDays + 86400#
            If clockTestWrapAt > 0# Then
                clockTestWraps = clockTestWraps + 1
                clockTestRawJump = rawSeconds - clockTimerLast
                clockTestCorrectedStep = rawSeconds + 86400# - clockTimerLast
            End If
        End If
    Else
        ' Keep zero available as the existing caller's uninitialized sentinel.
        clockTimerDays = 86400#
        clockTimerReady = True
    End If
    clockTimerLast = rawSeconds
    RPX_Clock = clockShift + clockTimerDays + rawSeconds
End Function

' Operator test only: zero disables; otherwise force Timer fallback and wrap.
' Reconfiguration is allowed only while idle, with TEST_INJECTION explicitly on.
' Align both modes to the previous reading so existing captures keep their epoch.
Public Sub RPX_TestClockWrap(ByVal wrapAtSeconds As Double)
    Dim previousSeconds As Double, rawSeconds As Double
    If RPX_Busy Then err.Raise 5, "RPX_Clock", "TEST_CLOCK_BUSY"
    If CLng(RPX_Setting("TEST_INJECTION", 0)) <> 1 Then err.Raise 5, "RPX_Clock", "TEST_INJECTION_DISABLED"
    If wrapAtSeconds < 0# Or wrapAtSeconds >= 86400# Then err.Raise 5, "RPX_Clock", "TEST_CLOCK_WRAP_RANGE"
    previousSeconds = RPX_Clock()
    clockTestWrapAt = wrapAtSeconds
    clockTestStarted = previousSeconds
    clockTestWraps = 0: clockTestRawJump = 0#: clockTestCorrectedStep = 0#
    clockTimerDays = 0#: clockTimerReady = False
    If wrapAtSeconds > 0# Then
        rawSeconds = CDbl(Timer)
        clockTestRawShift = 86400# - wrapAtSeconds - rawSeconds
        clockTimerLast = 86400# - wrapAtSeconds
        clockTimerReady = True
        clockShift = previousSeconds - clockTimerLast
    Else
#If VBA7 Then
#If Win64 Then
        clockShift = previousSeconds - CDbl(RPX_GetTickCount64()) / 1000#
        Exit Sub
#End If
#End If
        rawSeconds = CDbl(Timer)
        clockTimerLast = rawSeconds: clockTimerReady = True
        clockShift = previousSeconds - rawSeconds
    End If
End Sub

' Read-only evidence; deliberately retained after run_end for the operator runner.
Public Function RPX_TestClockStatus() As String
    Dim sampledSeconds As Double
    sampledSeconds = RPX_Clock()
    RPX_TestClockStatus = "active=" & Abs(CLng(clockTestWrapAt > 0#)) & ";wrap_at=" & CStr(clockTestWrapAt) & ";wraps=" & clockTestWraps
    RPX_TestClockStatus = RPX_TestClockStatus & ";elapsed=" & CStr(sampledSeconds - clockTestStarted) & ";raw_jump=" & CStr(clockTestRawJump) & ";corrected_step=" & CStr(clockTestCorrectedStep)
End Function
