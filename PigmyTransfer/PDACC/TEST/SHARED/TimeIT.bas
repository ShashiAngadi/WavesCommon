Attribute VB_Name = "basTimeIT"
Option Explicit

Declare Function QueryPerformanceCounter Lib "kernel32" _
                           (X As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" _
                           (X As Currency) As Boolean
Declare Function timeGetTime Lib "winmm.dll" () As Long

Private m_Starts As Currency
Private m_Ends As Currency
Private m_OverHead As Currency
Private m_Frequency As Currency


Public Sub StartTimer()
  
    
  QueryPerformanceCounter m_Starts
  QueryPerformanceCounter m_Ends
  
  m_OverHead = m_Ends - m_Starts        ' determine API overhead
  
  m_Starts = 0
  m_Ends = 0
  
  QueryPerformanceCounter m_Starts  ' time loop

End Sub
Public Sub StopTimer()
  
  QueryPerformanceCounter m_Ends
  QueryPerformanceFrequency m_Frequency
  
  'Debug.Print "("; m_Starts; "-"; m_Ends; "-"; m_OverHead; ") /"; m_Frequency
  'Debug.Print "100 additions took";
  Debug.Print Format((m_Ends - m_Starts - m_OverHead) / m_Frequency, "00.00"); " seconds"
   
  m_Ends = 0
  m_Starts = 0
  m_OverHead = 0
  m_Frequency = 0
  
End Sub

