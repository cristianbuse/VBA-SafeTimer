VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents st As SafeTimer
Attribute st.VB_VarHelpID = -1

Public Sub StartDemo()
    Set st = New SafeTimer
    st.StartTimer 50
End Sub

Private Sub st_TimerCall(ByVal wTime As Long, varData As Variant)
    Debug.Print wTime
End Sub
