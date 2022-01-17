VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMessage As String

Public Property Get Message() As String
    Message = pMessage
End Property

Public Property Let Message(RHS As String)
    pMessage = RHS
End Property

Public Function GetMessage() As String
    GetMessage = pMessage
End Function

Public Function GetValue() As Long
    GetValue = 1
End Function

Public Sub MethodSub()
    'Debug.Print "MethodSub"
End Sub

Public Function MethodFunction()
    'Debug.Print "MethodFunction"
End Function

Public Sub MethodArg(Arg)
    'Debug.Print "MethodArg"
End Sub
