VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "escripta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim login, pass1 As Integer
Dim ctr As Integer

Function escriptar(valor As String) As String
 On Error Resume Next
 Dim PassNew As String
 Dim Passtemp As String
 pass1 = Len(Trim(valor))
 ctr = 1
 Do While ctr <= pass1
 PassNew = CStr(PassNew) & Chr((Asc(Mid(Trim(valor), ctr, 1)) + 7))
 ctr = ctr + 1
 Loop
 escriptar = PassNew
End Function

Function desescriptar(valor As String) As String
 On Error Resume Next
 Dim Passlength As Integer, Cntr As Integer
 Dim tempChar As String
 Dim OldPass As String
 Cntr = 1
 Passlength = Len(valor)
 Do While Cntr <= Passlength
 OldPass = OldPass + Chr((Asc(Mid(Trim(valor), Cntr, 1)) - 7))
 Cntr = Cntr + 1
 Loop
 desescriptar = OldPass
End Function


