Attribute VB_Name = "Func"
Option Explicit

Public Function Init(o As IClass, ParamArray x()) As Object
    Dim v()
    
    v = x

    Set Init = o.Constructor(v)
End Function
