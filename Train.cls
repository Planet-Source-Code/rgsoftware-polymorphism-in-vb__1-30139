VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Train"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ITransportation

Private Function ITransportation_BuildEngine() As Variant
'
End Function

Private Property Let ITransportation_FuelType(ByVal RHS As Double)
'
End Property

Private Property Get ITransportation_FuelType() As Double
'
End Property

Private Property Let ITransportation_FuelLevel(ByVal RHS As Double)
'
End Property

Private Property Get ITransportation_FuelLevel() As Double
'
End Property

Private Function ITransportation_StartUp() As Variant
'
End Function
