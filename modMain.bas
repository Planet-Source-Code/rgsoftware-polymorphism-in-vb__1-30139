Attribute VB_Name = "modMain"

Sub Main()
    
    Dim oAnything As Automobile
    Dim oVehicle As ITransportation
    
    Set oAnything = New Automobile
    Set oVehicle = oAnything
    
    'Both objects refer to the same object instance

    oAnything.HorsePower = 400 'Automobile-specific property
    oVehicle.StartUp           'Use polymorphism on this object

    Set oAnything = Nothing
    Set oVehicle = Nothing
    
End Sub
