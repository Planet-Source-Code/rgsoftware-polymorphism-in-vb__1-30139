<div align="center">

## Polymorphism in VB


</div>

### Description

In this tutorial, I will show you how to support polymorphism in a COM compliant form, in Visual Basic, much like Java and C++.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-26 17:51:22
**By**             |[RGSoftware](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rgsoftware.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Object Oriented Programming \(OOP\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/object-oriented-programming-oop__1-47.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Polymorphi4411712262001\.zip](https://github.com/Planet-Source-Code/rgsoftware-polymorphism-in-vb__1-30139/archive/master.zip)





### Source Code

```

In this tutorial, I will show you how to support polymorphism in a COM compliant form, in Visual Basic, much like Java and C++.
About the only difference between other OO languages and VB's polymorphism, is that VB will shield you from implementing the default interface of a class.
To begin with, you will need to design an Interface. An Interface is simply a class (or an IDL) that has basic method/property prototypes - without any actual blocks of code - just empty members.
Once your interface class has been defined, you will then be able to create your worker classes which will implement and extend the basic members of your interface class.
Any one class can implement several interfaces, at the same time and any interface can be implemented by several other classes. This means that you can early bind to the same interface on two or more completely different objects! Why would you do this? Because it can prevent many runtime errors (VB will verify all property and methods during compile time, just as if you've added a reference to your object).
VB can support polymorphism via late binding, also. But of course, the problem often encountered with late binding is that VB won't be able to verify all the properties and methods for you. Also, you can run into method overloading problems in VB with late binding. For instance, if two objects use the same method name (e.g. "LoadFile(FileName As String)" and "LoadFile("FileNames() As String)"), nothing will prevent VB from passing a different data type (an array in this example) to the other method. This won't generate compile-time errors, it will generate RUN-TIME errors that are unexpected!
In order to use interfaces in VB, you must use the Implements keyword in a class, in addition to the default interface (which can be defined in an IDL (interface definition language) file, not just in a VB class).
In the code example that I've provided, there are four classes. ITransportation is the object Interface. Then there are 3 other classes named
"Plane", "Train" and "Automobile" - all of which implement the interface class.
When using polymorphism, it is mandatory for each class to support all of the members of the interface. In other words, the classes are allowed to implement more methods/properties and extend the ITransportation interface, but never allowed to remove methods/properties that are already defined in the Interface class.
So go ahead now and experiment with the modMain.bas module to see how polymorphism works! Try adding methods and properties to the Train, Plane and Automobile classes.
Send questions and comments to rgardner@rgsoftware.com
###
Copyright 2000 by Cobalt A.I. Software http://www.cobaltai.com
```

