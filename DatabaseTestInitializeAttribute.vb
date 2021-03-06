﻿Namespace DatabaseObjects.UnitTestExtensions

    ''' <summary>
    ''' A method marked with this attribute is executed after the [TestInitialize] and before the [TestMethod]. 
    ''' The method is expected to accept an argument of type DatabaseObjects.Database.
    ''' Only one method can be marked with this attribute per class.
    ''' </summary>
    <AttributeUsage(AttributeTargets.Method, AllowMultiple:=False, Inherited:=False)> _
    Public NotInheritable Class DatabaseTestInitializeAttribute
        Inherits Attribute

        Public Sub New()

        End Sub

    End Class

End Namespace

