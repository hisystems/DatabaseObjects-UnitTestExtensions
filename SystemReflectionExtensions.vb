Imports System.Runtime.CompilerServices
Imports System.IO

Namespace DatabaseObjects.UnitTestExtensions

    Friend Module SystemReflectionExtensions

        <Extension()>
        Public Function FullName(ByVal methodInfo As System.Reflection.MethodInfo) As String

            Return methodInfo.DeclaringType.FullName & System.Type.Delimiter & methodInfo.Name

        End Function

        <Extension()>
        Public Function DirectoryLocation(ByVal assembly As System.Reflection.Assembly) As String

            Return Path.GetDirectoryName(assembly.Location)

        End Function

        ''' <summary>
        ''' Returns the method that has an attribute of type T.
        ''' Returns Nothing if there is no matching method that has the custom attribute specified.
        ''' Throws an exception if the attribute is specified multiple times on the same type.
        ''' </summary>
        ''' <typeparam name="T">The attribute type</typeparam>
        <Extension()>
        Public Function GetSingleMethodWithCustomAttribute(Of T)(sourceType As Type) As System.Reflection.MethodInfo

            Dim foundMethods =
                sourceType.GetMethods() _
                .Select(Function(method) New With {.Method = method, .CustomAttributes = method.GetCustomAttributes(GetType(T), inherit:=False)}) _
                .Where(Function(methodAndAttribute) methodAndAttribute.CustomAttributes.Length > 0) _
                .Select(Function(methodAndAttribute) methodAndAttribute.Method) _
                .ToArray()

            If foundMethods.Length = 0 Then
                Return Nothing
            ElseIf foundMethods.Length > 1 Then
                Throw New ArgumentException("Attribute " & GetType(T).FullName & " has been specified multiple times on " & sourceType.FullName)
            Else
                Return foundMethods(0)
            End If

        End Function

    End Module

End Namespace

