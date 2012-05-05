Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace DatabaseObjects.UnitTestExtensions

    ''' <summary>
    ''' Allows functions marked with [TestMethod] to be executed with different databases.
    ''' Each test method is executed multiple times for each database specified.
    ''' The method must accept an argument of type DatabaseObjects.Database.
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()>
    Public Class DatabaseTestClassAttribute
        Inherits TestClassExtensionAttribute

        <NonSerialized()>
        Private _execution As DatabaseTestExtensionExecution = Nothing

        Public Property ConnectionStringNames As String()

        Public Sub New()

        End Sub

        Public Overrides ReadOnly Property ExtensionId As System.Uri
            Get

                Return New Uri("urn:DatabaseTestClassAttribute")

            End Get
        End Property

        Public Overrides Function GetExecution() As Microsoft.VisualStudio.TestTools.UnitTesting.TestExtensionExecution

            If _execution Is Nothing Then
                _execution = New DatabaseTestExtensionExecution(Me)
            End If

            Return _execution

        End Function

    End Class

End Namespace

