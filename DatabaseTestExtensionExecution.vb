Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace DatabaseObjects.UnitTestExtensions

    Public Class DatabaseTestExtensionExecution
        Inherits TestExtensionExecution

        Private _databaseTestClass As DatabaseTestClassAttribute
        Private _execution As TestExecution
        Private _testClassInstance As Object

        Public Sub New(databaseTestClass As DatabaseTestClassAttribute)

            If databaseTestClass Is Nothing Then
                Throw New ArgumentNullException
            End If

            Me._databaseTestClass = databaseTestClass

        End Sub

        Public Overrides Sub Initialize(execution As TestExecution)

            Me._execution = execution
            AddHandler execution.AfterTestInitialize, AddressOf AfterTestInitialize

        End Sub

        Private Sub AfterTestInitialize(sender As Object, e As Microsoft.VisualStudio.TestTools.UnitTesting.AfterTestInitializeEventArgs)

            _testClassInstance = e.Instance

        End Sub

        Public Overrides Function CreateTestMethodInvoker(context As Microsoft.VisualStudio.TestTools.UnitTesting.TestMethodInvokerContext) As Microsoft.VisualStudio.TestTools.UnitTesting.ITestMethodInvoker

            Return New DatabaseTestMethodInvoker(_databaseTestClass, _testClassInstance, context)

        End Function

        Public Overrides Sub Dispose()

            RemoveHandler _execution.AfterTestInitialize, AddressOf AfterTestInitialize

        End Sub

    End Class

End Namespace

