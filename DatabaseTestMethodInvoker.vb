Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Configuration
Imports System.Reflection

Namespace DatabaseObjects.UnitTestExtensions

    Public Class DatabaseTestMethodInvoker
        Implements ITestMethodInvoker

        Private _databaseTestClass As DatabaseTestClassAttribute
        Private _invokerContext As TestMethodInvokerContext
        Private _testClassInstance As Object

        Public Sub New(databaseTestClass As DatabaseTestClassAttribute, testClassInstance As Object, invokerContext As TestMethodInvokerContext)

            If databaseTestClass Is Nothing Then
                Throw New ArgumentNullException
            ElseIf invokerContext Is Nothing Then
                Throw New ArgumentNullException
            ElseIf testClassInstance Is Nothing Then
                Throw New ArgumentNullException
            End If

            Me._databaseTestClass = databaseTestClass
            Me._invokerContext = invokerContext
            Me._testClassInstance = testClassInstance

        End Sub

        Private Function Invoke(ParamArray parameters() As Object) As Microsoft.VisualStudio.TestTools.UnitTesting.TestMethodInvokerResult Implements Microsoft.VisualStudio.TestTools.UnitTesting.ITestMethodInvoker.Invoke

            Dim testMethod = Me._invokerContext.TestMethodInfo
            Dim testMethodParmeters = testMethod.GetParameters()
            Dim databaseObjectsAssembly = System.Reflection.Assembly.LoadFile(System.IO.Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().DirectoryLocation, "DatabaseObjects.dll"))
            Dim databaseObjectsDatabaseType = databaseObjectsAssembly.GetType("DatabaseObjects.Database")
            Dim databaseObjectsConnectionType = databaseObjectsAssembly.GetType("DatabaseObjects.Database+ConnectionType")

            'If no DatabaseObjects.Database argument is specified execute as per normal
            If testMethodParmeters.Length = 0 Then
                Return Me._invokerContext.InnerInvoker.Invoke(Nothing)
            ElseIf testMethodParmeters.Length > 1 Or Not testMethodParmeters(0).ParameterType.Equals(databaseObjectsDatabaseType) Then
                Throw New InvalidOperationException(testMethod.FullName & " must accept one argument of type " & databaseObjectsDatabaseType.FullName)
            Else
                Dim result As New TestMethodInvokerResult
                Dim databaseTestInitializeMethod As System.Reflection.MethodInfo = _testClassInstance.GetType.GetSingleMethodWithCustomAttribute(Of DatabaseTestInitializeAttribute)()
                Dim databaseTestCleanupMethod As System.Reflection.MethodInfo = _testClassInstance.GetType.GetSingleMethodWithCustomAttribute(Of DatabaseTestCleanupAttribute)()
                Dim databaseConstructor = databaseObjectsDatabaseType.GetConstructor(New System.Type() {GetType(String), databaseObjectsConnectionType})
                Dim databaseConstructorWithIDbConnection = databaseObjectsDatabaseType.GetConstructor(New System.Type() {GetType(System.Data.IDbConnection), databaseObjectsConnectionType})

                For Each database In
                    _databaseTestClass.ConnectionStringNames _
                    .Select(Function(connectionStringName) System.Configuration.ConfigurationManager.ConnectionStrings(connectionStringName)) _
                    .Select(Function(connection) New With {
                        .Name = connection.Name,
                        .Value = CreateDatabaseInstance(connection, databaseObjectsConnectionType, databaseConstructor, databaseConstructorWithIDbConnection)
                    })

                    Me._invokerContext.TestContext.WriteLine("_______________________________________")
                    Me._invokerContext.TestContext.WriteLine("Testing " & database.Name & "...")

                    If databaseTestInitializeMethod IsNot Nothing Then
                        Try
                            databaseTestInitializeMethod.Invoke(_testClassInstance, {database.Value})
                        Catch ex As Exception
                            result.Exception = New Exception("Initialization method " & databaseTestInitializeMethod.FullName & " threw exception. " & ex.Message, ex)
                            Exit For
                        End Try
                    End If

                    Try
                        result = Me._invokerContext.InnerInvoker.Invoke(database.Value)
                    Catch ex As Exception
                        result.Exception = New Exception("Test method " & Me._invokerContext.TestMethodInfo.FullName & " threw exception. " & ex.Message, ex)
                        Exit For
                    End Try

                    If databaseTestCleanupMethod IsNot Nothing Then
                        Try
                            databaseTestCleanupMethod.Invoke(_testClassInstance, {database.Value})
                        Catch ex As Exception
                            result.Exception = New Exception("Cleanup method " & databaseTestCleanupMethod.FullName & " threw exception. " & ex.Message, ex)
                            Exit For
                        End Try
                    End If
                Next

                Return result
            End If

        End Function

        ''' <summary>
        ''' The connection string's provider name may be of the following formats:
        ''' providerName="MySQL" or providerName="SqlLite;System.Data.SQLite.SQLiteConnection, System.Data.SQLite".
        ''' In the first case the DatabaseObject.Database(string, ConnectionType) constructor should be called.
        ''' In the second case the DatabaseObject.Database(IDbConnection, ConnectionType) constructor should be called
        ''' using the type specified in the providerName which should be of type IDbConnection. 
        ''' </summary>
        Private Shared Function CreateDatabaseInstance(connectionString As ConnectionStringSettings, databaseObjectsConnectionType As Type, databaseConstructor As System.Reflection.ConstructorInfo, databaseConstructorWithIDbConnection As System.Reflection.ConstructorInfo) As Object

            If System.Enum.IsDefined(databaseObjectsConnectionType, connectionString.ProviderName) Then
                Return databaseConstructor.Invoke({connectionString.ConnectionString, System.Enum.Parse(databaseObjectsConnectionType, connectionString.ProviderName)})
            Else
                Dim connectionTypeAndProvider = connectionString.ProviderName.Split(";"c)
                If connectionTypeAndProvider.Length <> 2 Then
                    Throw New InvalidOperationException(connectionString.Name + ".ProviderName " + connectionString.ProviderName + " is invalid")
                End If

                Dim connectionType = System.Enum.Parse(databaseObjectsConnectionType, connectionTypeAndProvider(0))
                Dim connectionObjectTypeAndAssembly = connectionTypeAndProvider(1).Replace(" ", String.Empty).Split(",")
                Dim connectionObjectAssembly = Assembly.LoadFile(System.IO.Path.Combine(Assembly.GetExecutingAssembly().DirectoryLocation, connectionObjectTypeAndAssembly(1) + ".dll"))
                Dim connectionObject = connectionObjectAssembly.GetType(connectionObjectTypeAndAssembly(0), throwOnError:=True)
                Dim connectionObjectConstructor = connectionObject.GetConstructor({GetType(String)})
                If connectionObjectConstructor Is Nothing Then
                    Throw New InvalidOperationException("A public constructor with a single string argument could not be found for " & connectionObject.FullName)
                End If

                Dim connectionObjectInstance = connectionObjectConstructor.Invoke({connectionString.ConnectionString})  'Invokes the Constructor passing the connection string as an argument - it assumes this is the case.

                Return databaseConstructorWithIDbConnection.Invoke({connectionObjectInstance, connectionType})
            End If

        End Function

    End Class

End Namespace