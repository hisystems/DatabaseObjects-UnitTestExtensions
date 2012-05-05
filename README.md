DatabaseObjects UnitTestExtensions
==================================

Overview
--------
Unit test extensions for Visual Studio 2010 that provides the ability to test multiple databases with one function. The databases to test are specified by DatabaseTestClassAttribute by specifying the connection string names from the app.config. Each <TestMethod> is optionally passed a DatabaseObjects.Database argument. Methods marked with <DatabaseTestInitialize> and <DatabaseTestCleanup> can also be specified that will run before and after the <TestMethod>. These initialization classes also accept a DatabaseObjects.Database argument.

License
-------
The library can be used for commercial and non-commercial purposes.

Setup
-----
See the **Setup** section in https://github.com/hisystems/DatabaseObjects-UnitTests/blob/master/README.md for details on how to install the unit test extension.
