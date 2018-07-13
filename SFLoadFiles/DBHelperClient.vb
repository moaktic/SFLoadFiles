#Region "DIRECTIVES"

Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.Common
Imports System.Collections
Imports System.Configuration

#End Region

Public Class DBHelperClient

#Region "DECLARATIONS"

    Private oFactory As DbProviderFactory
    Private oConnection As DbConnection
    Private oConnectionState As ConnectionState
    Public oCommand As DbCommand
    Private oParameter As DbParameter
    Private oTransaction As DbTransaction
    Private mblTransaction As Boolean

    Private Shared ReadOnly S_CONNECTION As String = ConfigurationManager.AppSettings("DATA.CONNECTIONSTRINGC")
    Private Shared ReadOnly S_PROVIDER As String = ConfigurationManager.AppSettings("DATA.PROVIDER")

#End Region

#Region "ENUMERATORS"

    Public Enum TransactionType As UInteger
        Open = 1
        Commit = 2
        Rollback = 3
    End Enum

#End Region

#Region "STRUCTURES"

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command
    '''Input			:	
    '''OutPut			:	
    '''Comments			:	
    ''' </summary>
    Public Structure Parameters
        Public ParamName As String
        Public ParamValue As Object
        Public ParamDirection As ParameterDirection

        Public Sub New(ByVal Name As String, ByVal Value As Object, ByVal Direction As ParameterDirection)
            ParamName = Name
            ParamValue = Value
            ParamDirection = Direction
        End Sub

        Public Sub New(ByVal Name As String, ByVal Value As Object)
            ParamName = Name
            ParamValue = Value
            ParamDirection = ParameterDirection.Input
        End Sub
    End Structure

#End Region

#Region "CONSTRUCTOR"

    Public Sub New()
        oFactory = DbProviderFactories.GetFactory(S_PROVIDER)
    End Sub

#End Region

#Region "DESTRUCTOR"

    Protected Overrides Sub Finalize()
        Try
            oFactory = Nothing
        Finally
            MyBase.Finalize()
        End Try
    End Sub

#End Region

#Region "CONNECTIONS"

    ''' <summary>
    '''Description	    :	This function is used to Open Database Connection
    '''Input			:	NA
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Public Sub EstablishFactoryConnection()
        '
        '            // This check is not required as it will throw "Invalid Provider Exception" on the contructor itself.
        '            if (0 == DbProviderFactories.GetFactoryClasses().Select("InvariantName='" + S_PROVIDER + "'").Length)
        '                throw new Exception("Invalid Provider");
        '            

        oConnection = oFactory.CreateConnection()

        If oConnection.State = ConnectionState.Closed Then
            oConnection.ConnectionString = ConfigurationManager.AppSettings("DATA.CONNECTIONSTRINGC")
            oConnection.Open()
            oConnectionState = ConnectionState.Open
        End If
    End Sub

    ''' <summary>
    '''Description	    :	This function is used to Close Database Connection
    '''Input			:	NA
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Public Sub CloseFactoryConnection()
        'check for an open connection            
        Try
            If oConnection.State = ConnectionState.Open Then
                oConnection.Close()
                oConnectionState = ConnectionState.Closed
            End If
        Catch oDbErr As DbException
            'catch any SQL server data provider generated error messag
            Throw New Exception(oDbErr.Message)
        Catch oNullErr As System.NullReferenceException
            Throw New Exception(oNullErr.Message)
        Finally
            If oConnection IsNot Nothing Then
                oConnection.Dispose()
            End If
        End Try
    End Sub

#End Region

#Region "TRANSACTION"

    ''' <summary>
    '''Description	    :	This function is used to Handle Transaction Events


    '''Input			:	Transaction Event Type
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Public Sub TransactionHandler(ByVal veTransactionType As TransactionType)
        Select Case veTransactionType
            Case TransactionType.Open
                'open a transaction
                Try
                    oTransaction = oConnection.BeginTransaction()
                    mblTransaction = True
                Catch oErr As InvalidOperationException
                    Throw New Exception("@TransactionHandler - " & oErr.Message)
                End Try
                Exit Select

            Case TransactionType.Commit
                'commit the transaction
                If oTransaction.Connection IsNot Nothing Then
                    Try
                        oTransaction.Commit()
                        mblTransaction = False
                    Catch oErr As InvalidOperationException
                        Throw New Exception("@TransactionHandler - " & oErr.Message)
                    End Try
                End If
                Exit Select

            Case TransactionType.Rollback
                'rollback the transaction
                Try
                    If mblTransaction Then
                        oTransaction.Rollback()
                    End If
                    mblTransaction = False
                Catch oErr As InvalidOperationException
                    Throw New Exception("@TransactionHandler - " & oErr.Message)
                End Try
                Exit Select
        End Select

    End Sub

#End Region

#Region "COMMANDS"

#Region "PARAMETERLESS METHODS"

    ''' <summary>
    '''Description	    :	This function is used to Prepare Command For Execution


    '''Input			:	Transaction, Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	NA
    '''Comments			:	Has to be changed/removed if object based array concept is removed.
    ''' </summary>
    Private Sub PrepareCommand(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String)

        If oConnection.State <> ConnectionState.Open Then
            oConnection.ConnectionString = S_CONNECTION
            oConnection.Open()
            oConnectionState = ConnectionState.Open
        End If

        If oCommand Is Nothing Then
            oCommand = oFactory.CreateCommand()
        End If

        oCommand.Connection = oConnection
        oCommand.CommandText = cmdText
        oCommand.CommandType = cmdType
        oCommand.CommandTimeout = 4800

        If blTransaction Then
            oCommand.Transaction = oTransaction
        End If
    End Sub

#End Region

#Region "OBJECT BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to Prepare Command For Execution


    '''Input			:	Transaction, Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Private Sub PrepareCommand(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,))

        If oConnection.State <> ConnectionState.Open Then
            oConnection.ConnectionString = S_CONNECTION
            oConnection.Open()
            oConnectionState = ConnectionState.Open
        End If

        If oCommand Is Nothing Then
            oCommand = oFactory.CreateCommand()
        End If

        oCommand.Connection = oConnection
        oCommand.CommandText = cmdText
        oCommand.CommandType = cmdType
        oCommand.CommandTimeout = 4800

        If blTransaction Then
            oCommand.Transaction = oTransaction
        End If

        If cmdParms IsNot Nothing Then
            CreateDBParameters(cmdParms)
        End If
    End Sub

#End Region

#Region "STRUCTURE BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to Prepare Command For Execution


    '''Input			:	Transaction, Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Private Sub PrepareCommand(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters())

        If oConnection.State <> ConnectionState.Open Then
            oConnection.ConnectionString = S_CONNECTION
            oConnection.Open()
            oConnectionState = ConnectionState.Open
        End If

        oCommand = oFactory.CreateCommand()
        oCommand.Connection = oConnection
        oCommand.CommandText = cmdText
        oCommand.CommandType = cmdType
        oCommand.CommandTimeout = 4800

        If blTransaction Then
            oCommand.Transaction = oTransaction
        End If

        If cmdParms IsNot Nothing Then
            CreateDBParameters(cmdParms)
        End If
    End Sub

#End Region

#End Region

#Region "PARAMETER METHODS"

#Region "OBJECT BASED"

    ''' <summary>
    '''Description	    :	This function is used to Create Parameters for the Command For Execution

    '''Input			:	2-Dimensional Parameter Array
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Private Sub CreateDBParameters(ByVal colParameters As Object(,))
        For i As Integer = 0 To colParameters.Length \ 2 - 5
            oParameter = oCommand.CreateParameter()
            oParameter.ParameterName = colParameters(i, 0).ToString()
            oParameter.Value = colParameters(i, 1)
            oCommand.Parameters.Add(oParameter)
        Next
    End Sub

#End Region

#Region "STRUCTURE BASED"

    ''' <summary>
    '''Description	    :	This function is used to Create Parameters for the Command For Execution


    '''Input			:	2-Dimensional Parameter Array
    '''OutPut			:	NA
    '''Comments			:	
    ''' </summary>
    Private Sub CreateDBParameters(ByVal colParameters As Parameters())
        For i As Integer = 0 To colParameters.Length - 1
            Dim oParam As Parameters = CType(colParameters(i), Parameters)

            oParameter = oCommand.CreateParameter()
            oParameter.ParameterName = oParam.ParamName
            oParameter.Value = oParam.ParamValue
            oParameter.Direction = oParam.ParamDirection

            oCommand.Parameters.Add(oParameter)
        Next
    End Sub

#End Region

#End Region

#Region "EXCEUTE METHODS"

#Region "PARAMETERLESS METHODS"

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    '''                     Has to be changed/removed if object based array concept is removed.
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal cmdType As CommandType, ByVal cmdText As String) As Integer
        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText)

            Return oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Transaction, Command Type, Command Text, 2-Dimensional Parameter Array, Clear Paramaeters
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    '''                     Has to be changed/removed if object based array concept is removed.
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String) As Integer
        Try
            PrepareCommand(blTransaction, cmdType, cmdText)
            Dim val As Integer = oCommand.ExecuteNonQuery()

            Return val
        Catch ex As Exception
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

#End Region

#Region "OBJECT BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array, Clear Parameters
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,), ByVal blDisposeCommand As Boolean) As Integer
        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Count of Records Affected
    '''Comments			:	Overloaded method. 
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,)) As Integer
        Return ExecuteNonQuery(cmdType, cmdText, cmdParms, True)
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Transaction, Command Type, Command Text, 2-Dimensional Parameter Array, Clear Paramaeters
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,), ByVal blDisposeCommand As Boolean) As Integer
        Try

            PrepareCommand(blTransaction, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Transaction, Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Count of Records Affected
    '''Comments			:	Overloaded function. 
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,)) As Integer
        Return ExecuteNonQuery(blTransaction, cmdType, cmdText, cmdParms, True)
    End Function

#End Region

#Region "STRUCTURE BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Command Type, Command Text, Parameter Structure Array, Clear Parameters
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters(), ByVal blDisposeCommand As Boolean) As Integer
        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Command Type, Command Text, Parameter Structure Array
    '''OutPut			:	Count of Records Affected
    '''Comments			:	Overloaded method. 
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters()) As Integer
        Return ExecuteNonQuery(cmdType, cmdText, cmdParms, True)
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Transaction, Command Type, Command Text, Parameter Structure Array, Clear Parameters
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters(), ByVal blDisposeCommand As Boolean) As Integer
        Try

            PrepareCommand(blTransaction, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to Execute the Command


    '''Input			:	Transaction, Command Type, Command Text, Parameter Structure Array
    '''OutPut			:	Count of Records Affected
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteNonQuery(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters()) As Integer
        Return ExecuteNonQuery(blTransaction, cmdType, cmdText, cmdParms, True)
    End Function

#End Region

#End Region

#Region "READER METHODS"

#Region "PARAMETERLESS METHODS"

    ''' <summary>
    '''Description	    :	This function is used to fetch data using Data Reader	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Data Reader
    '''Comments			:	
    '''                     Has to be changed/removed if object based array concept is removed.
    ''' </summary>
    Public Function ExecuteReader(ByVal cmdType As CommandType, ByVal cmdText As String) As DbDataReader

        ' we use a try/catch here because if the method throws an exception we want to 
        ' close the connection throw code, because no datareader will exist, hence the 
        ' commandBehaviour.CloseConnection will not work
        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText)
            Dim dr As DbDataReader = oCommand.ExecuteReader(CommandBehavior.CloseConnection)
            oCommand.Parameters.Clear()

            Return dr
        Catch ex As Exception
            CloseFactoryConnection()
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

#End Region

#Region "OBJECT BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to fetch data using Data Reader	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Data Reader
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteReader(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms(,) As Object) As DbDataReader

        ' we use a try/catch here because if the method throws an exception we want to 
        ' close the connection throw code, because no datareader will exist, hence the 
        ' commandBehaviour.CloseConnection will not work

        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText, cmdParms)
            Dim dr As DbDataReader = oCommand.ExecuteReader(CommandBehavior.CloseConnection)
            oCommand.Parameters.Clear()

            Return dr
        Catch ex As Exception
            CloseFactoryConnection()
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

#End Region

#Region "STRUCTURE BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to fetch data using Data Reader	


    '''Input			:	Command Type, Command Text, Parameter AStructure Array
    '''OutPut			:	Data Reader
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteReader(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters()) As DbDataReader

        ' we use a try/catch here because if the method throws an exception we want to 
        ' close the connection throw code, because no datareader will exist, hence the 
        ' commandBehaviour.CloseConnection will not work
        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteReader(CommandBehavior.CloseConnection)
        Catch ex As Exception
            CloseFactoryConnection()
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

#End Region

#End Region

#Region "ADAPTER METHODS"

#Region "PARAMETERLESS METHODS"

    ''' <summary>
    '''Description	    :	This function is used to fetch data using Data Adapter	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Data Set
    '''Comments			:	
    '''                     Has to be changed/removed if object based array concept is removed.
    ''' </summary>
    Public Function DataAdapter(ByVal cmdType As CommandType, ByVal cmdText As String) As DataSet

        ' we use a try/catch here because if the method throws an exception we want to 
        ' close the connection throw code, because no datareader will exist, hence the 
        ' commandBehaviour.CloseConnection will not work
        Dim dda As DbDataAdapter = Nothing
        Try
            EstablishFactoryConnection()
            dda = oFactory.CreateDataAdapter()
            PrepareCommand(False, cmdType, cmdText)

            dda.SelectCommand = oCommand
            Dim ds As New DataSet()
            dda.Fill(ds)
            Return ds
        Catch ex As Exception
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

#End Region

#Region "OBJECT BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to fetch data using Data Adapter	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Data Set
    '''Comments			:	
    ''' </summary>
    Public Function DataAdapter(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,)) As DataSet

        ' we use a try/catch here because if the method throws an exception we want to 
        ' close the connection throw code, because no datareader will exist, hence the 
        ' commandBehaviour.CloseConnection will not work
        Dim dda As DbDataAdapter = Nothing
        Try
            EstablishFactoryConnection()
            dda = oFactory.CreateDataAdapter()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            dda.SelectCommand = oCommand
            Dim ds As New DataSet()
            dda.Fill(ds)
            Return ds
        Catch ex As Exception
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

#End Region

#Region "STRUCTURE BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to fetch data using Data Adapter	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Data Set
    '''Comments			:	
    ''' </summary>
    Public Function DataAdapter(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters()) As DataSet

        ' we use a try/catch here because if the method throws an exception we want to 
        ' close the connection throw code, because no datareader will exist, hence the 
        ' commandBehaviour.CloseConnection will not work
        Dim dda As DbDataAdapter = Nothing
        Try
            EstablishFactoryConnection()
            dda = oFactory.CreateDataAdapter()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            dda.SelectCommand = oCommand
            Dim ds As New DataSet()

            dda.Fill(ds)
            Return ds
        Catch ex As Exception
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

#End Region

#End Region

#Region "SCALAR METHODS"

#Region "PARAMETERLESS METHODS"

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal cmdType As CommandType, ByVal cmdText As String) As Object
        Try
            EstablishFactoryConnection()

            PrepareCommand(False, cmdType, cmdText)

            Dim val As Object = oCommand.ExecuteScalar()

            Return val
        Catch ex As Exception
            Throw ex
        Finally
            If oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

#End Region

#Region "OBJECT BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,), ByVal blDisposeCommand As Boolean) As Object
        Try

            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteScalar()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	Overloaded Method. 
    ''' </summary>
    Public Function ExecuteScalar(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,)) As Object
        Return ExecuteScalar(cmdType, cmdText, cmdParms, True)
    End Function

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,), ByVal blDisposeCommand As Boolean) As Object
        Try

            PrepareCommand(blTransaction, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteScalar()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Object(,)) As Object
        Return ExecuteScalar(blTransaction, cmdType, cmdText, cmdParms, True)
    End Function

#End Region

#Region "STRUCTURE BASED PARAMETER ARRAY"

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters(), ByVal blDisposeCommand As Boolean) As Object
        Try
            EstablishFactoryConnection()
            PrepareCommand(False, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteScalar()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
            CloseFactoryConnection()
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	Overloaded Method. 
    ''' </summary>
    Public Function ExecuteScalar(ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters()) As Object
        Return ExecuteScalar(cmdType, cmdText, cmdParms, True)
    End Function

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters(), ByVal blDisposeCommand As Boolean) As Object
        Try

            PrepareCommand(blTransaction, cmdType, cmdText, cmdParms)

            Return oCommand.ExecuteScalar()
        Catch ex As Exception
            Throw ex
        Finally
            If blDisposeCommand AndAlso oCommand IsNot Nothing Then
                oCommand.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    '''Description	    :	This function is used to invoke Execute Scalar Method	


    '''Input			:	Command Type, Command Text, 2-Dimensional Parameter Array
    '''OutPut			:	Object
    '''Comments			:	
    ''' </summary>
    Public Function ExecuteScalar(ByVal blTransaction As Boolean, ByVal cmdType As CommandType, ByVal cmdText As String, ByVal cmdParms As Parameters()) As Object
        Return ExecuteScalar(blTransaction, cmdType, cmdText, cmdParms, True)
    End Function

#End Region

#End Region

End Class
