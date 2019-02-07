'________________________________________________________________________________________
' Author: THIAGO CARRERE.
' Product:  Main.
' Last Release:  13/01/2019.
' Copyright ©2010-2019 All Rights Reserved.
'________________________________________________________________________________________

' TERMS OF USE:

' This program is free software. Redistribution in the form of source code only,
' and any modification is absolutely permitted.
' The author accepts no liability for anything that may result from the usage of
' this product. 
' This notice must not be removed or altered in any way whatsoever.
'________________________________________________________________________________________

Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Win32
Imports System.IO
Imports Shell32
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Configuration
Imports System.Management
Imports System.Windows.Forms
Imports System.Net.NetworkInformation

Public Class Main



#Region "BANCO DE DADOS"


    Dim conexaoOLEDB As OleDbConnection
    Dim comandoOLEDB As OleDbCommand
    Dim adapterOLEDB As OleDbDataAdapter
    Dim conexaoSQL As SqlConnection
    Dim comandoSQL As SqlCommand
    Dim adapterSQL As SqlDataAdapter
    Dim dtSet As DataSet
		

		''' <summary>
    ''' Função que executa comando sql em base de dados access, no projeto já deve existir o arquivo app.config, com a string de conexão.
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna verdadeiro se tudo ocorrer bem, e falso caso dê algo errado.</returns>
    Public Function BD_retornoBooleanoAccess(ByVal comando As String) As Boolean

        'conexaoOLEDB = New OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)
        'Dim chave As String = System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString
        'chave = chave.Replace("a", "5")
        'conexaoOLEDB = New OleDbConnection(chave)

        unlock()'refere-se a uma função que pode ser criada e utilizada a criterio do desenvolvedor para desbloquear bases com senha*

        Try
            conexaoOLEDB.Open()
            comandoOLEDB = New OleDbCommand(comando, conexaoOLEDB)
            comandoOLEDB.ExecuteNonQuery()
            BD_retornoBooleanoAccess = True
        Catch ex As Exception
            BD_retornoBooleanoAccess = False
        Finally
            conexaoOLEDB.Close()
        End Try

        Return BD_retornoBooleanoAccess

    End Function


    ''' <summary>
    ''' Função que executa comando sql em base de dados access, no projeto já deve existir o arquivo app.config, com a string de conexão.
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna um dataset populado caso dê tudo certo.</returns>
    Public Function BD_retornoDataSetAccess(ByVal comando As String) As DataSet

        'conexaoOLEDB = New OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        unlock()

        Try
            conexaoOLEDB.Open()
            dtSet = New DataSet
            adapterOLEDB = New OleDbDataAdapter(comando, conexaoOLEDB)
            adapterOLEDB.Fill(dtSet)
            Return dtSet
        Catch ex As Exception
            Return Nothing
        Finally
            conexaoOLEDB.Close()
        End Try

    End Function


    ''' <summary>
    ''' Função que executa comando sql em base de dados access, no projeto já deve existir o arquivo app.config, com a string de conexão.
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna uma string caso dê tudo certo</returns>
    Public Function BD_retornoStringAccess(ByVal comando As String) As String

        'conexaoOLEDB = New OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        unlock()

        Try
            conexaoOLEDB.Open()
            comandoOLEDB = New OleDbCommand(comando, conexaoOLEDB)
            BD_retornoStringAccess = comandoOLEDB.ExecuteScalar()
            Return BD_retornoStringAccess
        Catch ex As Exception
            Return Nothing
        Finally
            conexaoOLEDB.Close()
        End Try

    End Function


    ''' <summary>
    ''' Função que executa comando sql em base de dados access, no projeto já deve existir o arquivo app.config, com a string de conexão.
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna um inteiro caso dê tudo certo</returns>
    Public Function BD_retornoInteiroAccess(ByVal comando As String) As Integer

        Dim retorno As Integer = 0
        'conexaoOLEDB = New OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        unlock()

        Try
            conexaoOLEDB.Open()
            comandoOLEDB = New OleDbCommand(comando, conexaoOLEDB)
            retorno = comandoOLEDB.ExecuteScalar()
            Return retorno
        Catch ex As Exception
            Return Nothing
        Finally
            conexaoOLEDB.Close()
        End Try

    End Function


    ''' <summary>
    ''' Função que efetua a cópia de uma base de dados access temporária, necessita ser configurada*
    ''' </summary>
    Public Sub BD_copiarBdTemp()

        'conexaoOLEDB = New OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        unlock()

        If conexaoOLEDB.DataSource <> "C:\BIBAC\BDBIB.mdb" Then
            File.Copy(conexaoOLEDB.DataSource, "C:\BIBAC\BDBIB.mdb")
        End If

    End Sub


    ''' <summary>
    ''' Função que deleta a cópia de uma base de dados access temporária, necessita ser configurada*
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BD_deletarBdTemp()

        'conexaoOLEDB = New OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        unlock()

        If conexaoOLEDB.DataSource <> "C:\BIBAC\BDBIB.mdb" Then
            File.Delete("C:\BIBAC\BDBIB.mdb")
        End If

    End Sub


    ''' <summary>
    ''' Função que executa comando sql em base de dados sql server
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna verdadeiro se tudo der certo e falso caso haja algum erro.</returns>
    Public Function BD_retornoBooleanoSQL(ByVal comando As String) As Boolean

        conexaoSQL = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)
        comandoSQL = New SqlCommand(comando, conexaoSQL)

        Try
            conexaoSQL.Open()
            comandoSQL.ExecuteNonQuery()
            BD_retornoBooleanoSQL = True
        Catch ex As Exception
            BD_retornoBooleanoSQL = False
        Finally
            conexaoSQL.Close()
        End Try

    End Function


    ''' <summary>
    ''' Função que executa comando sql em base de dados sql server
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna um dataset populado caso tudo dê certo</returns>
    Public Function BD_retornoDataSetSQL(ByVal comando As String) As DataSet

        'conexaoSQL = New SqlConnection("data source=" & servidor & ";initial catalog=" & nomeBase & ";integrated security=true")
        conexaoSQL = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        Try
            conexaoSQL.Open()
            dtSet = New DataSet
            adapterSQL = New SqlDataAdapter(comando, conexaoSQL)
            adapterSQL.Fill(dtSet)
            Return dtSet
        Catch ex As Exception
            Return Nothing
        Finally

            conexaoSQL.Close()

        End Try

    End Function


    ''' <summary>
    ''' Função que executa comando sql em base de dados sql server
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna um inteiro caso tudo dê certo</returns>
    Public Function BD_retornoInteiroSQL(ByVal comando As String) As Integer

        Dim retorno As Integer = 0
        'conexaoSQL = New SqlConnection("data source=" & servidor & ";initial catalog=" & nomeBase & ";integrated security=true")
        conexaoSQL = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        Try
            conexaoSQL.Open()
            comandoSQL = New SqlCommand(comando, conexaoSQL)
            retorno = comandoSQL.ExecuteScalar()
            Return retorno
        Catch ex As Exception
            Return Nothing
        Finally

            conexaoSQL.Close()

        End Try

    End Function


    ''' <summary>
    ''' Função que executa comando sql em base de dados sql server
    ''' </summary>
    ''' <param name="comando">Comando SQL</param>
    ''' <returns>retorna uma string caso tudo dê certo</returns>
    Public Function BD_retornoStringSQL(ByVal comando As String) As String

        Dim retorno As String = String.Empty
        conexaoSQL = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("conexaoBD").ConnectionString)

        Try
            conexaoSQL.Open()
            comandoSQL = New SqlCommand(comando, conexaoSQL)
            retorno = comandoSQL.ExecuteScalar()
            Return retorno
        Catch ex As Exception
            Return Nothing
        Finally

            conexaoSQL.Close()

        End Try

    End Function


    ''' <summary>
    ''' Função que efetua o backup de base de dados em sql server
    ''' </summary>
    ''' <param name="servidor">Nome do servidor ex: overture\sqlexpress</param>
    ''' <param name="nomeBase">Nome do banco de dados, ex: sisLoja</param>
    ''' <param name="unidade">Onde será efetuado o backup do banco, ex: C, D, H, T... </param>
    ''' <returns>retorna verdadeiro se tudo ocorrer bem e falso se houver algum erro</returns>
    Public Function BD_backupSQL(ByVal servidor As String, ByVal nomeBase As String, ByVal unidade As Char) As Boolean

        Dim SQL As String = "BACKUP DATABASE " & nomeBase & " TO  DISK = " & "N'" & unidade & ":\" & nomeBase & "'" & " WITH NOFORMAT, NOINIT,  NAME = " & "N'" & nomeBase & "-Full Database Backup'" & " , SKIP, NOREWIND, NOUNLOAD,  STATS = 10"

        If BD_retornoBooleanoSQL(SQL) Then
            BD_backupSQL = True
        Else
            BD_backupSQL = False
        End If

    End Function

#End Region


End Class
