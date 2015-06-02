'---------------------------------------------------------'
'--------------------Solution Imports---------------------'
'---------------------------------------------------------'

'---------------------------------------------------------'
'---------------------General Imports---------------------'
'---------------------------------------------------------'
Imports System.Data.SqlClient

Namespace DBConn

    Public Class DBConnection
        '---------------------------------------------------------'
        '------------------------Variables------------------------' x
        '---------------------------------------------------------'
        Private _ConnString As String
        Private _Connection As SqlConnection = Nothing
        Private _SQLCommand As SqlCommand = Nothing
        Private _SQLReader As SqlDataReader = Nothing
        Private _ResultList As List(Of List(Of Object))

        '---------------------------------------------------------'
        '------------------------Properties-----------------------'
        '---------------------------------------------------------'
        Public ReadOnly Property ResultList As List(Of List(Of Object))
            Get
                ResultList = _ResultList
            End Get
        End Property

        '---------------------------------------------------------'
        '-------------------------Methods-------------------------'
        '---------------------------------------------------------'
        Public Sub New()
            _ConnString = "Server=YOUR_INFO_HERE;Database=YOUR_INFO_HERE;User Id=YOUR_INFO_HERE;Password=YOUR_INFO_HERE;"
            _Connection = New SqlConnection
            _Connection.ConnectionString = _ConnString
        End Sub

        Public Sub RunQuery(ByVal QueryString As String)
            'Need to sanitize query string/catch if query not well formed
            _Connection.Open()
            _SQLCommand = New SqlCommand(QueryString, _Connection)
            _SQLReader = _SQLCommand.ExecuteReader
            _ResultList = New List(Of List(Of Object))
            If _SQLReader.HasRows Then
                Dim FirstRow As Boolean = True
                Dim i As Integer = 1
                Do While _SQLReader.Read
                    Dim LoopList As List(Of Object)
                    For j As Integer = 0 To _SQLReader.VisibleFieldCount - 1
                        If FirstRow = True Then
                            LoopList = New List(Of Object)
                            LoopList.Add(_SQLReader.GetValue(j))
                            _ResultList.Add(LoopList)
                        Else
                            LoopList = _ResultList.Item(j)
                            LoopList.Add(_SQLReader.GetValue(j))
                        End If
                    Next
                    If i = 1 Then
                        FirstRow = False
                    End If
                    i = i + 1
                Loop
            End If
            _SQLReader.Close()
            _Connection.Close()
        End Sub

    End Class

End Namespace

