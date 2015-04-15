'---------------------------------------------------------'
'--------------------Solution Imports---------------------'
'---------------------------------------------------------'

'---------------------------------------------------------'
'---------------------General Imports---------------------'
'---------------------------------------------------------'
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel

Namespace ExcelExport

    Public Class ExcelExport
        '---------------------------------------------------------'
        '------------------------Variables------------------------'
        '---------------------------------------------------------'
        Private _ExcelApp As Application
        Private _WrkBk As Workbook
        Private _WrkSheet As Worksheet

        '---------------------------------------------------------'
        '------------------------Properties-----------------------'
        '---------------------------------------------------------'

        '---------------------------------------------------------'
        '-------------------------Methods-------------------------'
        '---------------------------------------------------------'
        Public Sub New()
            _ExcelApp = New Application
            _ExcelApp.Visible = False
        End Sub

        Public Sub RunExport(ByVal FilePath As String, ByVal ObjectList As List(Of Object), ByVal HeaderArray As String())
            'Be sure to sanitize the file name before use, if not determined by code'
            _WrkBk = _ExcelApp.Workbooks.Add
            _WrkSheet = _WrkBk.Worksheets.Item(1)

            'Create column headers based on header array passed in
            For i As Integer = 1 To HeaderArray.Count
                _WrkSheet.Cells(1, i) = HeaderArray(i - 1)
            Next
            Dim LoopObject As New Object
            'For every object in the input list, get attributes that match the header for that column, and write to Excel
            For i As Integer = 1 To ObjectList.Count
                LoopObject = ObjectList.Item(i - 1)
                For j As Integer = 1 To HeaderArray.Count
                    _WrkSheet.Cells(i + 1, j) = GetPropByName(LoopObject, HeaderArray(j - 1))
                Next
            Next
            'Save the workbook to the work directory
            _WrkBk.SaveAs(FilePath & "yourfilename.xlsx")
            _WrkBk.Close()
        End Sub

        'Get the value of an object's attribute by the property's name
        Private Function GetPropByName(ByVal ParentObject As Object, ByVal PropertyName As String) As Object
            Dim ObjectType As Type = ParentObject.GetType()
            Dim PropInfo As System.Reflection.PropertyInfo = ObjectType.GetProperty(PropertyName)
            Dim PropValue As Object = PropInfo.GetValue(ParentObject, Reflection.BindingFlags.GetProperty, Nothing, Nothing, Nothing)
            Return PropValue
        End Function

    End Class

End Namespace

