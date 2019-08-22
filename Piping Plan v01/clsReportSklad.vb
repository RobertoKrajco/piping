Imports System.Data.DataSet
Public Class clsReportSklad
    Inherits DataSet

    Private dtc As DataTable
    Public Sub New()

        dtc = New DataTable("dtVychystanie")

        'dtc.Columns.Add("id", GetType(Integer))
        dtc.Columns.Add("DatumStavby", GetType(Date))
        dtc.Columns.Add("Skupina")
        dtc.Columns.Add("JednotkaWJ")
        dtc.Columns.Add("PN")
        dtc.Columns.Add("QTY",GetType(Double))
        dtc.Columns.Add("VyplanovanyWJ")
        dtc.Columns.Add("VyplanovaneMnozstvo",GetType(Double))
        dtc.Columns.Add("VychystanieStatus",GetType(Boolean))
        dtc.Columns.Add("Linka")
        dtc.Columns.Add("ZaplanovanyPN")

        dtc.AcceptChanges()

        Tables.Add(dtc)

    End Sub

End Class
