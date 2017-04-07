Option Strict Off
Option Explicit On

Imports NetOffice.OutlookApi

Namespace CompuMaster.Data.Outlook

    Friend Class ItemTools

        Public Shared Function ParentFolder(item As NetOffice.COMObject) As MAPIFolder
            Return CType(CType(item, Object).Parent, MAPIFolder)
        End Function

        Public Shared Function ObjectClass(item As NetOffice.COMObject) As NetOffice.OutlookApi.Enums.OlObjectClass
            Return CType(CType(item, Object).Class, NetOffice.OutlookApi.Enums.OlObjectClass)
        End Function

        Public Shared Function Subject(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).Subject, String)
        End Function

    End Class

End Namespace