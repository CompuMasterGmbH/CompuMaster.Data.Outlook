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

        Public Shared Function Body(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).Body, String)
        End Function

        Public Shared Function HTMLBody(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).HTMLBody, String)
        End Function

        Public Shared Function RTFBody(item As NetOffice.COMObject) As Object
            Return CType(CType(item, Object).RTFBody, Object)
        End Function

        Public Shared Function BodyFormat(item As NetOffice.COMObject) As NetOffice.OutlookApi.Enums.OlBodyFormat
            Return CType(CType(item, Object).BodyFormat, NetOffice.OutlookApi.Enums.OlBodyFormat)
        End Function

        Public Shared Function CC(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).CC, String)
        End Function

        Public Shared Function BCC(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).BCC, String)
        End Function

        Public Shared Function [To](item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).To, String)
        End Function

        Public Shared Function TaskSubject(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).TaskSubject, String)
        End Function

        Public Shared Function SenderEmailAddress(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).SenderEmailAddress, String)
        End Function

        Public Shared Function SenderName(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).SenderName, String)
        End Function

        Public Shared Function SenderEmailType(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).SenderEmailType, String)
        End Function
        Public Shared Function EntryID(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).EntryID, String)
        End Function
        Public Shared Function ReceivedByName(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).ReceivedByName, String)
        End Function
        Public Shared Function ReceivedByEntryID(item As NetOffice.COMObject) As String
            Return CType(CType(item, Object).ReceivedByEntryID, String)
        End Function
        Public Shared Function CreationTime(item As NetOffice.COMObject) As DateTime
            Return CType(CType(item, Object).CreationTime, DateTime)
        End Function
        Public Shared Function UnRead(item As NetOffice.COMObject) As Boolean
            Return CType(CType(item, Object).UnRead, Boolean)
        End Function
        Public Shared Function ReceivedTime(item As NetOffice.COMObject) As DateTime
            Return CType(CType(item, Object).ReceivedTime, DateTime)
        End Function
        Public Shared Function LastModificationTime(item As NetOffice.COMObject) As DateTime
            Return CType(CType(item, Object).LastModificationTime, DateTime)
        End Function
        Public Shared Function ReminderTime(item As NetOffice.COMObject) As DateTime
            Return CType(CType(item, Object).ReminderTime, DateTime)
        End Function
        Public Shared Function SentOn(item As NetOffice.COMObject) As DateTime
            Return CType(CType(item, Object).SentOn, DateTime)
        End Function
        Public Shared Function Sensitivity(item As NetOffice.COMObject) As NetOffice.OutlookApi.Enums.OlSensitivity
            Return CType(CType(item, Object).Body, NetOffice.OutlookApi.Enums.OlSensitivity)
        End Function
        Public Shared Function Importance(item As NetOffice.COMObject) As NetOffice.OutlookApi.Enums.OlImportance
            Return CType(CType(item, Object).Body, NetOffice.OutlookApi.Enums.OlImportance)
        End Function
        Public Shared Sub Move(item As NetOffice.COMObject, destinationFolder As MAPIFolder)
            CType(item, Object).Move(destinationFolder)
        End Sub
        Public Shared Function Recipients(item As NetOffice.COMObject) As NetOffice.OutlookApi.Recipients
            Return CType(CType(item, Object).Body, NetOffice.OutlookApi.Recipients)
        End Function

        Public Shared Function ItemProperties(item As NetOffice.COMObject) As NetOffice.OutlookApi.ItemProperties
            Return CType(CType(item, Object).Body, NetOffice.OutlookApi.ItemProperties)
        End Function

    End Class

End Namespace