Option Strict On
Option Explicit On

Imports System
Imports NetOffice.OutlookApi
Imports System.Net

Namespace CompuMaster.Data.Outlook

    ''' <summary>
    ''' Represents a folder in Exchange
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Directory

        Private _outlookWrapper As OutlookApp
        Private _IsRootElementForSubFolderQuery As Boolean = False
        Private _folder As MAPIFolder
        Private _parentDirectory As Directory
        Private _parentFolder As MAPIFolder

        Public Sub New(outlookWrapper As OutlookApp, ByVal folder As MAPIFolder)
            _folder = folder
            _IsRootElementForSubFolderQuery = True
            _outlookWrapper = outlookWrapper
        End Sub

        Friend Sub New(outlookWrapper As OutlookApp, ByVal folder As MAPIFolder, ByVal parentFolder As MAPIFolder)
            _folder = folder
            _parentFolder = parentFolder
            If parentFolder Is Nothing Then
                _IsRootElementForSubFolderQuery = True
            End If
            _outlookWrapper = outlookWrapper
        End Sub

        Friend Sub New(outlookWrapper As OutlookApp, ByVal folder As MAPIFolder, ByVal parentDirectory As Directory)
            _folder = folder
            _parentDirectory = parentDirectory
            If parentDirectory Is Nothing Then
                _IsRootElementForSubFolderQuery = True
            End If
            _outlookWrapper = outlookWrapper
        End Sub

        ''' <summary>
        ''' Gives full access to the Exchange Managed API for this folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property OutlookFolder() As MAPIFolder
            Get
                Return _folder
            End Get
        End Property

        Public ReadOnly Property OutlookApp As OutlookApp
            Get
                Return _outlookWrapper
            End Get
        End Property

        Public Function Item(index As Integer) As CompuMaster.Data.Outlook.Item
            Return New CompuMaster.Data.Outlook.Item(Me._outlookWrapper, CType(Me.OutlookFolder.Items()(index), NetOffice.COMObject), Me)
        End Function

        Public Function ItemsRange(startIndex As Integer, length As Integer) As CompuMaster.Data.Outlook.Item()
            Dim Results As New List(Of NetOffice.COMObject)
            Dim MyCounter As Integer = 0
            For Each item As Object In Me.OutlookFolder.Items
                If MyCounter >= startIndex And MyCounter <= startIndex + length Then
                    Results.Add(CType(item, NetOffice.COMObject))
                End If
                MyCounter += 1
            Next
            Return Convert2Items(Me, Results)
        End Function

        Public Function ItemsAll() As CompuMaster.Data.Outlook.Item()
            Return Convert2Items(Me, Me.OutlookFolder.Items)
        End Function

        Private Function Convert2Items(dir As Directory, items As NetOffice.OutlookApi._Items) As Item()
            Dim Result As New List(Of Item)
            For Each item As Object In items
                Result.Add(New Item(dir.OutlookApp, CType(item, NetOffice.COMObject), dir))
            Next
            'For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Count) - 1
            '    Result.Add(New Item(dir.OutlookApp, CType(items(MyItemCounter), NetOffice.COMObject), dir))
            'Next
            Return Result.ToArray
        End Function
        Private Function Convert2Items(dir As Directory, items As List(Of NetOffice.COMObject)) As Item()
            Dim Result As New List(Of Item)
            For Each item As NetOffice.COMObject In items
                Result.Add(New Item(dir.OutlookApp, item, dir))
            Next
            Return Result.ToArray
        End Function

        Private Shared Sub ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(ByVal value As String, ByVal row As DataRow, ByVal columnName As String)
            If row.Table.Columns.Contains(columnName) = False Then
                row.Table.Columns.Add(columnName, GetType(String))
            End If
            row(columnName) = value
        End Sub

        Public Function ItemsRangeAsDataTable(startIndex As Integer, length As Integer) As DataTable
            Return ItemsAsDataTable(Me.ItemsRange(startIndex, length), False)
        End Function

        Public Function ItemsRangeAsDataTable(startIndex As Integer, length As Integer, includeExtendedItemProperties As Boolean) As DataTable
            Return ItemsAsDataTable(Me.ItemsRange(startIndex, length), includeExtendedItemProperties)
        End Function

        Public Function ItemsAllAsDataTable() As System.Data.DataTable
            Return ItemsAsDataTable(Me.ItemsAll, False)
        End Function

        Public Function ItemsAllAsDataTable(includeExtendedItemProperties As Boolean) As System.Data.DataTable
            Return ItemsAsDataTable(Me.ItemsAll, includeExtendedItemProperties)
        End Function

        ''' <summary>
        ''' List available items of a folder
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ItemsAsDataTable(items As Item(), includeExtendedItemProperties As Boolean) As System.Data.DataTable
            Dim Result As New System.Data.DataTable("Items")

            'Prepare table structure
            Result.Columns.Add("BCC", GetType(String))
            Result.Columns.Add("Body", GetType(String))
            Result.Columns.Add("BodyFormat", GetType(Enums.OlBodyFormat))
            Result.Columns.Add("CC", GetType(String))
            Result.Columns.Add("CreationTime", GetType(Date))
            Result.Columns.Add("EntryID", GetType(String))
            Result.Columns.Add("HTMLBody", GetType(String))
            Result.Columns.Add("Importance", GetType(Enums.OlImportance))
            Result.Columns.Add("IsAppointment", GetType(Boolean))
            Result.Columns.Add("LastModificationTime", GetType(Date))
            Result.Columns.Add("ObjectClassName", GetType(String))
            Result.Columns.Add("ParentFolderID", GetType(String))
            Result.Columns.Add("ReceivedByEntryID", GetType(String))
            Result.Columns.Add("ReceivedByName", GetType(String))
            Result.Columns.Add("ReceivedTime", GetType(Date))
            Result.Columns.Add("RTFBody", GetType(Object))
            Result.Columns.Add("SenderEmailAddress", GetType(String))
            Result.Columns.Add("SenderEmailType", GetType(String))
            Result.Columns.Add("SenderName", GetType(String))
            Result.Columns.Add("Sensitivity", GetType(Enums.OlSensitivity))
            Result.Columns.Add("SentOn", GetType(Date))
            Result.Columns.Add("Subject", GetType(String))
            Result.Columns.Add("TaskSubject", GetType(String))
            Result.Columns.Add("To", GetType(String))
            Result.Columns.Add("UnRead", GetType(Boolean))
            Result.Columns.Add("Recipients", GetType(NetOffice.OutlookApi.Recipient()))
            Result.Columns.Add("Start", GetType(Date))
            Result.Columns.Add("End", GetType(Date))
            Result.Columns.Add("StartUtc", GetType(Date))
            Result.Columns.Add("EndUtc", GetType(Date))
            Result.Columns.Add("Categories", GetType(String))
            Result.Columns.Add("Location", GetType(String))
            Result.Columns.Add("Organizer", GetType(String))
            Result.Columns.Add("Duration", GetType(String))
            Result.Columns.Add("BusyStatus", GetType(String))
            Result.Columns.Add("RequiredAttendees", GetType(String))
            Result.Columns.Add("ReminderMinutesBeforeStart", GetType(String))
            Result.Columns.Add("MessageClass", GetType(String))

            'Fill items into table
            For ItemCounter As Integer = 0 To items.Length - 1
                Dim NewRow As System.Data.DataRow = Result.NewRow
                NewRow("BCC") = items(ItemCounter).BCC
                NewRow("Body") = items(ItemCounter).Body
                NewRow("BodyFormat") = items(ItemCounter).BodyFormat
                NewRow("CC") = items(ItemCounter).CC
                NewRow("CreationTime") = items(ItemCounter).CreationTime
                NewRow("EntryID") = items(ItemCounter).EntryID
                NewRow("HTMLBody") = items(ItemCounter).HTMLBody
                NewRow("Importance") = items(ItemCounter).Importance
                NewRow("IsAppointment") = items(ItemCounter).IsAppointment
                NewRow("LastModificationTime") = items(ItemCounter).LastModificationTime
                NewRow("ObjectClassName") = items(ItemCounter).ObjectClassName
                NewRow("ParentFolderID") = items(ItemCounter).ParentFolderID
                NewRow("ReceivedByEntryID") = items(ItemCounter).ReceivedByEntryID
                NewRow("ReceivedByName") = items(ItemCounter).ReceivedByName
                NewRow("ReceivedTime") = items(ItemCounter).ReceivedTime
                NewRow("RTFBody") = items(ItemCounter).RTFBody
                NewRow("SenderEmailAddress") = items(ItemCounter).SenderEmailAddress
                NewRow("SenderEmailType") = items(ItemCounter).SenderEmailType
                NewRow("SenderName") = items(ItemCounter).SenderName
                NewRow("Sensitivity") = items(ItemCounter).Sensitivity
                NewRow("SentOn") = items(ItemCounter).SentOn
                NewRow("Subject") = items(ItemCounter).Subject
                NewRow("TaskSubject") = items(ItemCounter).TaskSubject
                NewRow("To") = items(ItemCounter).To
                NewRow("UnRead") = items(ItemCounter).UnRead
                NewRow("Start") = items(ItemCounter).Start
                NewRow("End") = items(ItemCounter).End
                NewRow("StartUtc") = items(ItemCounter).StartUtc
                NewRow("EndUtc") = items(ItemCounter).EndUtc
                NewRow("Categories") = items(ItemCounter).Categories
                NewRow("Location") = items(ItemCounter).Location
                NewRow("Organizer") = items(ItemCounter).Organizer
                NewRow("Duration") = items(ItemCounter).Duration
                NewRow("BusyStatus") = items(ItemCounter).BusyStatus
                NewRow("RequiredAttendees") = items(ItemCounter).RequiredAttendees
                NewRow("ReminderMinutesBeforeStart") = items(ItemCounter).ReminderMinutesBeforeStart
                NewRow("MessageClass") = items(ItemCounter).MessageClass
                If items(ItemCounter).Recipients IsNot Nothing Then
                    Dim Recipients As New List(Of NetOffice.OutlookApi.Recipient)
                    For RecipientsCounter As Integer = 1 To items(ItemCounter).Recipients.Count
                        Recipients.Add(items(ItemCounter).Recipients(RecipientsCounter))
                    Next
                    NewRow("Recipients") = Recipients.ToArray
                End If

                If includeExtendedItemProperties Then
                    For Each ExtendedPropertyName As String In items(ItemCounter).ItemPropertyNames
                        Dim ExtendedPropertyValue As Object = items(ItemCounter).ItemPropertyValues(ExtendedPropertyName)
                        If ExtendedPropertyValue IsNot Nothing Then
                            If NewRow.Table.Columns.Contains(ExtendedPropertyName) = False Then
                                NewRow.Table.Columns.Add(ExtendedPropertyName, ExtendedPropertyValue.GetType)
                                NewRow(ExtendedPropertyName) = ExtendedPropertyValue
                            Else
                                If IsDBNull(NewRow(ExtendedPropertyName)) Then 'never override major fields with extended property data
                                    NewRow(ExtendedPropertyName) = ExtendedPropertyValue
                                End If
                            End If
                        End If
                    Next
                End If

                Result.Rows.Add(NewRow)
            Next

            Return Result
        End Function

        '''' <summary>
        '''' List available items of a folder
        '''' </summary>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function ItemsAsDataTable(items As Item()) As System.Data.DataTable
        'Dim Result As New DataTable("items")
        'Dim ProcessedSchemas As New ArrayList
        'Dim Columns As New Hashtable
        ''Add all items into the result table with all of their properties as complete as possible
        'For Each MyItem As Item In items
        '    'Add required additional columns if not yet done
        '    If ProcessedSchemas.Contains(MyItem.ExchangeItem.Schema) = False Then
        '        For Each prop As NetOffice.OutlookApi.PropertyAccessor In MyItem.ExchangeItem.Schema
        '            Dim ColName As String = prop.Name
        '            If prop.Version <> 0 Then ColName &= "_V" & prop.Version
        '            If Not Result.Columns.Contains(ColName) Then
        '                If prop.Type.ToString.StartsWith("System.Nullable") Then
        '                    'Dataset doesn't support System.Nullable --> use System.Object
        '                    Columns.Add(ColName, New FolderItemPropertyToColumn(prop, Result.Columns.Add(ColName, GetType(Object))))
        '                Else
        '                    'Use the property type as regular
        '                    Columns.Add(ColName, New FolderItemPropertyToColumn(prop, Result.Columns.Add(ColName, prop.Type)))
        '                End If
        '            End If
        '        Next
        '    End If
        '    'Add item as new data row
        '    Dim row As System.Data.DataRow = Result.NewRow
        '    For Each key As Object In Columns.Keys
        '        Dim MyColumn As FolderItemPropertyToColumn = CType(Columns(key), FolderItemPropertyToColumn)
        '        If Not MyColumn.SchemaProperty Is Nothing Then
        '            Try
        '                If MyItem.ExchangeItem.Item(MyColumn.SchemaProperty) Is Nothing Then
        '                    row(MyColumn.Column) = DBNull.Value
        '                Else
        '                    Select Case MyItem.ExchangeItem.Item(MyColumn.SchemaProperty).GetType.ToString
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection
        '                            value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection)
        '                            'Dim comp As String = MyItem.Item(CType(Columns("Subject"), FolderItemPropertyToColumn).SchemaProperty).ToString
        '                            'If comp.IndexOf("Wezel") > -1 Then
        '                            '    Debug.Print(value.ToString)
        '                            'End If
        '                            'For Each valueKey As ExtendedProperty In value
        '                            '    Debug.Print(valueKey.PropertyDefinition.Name & "=" & valueKey.Value.ToString)
        '                            'Next
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.CompleteName).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.CompleteName
        '                            value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.CompleteName)
        '                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value.Title, row, MyColumn.Column.ColumnName & "_Title")
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.EmailAddressDictionary).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.EmailAddressDictionary
        '                            value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.EmailAddressDictionary)
        '                            If value.Contains(EmailAddressKey.EmailAddress1) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(EmailAddressKey.EmailAddress1).Address, row, MyColumn.Column.ColumnName & "_Email1")
        '                            If value.Contains(EmailAddressKey.EmailAddress2) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(EmailAddressKey.EmailAddress2).Address, row, MyColumn.Column.ColumnName & "_Email2")
        '                            If value.Contains(EmailAddressKey.EmailAddress3) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(EmailAddressKey.EmailAddress3).Address, row, MyColumn.Column.ColumnName & "_Email3")
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary
        '                            value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary)
        '                            If value.Contains(PhysicalAddressKey.Business) Then
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).Street, row, MyColumn.Column.ColumnName & "_Business_Street")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).PostalCode, row, MyColumn.Column.ColumnName & "_Business_PostalCode")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).City, row, MyColumn.Column.ColumnName & "_Business_City")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).State, row, MyColumn.Column.ColumnName & "_Business_State")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).CountryOrRegion, row, MyColumn.Column.ColumnName & "_Business_CountryOrRegion")
        '                            End If
        '                            If value.Contains(PhysicalAddressKey.Home) Then
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).Street, row, MyColumn.Column.ColumnName & "_Home_Street")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).PostalCode, row, MyColumn.Column.ColumnName & "_Home_PostalCode")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).City, row, MyColumn.Column.ColumnName & "_Home_City")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).State, row, MyColumn.Column.ColumnName & "_Home_State")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).CountryOrRegion, row, MyColumn.Column.ColumnName & "_Home_CountryOrRegion")
        '                            End If
        '                            If value.Contains(PhysicalAddressKey.Other) Then
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).Street, row, MyColumn.Column.ColumnName & "_Other_Street")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).PostalCode, row, MyColumn.Column.ColumnName & "_Other_PostalCode")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).City, row, MyColumn.Column.ColumnName & "_Other_City")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).State, row, MyColumn.Column.ColumnName & "_Other_State")
        '                                ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).CountryOrRegion, row, MyColumn.Column.ColumnName & "_Other_CountryOrRegion")
        '                            End If
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary
        '                            value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary)
        '                            If value.Contains(PhoneNumberKey.BusinessPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.BusinessPhone), row, MyColumn.Column.ColumnName & "_BusinessPhone")
        '                            If value.Contains(PhoneNumberKey.BusinessPhone2) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.BusinessPhone2), row, MyColumn.Column.ColumnName & "_BusinessPhone2")
        '                            If value.Contains(PhoneNumberKey.BusinessFax) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.BusinessFax), row, MyColumn.Column.ColumnName & "_BusinessFax")
        '                            If value.Contains(PhoneNumberKey.CompanyMainPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.CompanyMainPhone), row, MyColumn.Column.ColumnName & "_CompanyMainPhone")
        '                            If value.Contains(PhoneNumberKey.CarPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.CarPhone), row, MyColumn.Column.ColumnName & "_CarPhone")
        '                            If value.Contains(PhoneNumberKey.Callback) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.Callback), row, MyColumn.Column.ColumnName & "_Callback")
        '                            If value.Contains(PhoneNumberKey.AssistantPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.AssistantPhone), row, MyColumn.Column.ColumnName & "_AssistantPhone")
        '                            If value.Contains(PhoneNumberKey.HomeFax) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.HomeFax), row, MyColumn.Column.ColumnName & "_HomeFax")
        '                            If value.Contains(PhoneNumberKey.HomePhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.HomePhone), row, MyColumn.Column.ColumnName & "_HomePhone")
        '                            If value.Contains(PhoneNumberKey.HomePhone2) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.HomePhone2), row, MyColumn.Column.ColumnName & "_HomePhone2")
        '                            If value.Contains(PhoneNumberKey.MobilePhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.MobilePhone), row, MyColumn.Column.ColumnName & "_MobilePhone")
        '                            If value.Contains(PhoneNumberKey.OtherFax) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.OtherFax), row, MyColumn.Column.ColumnName & "_OtherFax")
        '                            If value.Contains(PhoneNumberKey.OtherTelephone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.OtherTelephone), row, MyColumn.Column.ColumnName & "_OtherTelephone")
        '                            If value.Contains(PhoneNumberKey.PrimaryPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.PrimaryPhone), row, MyColumn.Column.ColumnName & "_PrimaryPhone")
        '                            If value.Contains(PhoneNumberKey.RadioPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.RadioPhone), row, MyColumn.Column.ColumnName & "_RadioPhone")
        '                        Case Else
        '                            row(MyColumn.Column) = MyItem.ExchangeItem.Item(MyColumn.SchemaProperty)
        '                    End Select
        '                End If
        '            Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
        '                'Mark this column to be killed at the end because it only contains non-sense
        '                MyColumn.SchemaProperty = Nothing
        '            Catch ex As Microsoft.Exchange.WebServices.Data.ServiceVersionException
        '                'Mark this column to be killed at the end because it only contains non-sense
        '                MyColumn.SchemaProperty = Nothing
        '            End Try
        '        End If
        '    Next
        '    Result.Rows.Add(row)
        'Next
        ''Remove all columns which are marked to be deleted
        'For Each key As Object In Columns.Keys
        '    If CType(Columns(key), FolderItemPropertyToColumn).SchemaProperty Is Nothing Then
        '        'Missing data indicates a column to be deleted
        '        Result.Columns.Remove(CType(Columns(key), FolderItemPropertyToColumn).Column)
        '    End If
        'Next
        'Result.Columns("ID").Unique = True
        'Return Result
        'End Function

        ''' <summary>
        ''' Schema information to a column
        ''' </summary>
        ''' <remarks></remarks>
        Private Class FolderItemPropertyToColumn
            Public Sub New(ByVal schemaProperty As NetOffice.OutlookApi.PropertyAccessor, ByVal column As System.Data.DataColumn) ' PropertyDefinition
                Me.Column = column
                Me.SchemaProperty = schemaProperty
            End Sub
            Public Column As System.Data.DataColumn
            Public SchemaProperty As NetOffice.OutlookApi.PropertyAccessor
        End Class

        '''' <summary>
        '''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        '''' </summary>
        '''' <returns></returns>
        'Public Function ItemsAsExchangeItem() As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item)
        '    Return Me.OutlookFolder.FindItems(New ItemView(Integer.MaxValue)).Items
        'End Function

        '''' <summary>
        '''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        '''' </summary>
        '''' <returns></returns>
        'Public Function Items() As Item()
        '    Dim Result As New List(Of Item)
        '    For Each ExchangeItem As Microsoft.Exchange.WebServices.Data.Item In ItemsAsExchangeItem()
        '        Result.Add(New Item(Me._outlookWrapper, ExchangeItem, Me))
        '    Next
        '    Return Result.ToArray
        'End Function

        '''' <summary>
        '''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        '''' </summary>
        '''' <returns></returns>
        'Public Function Items(searchFilter As Microsoft.Exchange.WebServices.Data.SearchFilter, itemView As Microsoft.Exchange.WebServices.Data.ItemView) As Item()
        '    Dim Result As New List(Of Item)
        '    For Each ExchangeItem As Microsoft.Exchange.WebServices.Data.Item In ItemsAsExchangeItem(searchFilter, itemView)
        '        Result.Add(New Item(Me._outlookWrapper, ExchangeItem, Me))
        '    Next
        '    Return Result.ToArray
        'End Function

        '''' <summary>
        '''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        '''' </summary>
        '''' <returns></returns>
        'Public Function MailboxItems(searchFilter As Microsoft.Exchange.WebServices.Data.SearchFilter, itemView As Microsoft.Exchange.WebServices.Data.ItemView) As Item()
        '    Dim searchFolder As Directory = Me.InitialRootDirectory.SelectSubFolder("AllItems", False, Me._outlookWrapper.DirectorySeparatorChar)
        '    Dim Result As New List(Of Item)
        '    For Each ExchangeItem As Microsoft.Exchange.WebServices.Data.Item In searchFolder.ItemsAsExchangeItem(searchFilter, itemView)
        '        Result.Add(New Item(Me._outlookWrapper, ExchangeItem, Me))
        '    Next
        '    Return Result.ToArray
        'End Function

        ''' Total amount of items of a folder
        Public Function ItemCount() As Integer
            Return Me.OutlookFolder.Items.Count
        End Function

        ''' <summary>
        ''' Number of unread items of a folder
        ''' </summary>
        ''' <returns></returns>
        Public Function ItemUnreadCount() As Integer
            Return Me.OutlookFolder.UnReadItemCount
        End Function

        Private _ParentFolderID As String
        ''' <summary>
        ''' The unique ID of the parent folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ParentFolderID() As String
            Get
                Return Me.ParentDirectory.FolderID
            End Get
        End Property

        Private _FolderID As String
        ''' <summary>
        ''' The unique ID of the folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FolderID() As String
            Get
                If _FolderID Is Nothing Then
                    _FolderID = _folder.EntryID '.Id.UniqueId
                End If
                Return _FolderID
            End Get
        End Property

        Private _FolderClass As String
        ''' <summary>
        ''' The folder class name
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FolderClass() As String
            Get
                If _FolderClass Is Nothing Then
                    _FolderClass = _folder.Class.ToString '.FolderClass
                End If
                Return _FolderClass
            End Get
            Set(ByVal value As String)
                _FolderClass = value
                '_folder.FolderClass = value
                '_folder.Class = Enums.OlObjectClass.olItems ' value
                Me.CachedFolderDisplayPath = Nothing
            End Set
        End Property

        Private _DisplayName As String
        ''' <summary>
        ''' The display name of the folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DisplayName() As String
            Get
                If _DisplayName Is Nothing Then
                    _DisplayName = System.IO.Path.GetFileName(_folder.FolderPath) '.DisplayName
                End If
                Return _DisplayName
            End Get
            Set(ByVal value As String)
                _DisplayName = value
                Me.CachedFolderDisplayPath = Nothing
            End Set
        End Property

        Private CachedFolderDisplayPath As String = Nothing
        ''' <summary>
        ''' The display path of the folder separated by back-slashes (\)
        ''' </summary>
        ''' <returns>Existing back-slashes in folder's display names might confuse here - in case that back-slahes are possible in display names</returns>
        Public ReadOnly Property DisplayPath As String
            Get
                If Me.ParentDirectory IsNot Nothing Then
                    Return Me.ParentDirectory.DisplayPath & "\" & Me.DisplayName
                Else
                    Return Me.DisplayName
                End If
            End Get
        End Property

        Private CachedFolderPath As String = Nothing
        ''' <summary>
        ''' The path of the folder separated by back-slashes (\)
        ''' </summary>
        ''' <returns>Existing back-slashes in folder's display names might confuse here - in case that back-slahes are possible in display names</returns>
        Public ReadOnly Property Path As String
            Get
                If Me.ParentDirectory IsNot Nothing Then
                    If Me.ParentDirectory.Path = Nothing Then
                        Return Me.DisplayName
                    Else
                        Return Me.ParentDirectory.Path & "\" & Me.DisplayName
                    End If
                Else
                    Return "" 'This is the root folder
                End If
            End Get
        End Property

        ''' <summary>
        ''' The relative folder path starting from your initial top directory
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>Parent folder structure can't be looked up till root folder, that's why it's only up to your initial top directory</remarks>
        Public ReadOnly Property ParentDirectory As Directory
            Get
                If _parentDirectory Is Nothing AndAlso _parentFolder IsNot Nothing Then
                    _parentDirectory = New Directory(_outlookWrapper, _parentFolder)
                ElseIf _parentDirectory Is Nothing AndAlso _parentFolder Is Nothing Then
                    '_parentFolder = New Directory(Me.ExchangeFolder.ParentFolderId.UniqueId)
                End If
                Return _parentDirectory
            End Get
        End Property

        '''' <summary>
        '''' The default view for folders
        '''' </summary>
        '''' <returns></returns>
        'Friend Shared Function DefaultFolderView(folderTraversal As FolderTraversal, offSet As Integer) As FolderView
        '    Dim Result As New FolderView(Integer.MaxValue, offSet)
        '    Result.PropertySet = DefaultPropertySet()
        '    Result.Traversal = folderTraversal
        '    Return Result
        'End Function

        'Friend Shared Function DefaultPropertySet() As PropertySet
        '    Dim AdditionalProperties As New List(Of Microsoft.Exchange.WebServices.Data.PropertyDefinition)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.ChildFolderCount)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.TotalCount)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.UnreadCount)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.FolderClass)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.Id)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.ParentFolderId)
        '    AdditionalProperties.Add(Microsoft.Exchange.WebServices.Data.FolderSchema.DisplayName)
        '    Return New PropertySet(BasePropertySet.FirstClassProperties, AdditionalProperties.ToArray)
        'End Function

        Private _SubFolders As List(Of Directory)
        Public ReadOnly Property SubFolders As Directory()
            Get
                If _SubFolders Is Nothing Then
                    'fill from hierarchy list
                    _SubFolders = New List(Of Directory)
                    For Each folder As MAPIFolder In Me.OutlookFolder.Folders
                        Dim childDir As New Directory(_outlookWrapper, folder)
                        childDir._Internal_SetParentDirectory(Me)
                        _SubFolders.Add(childDir)
                    Next
                End If
                Return _SubFolders.ToArray
            End Get
        End Property

        Public Function LookupSubDirectory(searchedFolderID As String) As Directory
            For Each SubFolder As Directory In Me.SubFolders
                If searchedFolderID = SubFolder.FolderID Then Return SubFolder
                If SubFolder.SubFolderCount > 0 Then
                    Dim FoundResult As Directory = SubFolder.LookupSubDirectory(searchedFolderID)
                    If FoundResult IsNot Nothing Then
                        Return FoundResult
                    End If
                End If
            Next
            Return Nothing
        End Function

        Private Sub _Internal_SetParentDirectory(parentDirectory As Directory)
            Me._parentDirectory = parentDirectory
            Me._parentFolder = parentDirectory.OutlookFolder
            Me._ParentFolderID = parentDirectory.FolderID
        End Sub

        Public ReadOnly Property InitialRootDirectory As Directory
            Get
                If _parentDirectory IsNot Nothing Then
                    Return _parentDirectory.InitialRootDirectory
                Else
                    Return Me
                End If
            End Get
        End Property

        ''' <summary>
        ''' Lookup a directory based on its directory structure
        ''' </summary>
        ''' <param name="subfolder">A string containing the relative folder path, e.g. &quot;Inbox\Done&quot;</param>
        ''' <param name="searchCaseInsensitive">Ignore upper/lower case differences</param>
        ''' <returns></returns>
        Public Function SelectSubFolder(subFolder As String, ByVal searchCaseInsensitive As Boolean) As Directory
            Return Me.SelectSubFolder(subFolder, searchCaseInsensitive, False, Me._outlookWrapper.DirectorySeparatorChar)
        End Function

        ''' <summary>
        ''' Lookup a directory based on its directory structure
        ''' </summary>
        ''' <param name="subfolder">A string containing the relative folder path, e.g. &quot;Inbox\Done&quot;</param>
        ''' <param name="searchCaseInsensitive">Ignore upper/lower case differences</param>
        ''' <param name="autoCreateFolder">Create folder if not yet existing</param>
        ''' <returns></returns>
        Public Function SelectSubFolder(subFolder As String, ByVal searchCaseInsensitive As Boolean, autoCreateFolder As Boolean) As Directory
            Return Me.SelectSubFolder(subFolder, searchCaseInsensitive, autoCreateFolder, Me._outlookWrapper.DirectorySeparatorChar)
        End Function

        ''' <summary>
        ''' Lookup a directory based on its directory structure
        ''' </summary>
        ''' <param name="subfolder">A string containing the relative folder path, e.g. &quot;Inbox\Done&quot;</param>
        ''' <param name="searchCaseInsensitive">Ignore upper/lower case differences</param>
        ''' <param name="autoCreateFolder">Create folder if not yet existing</param>
        ''' <param name="directorySeparatorChar"></param>
        ''' <returns></returns>
        Private Function SelectSubFolder(subFolder As String, ByVal searchCaseInsensitive As Boolean, autoCreateFolder As Boolean, directorySeparatorChar As Char) As Directory
            If subFolder = Nothing Then
                Return Me
            ElseIf subFolder.StartsWith(directorySeparatorChar) Then
                Throw New ArgumentException("subFolder can't start with a directorySeparatorChar", "subFolder")
            Else
                Dim subfoldersSplitted As String() = subFolder.Split(directorySeparatorChar)
                Dim nextSubFolder As String = subfoldersSplitted(0)
                For Each mySubfolder As Directory In Me.SubFolders
                    If mySubfolder.DisplayName = nextSubFolder OrElse (searchCaseInsensitive AndAlso mySubfolder.DisplayName.ToLowerInvariant = nextSubFolder.ToLowerInvariant) Then
                        If subfoldersSplitted.Length > 1 Then
                            'recursive call required
                            Return mySubfolder.SelectSubFolder(String.Join(directorySeparatorChar, subfoldersSplitted, 1, subfoldersSplitted.Length - 1), searchCaseInsensitive, autoCreateFolder, directorySeparatorChar)
                        Else
                            'this is the last recursion - just return our current path item
                            Return mySubfolder
                        End If
                    End If
                Next
                If autoCreateFolder = True Then
                    Dim NewFolderName As String = subfoldersSplitted(0)
                    Me.CreateSubFolder(NewFolderName)
                    For Each mySubfolder As Directory In Me.SubFolders
                        If mySubfolder.DisplayName = nextSubFolder OrElse (searchCaseInsensitive AndAlso mySubfolder.DisplayName.ToLowerInvariant = nextSubFolder.ToLowerInvariant) Then
                            If subfoldersSplitted.Length > 1 Then
                                'recursive call required
                                Return mySubfolder.SelectSubFolder(String.Join(directorySeparatorChar, subfoldersSplitted, 1, subfoldersSplitted.Length - 1), searchCaseInsensitive, autoCreateFolder, directorySeparatorChar)
                            Else
                                'this is the last recursion - just return our current path item
                                Return mySubfolder
                            End If
                        End If
                    Next
                    Throw New System.Exception("Folder """ & NewFolderName & """ couldn't be created in " & Me.DisplayPath)
                Else
                    Throw New System.Exception("Folder """ & subFolder & """ hasn't been found in " & Me.DisplayPath)
                End If
            End If
        End Function

        Public Sub CreateSubFolder(newFolderName As String)
            Me.OutlookFolder.Folders.Add(newFolderName)
            _SubFolders = Nothing 'reset subfolders cache
        End Sub

        ''' <summary>
        ''' The number of child folders
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>For performance reasons, SubFolderCount doesn't load the full subfolder data in case it hasn't been loaded yet</remarks>
        Public ReadOnly Property SubFolderCount() As Integer
            Get
                If _SubFolders Is Nothing Then
                    Return _folder.Folders.Count
                Else
                    Return _SubFolders.Count
                End If
            End Get
        End Property

        '''' <summary>
        '''' Save changes to this folder
        '''' </summary>
        '''' <remarks></remarks>
        'Public Sub Save()
        '    _folder.Update()
        'End Sub

        '''' <summary>
        '''' Save this folder as sub folder of the specified one
        '''' </summary>
        '''' <param name="parentFolder"></param>
        '''' <remarks></remarks>
        'Public Sub Save(ByVal parentFolder As Directory)
        '    _folder.Save(New FolderId(parentFolder.ID))
        'End Sub

        Public Overrides Function ToString() As String
            Return Me.DisplayPath
        End Function

        Private _ExtendedData As Generic.Dictionary(Of String, Object)
        'Public Function ExtendedData() As Generic.Dictionary(Of String, Object)
        '    If _ExtendedData Is Nothing Then
        '        'Load first class props
        '        'Dim propSet As New Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.BasePropertySet.FirstClassProperties)
        '        'Microsoft.Exchange.WebServices.Data.EmailMessage.Bind(_service.CreateConfiguredExchangeService, _exchangeItem.Id, propSet)
        '        _ExtendedData = New Generic.Dictionary(Of String, Object)
        '        'Add all items into the result table with all of their properties as complete as possible
        '        'Add required additional columns if not yet done
        '        For Each prop As PropertyDefinition In Me.OutlookFolder.Schema
        '            Dim ColName As String = prop.Name
        '            If prop.Version <> 0 Then ColName &= "_V" & prop.Version
        '            Try
        '                If Me.OutlookFolder.Item(prop) Is Nothing Then
        '                    _ExtendedData.Add(ColName, Nothing)
        '                Else
        '                    Select Case Me.OutlookFolder.Item(prop).GetType.ToString
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection
        '                            value = CType(Me.OutlookFolder.Item(prop), Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection)
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.CompleteName).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.CompleteName
        '                            value = CType(Me.OutlookFolder.Item(prop), Microsoft.Exchange.WebServices.Data.CompleteName)
        '                            _ExtendedData.Add(ColName & "_Title", value.Title)
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.EmailAddressDictionary).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.EmailAddressDictionary
        '                            value = CType(Me.OutlookFolder.Item(prop), Microsoft.Exchange.WebServices.Data.EmailAddressDictionary)
        '                            If value.Contains(EmailAddressKey.EmailAddress1) Then _ExtendedData.Add(ColName & "_Email1", value(EmailAddressKey.EmailAddress1).Address)
        '                            If value.Contains(EmailAddressKey.EmailAddress2) Then _ExtendedData.Add(ColName & "_Email2", value(EmailAddressKey.EmailAddress2).Address)
        '                            If value.Contains(EmailAddressKey.EmailAddress3) Then _ExtendedData.Add(ColName & "_Email3", value(EmailAddressKey.EmailAddress3).Address)
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary
        '                            value = CType(Me.OutlookFolder.Item(prop), Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary)
        '                            If value.Contains(PhysicalAddressKey.Business) Then
        '                                _ExtendedData.Add(ColName & "_Business_Street", value(PhysicalAddressKey.Business).Street)
        '                                _ExtendedData.Add(ColName & "_Business_PostalCode", value(PhysicalAddressKey.Business).PostalCode)
        '                                _ExtendedData.Add(ColName & "_Business_City", value(PhysicalAddressKey.Business).City)
        '                                _ExtendedData.Add(ColName & "_Business_State", value(PhysicalAddressKey.Business).State)
        '                                _ExtendedData.Add(ColName & "_Business_CountryOrRegion", value(PhysicalAddressKey.Business).CountryOrRegion)
        '                            End If
        '                            If value.Contains(PhysicalAddressKey.Home) Then
        '                                _ExtendedData.Add(ColName & "_Home_Street", value(PhysicalAddressKey.Home).Street)
        '                                _ExtendedData.Add(ColName & "_Home_PostalCode", value(PhysicalAddressKey.Home).PostalCode)
        '                                _ExtendedData.Add(ColName & "_Home_City", value(PhysicalAddressKey.Home).City)
        '                                _ExtendedData.Add(ColName & "_Home_State", value(PhysicalAddressKey.Home).State)
        '                                _ExtendedData.Add(ColName & "_Home_CountryOrRegion", value(PhysicalAddressKey.Home).CountryOrRegion)
        '                            End If
        '                            If value.Contains(PhysicalAddressKey.Other) Then
        '                                _ExtendedData.Add(ColName & "_Other_Street", value(PhysicalAddressKey.Other).Street)
        '                                _ExtendedData.Add(ColName & "_Other_PostalCode", value(PhysicalAddressKey.Other).PostalCode)
        '                                _ExtendedData.Add(ColName & "_Other_City", value(PhysicalAddressKey.Other).City)
        '                                _ExtendedData.Add(ColName & "_Other_State", value(PhysicalAddressKey.Other).State)
        '                                _ExtendedData.Add(ColName & "_Other_CountryOrRegion", value(PhysicalAddressKey.Other).CountryOrRegion)
        '                            End If
        '                        Case GetType(Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary).ToString
        '                            Dim value As Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary
        '                            value = CType(Me.OutlookFolder.Item(prop), Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary)
        '                            If value.Contains(PhoneNumberKey.BusinessPhone) Then _ExtendedData.Add(ColName & "_BusinessPhone", value(PhoneNumberKey.BusinessPhone))
        '                            If value.Contains(PhoneNumberKey.BusinessPhone2) Then _ExtendedData.Add(ColName & "_BusinessPhone2", value(PhoneNumberKey.BusinessPhone2))
        '                            If value.Contains(PhoneNumberKey.BusinessFax) Then _ExtendedData.Add(ColName & "_BusinessFax", value(PhoneNumberKey.BusinessFax))
        '                            If value.Contains(PhoneNumberKey.CompanyMainPhone) Then _ExtendedData.Add(ColName & "_CompanyMainPhone", value(PhoneNumberKey.CompanyMainPhone))
        '                            If value.Contains(PhoneNumberKey.CarPhone) Then _ExtendedData.Add(ColName & "_CarPhone", value(PhoneNumberKey.CarPhone))
        '                            If value.Contains(PhoneNumberKey.Callback) Then _ExtendedData.Add(ColName & "_Callback", value(PhoneNumberKey.Callback))
        '                            If value.Contains(PhoneNumberKey.AssistantPhone) Then _ExtendedData.Add(ColName & "_AssistantPhone", value(PhoneNumberKey.AssistantPhone))
        '                            If value.Contains(PhoneNumberKey.HomeFax) Then _ExtendedData.Add(ColName & "_HomeFax", value(PhoneNumberKey.HomeFax))
        '                            If value.Contains(PhoneNumberKey.HomePhone) Then _ExtendedData.Add(ColName & "_HomePhone", value(PhoneNumberKey.HomePhone))
        '                            If value.Contains(PhoneNumberKey.HomePhone2) Then _ExtendedData.Add(ColName & "_HomePhone2", value(PhoneNumberKey.HomePhone2))
        '                            If value.Contains(PhoneNumberKey.MobilePhone) Then _ExtendedData.Add(ColName & "_MobilePhone", value(PhoneNumberKey.MobilePhone))
        '                            If value.Contains(PhoneNumberKey.OtherFax) Then _ExtendedData.Add(ColName & "_OtherFax", value(PhoneNumberKey.OtherFax))
        '                            If value.Contains(PhoneNumberKey.OtherTelephone) Then _ExtendedData.Add(ColName & "_OtherTelephone", value(PhoneNumberKey.OtherTelephone))
        '                            If value.Contains(PhoneNumberKey.PrimaryPhone) Then _ExtendedData.Add(ColName & "_PrimaryPhone", value(PhoneNumberKey.PrimaryPhone))
        '                            If value.Contains(PhoneNumberKey.RadioPhone) Then _ExtendedData.Add(ColName & "_RadioPhone", value(PhoneNumberKey.RadioPhone))
        '                        Case Else
        '                            _ExtendedData.Add(ColName, Me.OutlookFolder.Item(prop))
        '                    End Select
        '                End If
        '            Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
        '                'Mark this column to be killed at the end because it only contains non-sense
        '            Catch ex As Microsoft.Exchange.WebServices.Data.ServiceVersionException
        '                'Mark this column to be killed at the end because it only contains non-sense
        '            Catch ex As NullReferenceException
        '                _ExtendedData.Add(ColName, Nothing)
        '            End Try
        '        Next
        '    End If
        '    Return _ExtendedData
        'End Function

        Public Delegate Sub DirectoryAction(dir As Directory)

        Public Sub ForDirectoryAndEachSubDirectory(actions As DirectoryAction)
            ForDirectoryAndEachSubDirectory(Me, actions)
        End Sub

        Public Shared Sub ForDirectoryAndEachSubDirectory(dir As Directory, actions As DirectoryAction)
            actions(dir)
            For Each dirItem As Directory In dir.SubFolders
                ForDirectoryAndEachSubDirectory(dirItem, actions)
            Next
        End Sub

        Public Sub ForEachSubDirectory(actions As DirectoryAction)
            ForEachSubDirectory(Me, actions)
        End Sub

        Public Shared Sub ForEachSubDirectory(dir As Directory, actions As DirectoryAction)
            For Each dirItem As Directory In dir.SubFolders
                actions(dir)
                ForEachSubDirectory(dirItem, actions)
            Next
        End Sub

    End Class

End Namespace