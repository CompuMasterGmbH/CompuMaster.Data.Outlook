Option Strict On
Option Explicit On

Namespace CompuMaster.Data.Outlook

    ''' <summary>
    ''' Represents an attendee or recipient
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Recipient
        Public Sub New(ByVal emailAddress As String)
            Me.EMailAddress = emailAddress
        End Sub
        Public Sub New(ByVal name As String, ByVal emailAddress As String)
            _Name = name
            Me.EMailAddress = emailAddress
        End Sub
        Private _Name As String
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property
        Private _EMailAddress As String
        Public Property EMailAddress() As String
            Get
                Return _EMailAddress
            End Get
            Set(ByVal value As String)
                If value = Nothing Then Throw New ArgumentNullException("value")
                _EMailAddress = value
            End Set
        End Property
    End Class

End Namespace