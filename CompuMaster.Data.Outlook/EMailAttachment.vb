Option Strict On
Option Explicit On

Namespace CompuMaster.Data.Outlook

    Public Class EMailAttachment
        Public Property FileName As String
        Public Property FilePath As String
        Public Property FileData As Byte()
        Public Property FileStream As IO.Stream

        Public Sub New(FilePath As String)
            Me.FilePath = FilePath
        End Sub

        Public Sub New(FileName As String, FileData() As Byte)
            Me.FileName = FileName
            Me.FileData = FileData
        End Sub

        Public Sub New(FileName As String, FilePath As String)
            Me.FileName = FileName
            Me.FilePath = FilePath
        End Sub

        Public Sub New(FileName As String, FileStream As IO.Stream)
            Me.FileName = FileName
            Me.FileStream = FileStream
        End Sub

    End Class

End Namespace