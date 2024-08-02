' COPYRIGHT (C) 2004, Mercury Solutions, Inc.
' MODULE:   CDasFile.vb
' AUTHOR:   Denny Forsberg
' PURPOSE:  This class implements the mapping class to hold objects by
'           a key that is of type string.
' REVISION:
'   
Imports System
Imports System.Collections.DictionaryBase

Public Class InterfaceDictionary
    Inherits DictionaryBase

    Default Public Property Item(ByVal Key As String) As [Object]
        Get
            Return CType(Dictionary(Key), Object)
        End Get
        Set(ByVal Value As [Object])
            Dictionary(Key) = Value
        End Set
    End Property

    Public ReadOnly Property Keys() As ICollection
        Get
            Return Dictionary.Keys
        End Get
    End Property

    Public ReadOnly Property Values() As ICollection
        Get
            Return Dictionary.Values
        End Get
    End Property

    Public Sub Replace(ByVal key As String, ByVal value As Object)
        If dictionary.Contains(key) Then
            dictionary.Remove(key)
        End If
        dictionary.Add(key, value)
    End Sub

    Public Sub Add(ByVal key As String, ByVal value As Object)
        Try
            Dictionary.Add(key, value)
        Catch e As Exception
            'MessageBox.Show("Source already ingested!")
        End Try
    End Sub

    Public Function Contains(ByVal key As String) As Boolean
        Return Dictionary.Contains(key)
    End Function

    Public Sub Remove(ByVal key As String)
        Try
            Dictionary.Remove(key)
        Catch ex As Exception
            'key doesn't exist
        End Try
    End Sub

    Protected Overrides Sub OnInsert(ByVal key As Object, ByVal value As Object)
        If Not key.GetType() Is Type.GetType("System.String") Then
            Throw New ArgumentException("key must be of type String.", "key")
        End If
    End Sub

    Protected Overrides Sub OnRemove(ByVal key As Object, ByVal value As Object)
        If Not key.GetType() Is Type.GetType("System.String") Then
            Throw New ArgumentException("key must be of type String.", "key")
        End If
    End Sub

    Protected Overrides Sub OnSet(ByVal key As Object, ByVal oldValue As Object, ByVal newValue As Object)
        If Not key.GetType() Is Type.GetType("System.String") Then
            Throw New ArgumentException("key must be of type String.", "key")
        End If
    End Sub

End Class