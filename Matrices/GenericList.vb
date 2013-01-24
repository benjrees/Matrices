Public Class GenericList(Of ItemType)
    Inherits CollectionBase

    Public Function Add(ByVal value As ItemType) _
      As Double
        Return List.Add(value)
    End Function

    Public Sub Remove(ByVal value As ItemType)
        List.Remove(value)
    End Sub

    Public ReadOnly Property Item( _
      ByVal index As Double) As ItemType
        Get
            ' The appropriate item is retrieved from 
            ' the List object and explicitly cast to 
            ' the appropriate type, and then returned.
            Return CType(List.Item(index), ItemType)
        End Get
    End Property

    Public Function ToArray() As ItemType()

        Dim retValues(List.Count - 1) As ItemType
        For i As Integer = 0 To List.Count - 1
            retValues(i) = CType(List.Item(i), ItemType)
        Next

        Return retValues
    End Function

End Class
