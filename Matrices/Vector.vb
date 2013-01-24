Public Class Vector

    Private mMat As Matrix
    Private mOrient As OrientationType


    Public Enum OrientationType
        horizontal
        vertical
    End Enum

    Public Sub New(ByVal iValues As GenericList(Of Double), ByVal iOrient As Vector.OrientationType)

        mOrient = iOrient

        If iOrient = OrientationType.horizontal Then
            Dim vals(0, iValues.Count - 1) As Double
            For j As Integer = 0 To iValues.Count - 1
                vals(0, j) = iValues.Item(j)
            Next

            mMat = New Matrix(vals)
        Else
            Dim vals(iValues.Count - 1, 0) As Double
            For i As Integer = 0 To iValues.Count - 1
                vals(i, 0) = iValues.Item(i)
            Next

            mMat = New Matrix(vals)

        End If


    End Sub

    Public Sub New(ByVal iMatrix As Matrix)




    End Sub

    Public Function Item(ByVal index As Integer) As Double
        If mOrient = OrientationType.horizontal Then
            Return mMat.Value(0, index)
        Else
            Return mMat.Value(index, 0)
        End If
    End Function


    Public Property Orientation() As OrientationType
        Get
            Return mOrient
        End Get
        Set(ByVal value As OrientationType)
            mOrient = value
        End Set
    End Property


    Public Function Count() As Integer
        If mOrient = OrientationType.horizontal Then
            Return mMat.N
        Else
            Return mMat.M
        End If
    End Function


    Public Shared Function Transpose(ByVal iV As Vector) As Vector

        Dim vals As New GenericList(Of Double)

        For i As Integer = 0 To iV.Count - 1
            vals.Add(iV.Item(i))
        Next

        If iV.Orientation = OrientationType.horizontal Then
            Return New Vector(vals, OrientationType.vertical)
        Else
            Return New Vector(vals, OrientationType.horizontal)
        End If

    End Function

    Public Shared Operator *(ByVal iA As Vector, ByVal c As Double) As Vector

        Dim result As New GenericList(Of Double)
        For i As Integer = 0 To iA.Count - 1
            result.Add(iA.Item(i) * c)
        Next

        Return New Vector(result, iA.Orientation)

    End Operator

    Public Shared Operator *(ByVal iA As Vector, ByVal iB As Vector) As Double

        Dim result As Double = 0
        For i As Integer = 0 To iA.Count - 1
            result += iA.Item(i) * iB.Item(i)
        Next

        Return result

    End Operator

    Public Shared Operator -(ByVal iA As Vector, ByVal iB As Vector) As Vector

        Dim result As New GenericList(Of Double)
        For i As Integer = 0 To iA.Count - 1
            result.Add(iA.Item(i) - iB.Item(i))
        Next

        Return New Vector(result, iA.Orientation)

    End Operator


End Class
