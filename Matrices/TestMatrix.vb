Public Class TestMatrix

    Public Shared Sub Main()

        System.Windows.Forms.Application.Run(New TestMatrix)

    End Sub


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        DataGridView1.Rows.Add(1, 2, 3)
        DataGridView1.Rows.Add(4, 5, 5)
        DataGridView1.Rows.Add(8, 1, 2)

        DataGridView2.Rows.Add()
        DataGridView2.Rows.Add()
        DataGridView2.Rows.Add()


    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim iArray(3 - 1, 3 - 1) As Double

        Dim c1, c2 As Integer
        c1 = 0
        
        For Each row As System.Windows.Forms.DataGridViewRow In DataGridView1.Rows
            c2 = 0
            For Each cell As System.Windows.Forms.DataGridViewCell In row.Cells
                iArray(c1, c2) = cell.Value
                c2 += 1
            Next
            c1 += 1
        Next

        Dim A As New Matrix(iArray)

        Dim ASQ As Vector = Matrix.Eigenvalues(A)

        If Not ASQ Is Nothing Then

            c1 = 0
            For Each row As System.Windows.Forms.DataGridViewRow In DataGridView2.Rows
                c2 = 0

                row.Cells(0).Value = ASQ.Item(c1)
                c1 += 1
            Next

        End If




    End Sub
End Class