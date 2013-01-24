Imports System.Math

Public Class Matrix

    Private myValues As Double(,)     ' the actual matrix values

    Public Function Row(ByVal i As Integer) As Vector

        Dim vals As New GenericList(Of Double)
        For j As Integer = 0 To Me.N - 1
            vals.Add(Me.Value(i, j))
        Next

        Dim retVals As New Vector(vals, Vector.OrientationType.horizontal)

        Return retVals

    End Function

    Public Function Column(ByVal j As Integer) As Vector

        Dim vals As New GenericList(Of Double)
        For i As Integer = 0 To Me.M - 1
            vals.Add(Me.Value(i, j))
        Next

        Dim retVals As New Vector(vals, Vector.OrientationType.vertical)

        Return retVals

    End Function


    Public Property Values() As Double(,)
        Get
            Return myValues
        End Get
        Set(ByVal value As Double(,))
            myValues = value
        End Set
    End Property

    Public Property Value(ByVal i As Integer, ByVal j As Integer) As Double
        Get
            Return myValues(i, j)
        End Get
        Set(ByVal iValue As Double)
            myValues(i, j) = iValue
        End Set
    End Property


    Public ReadOnly Property M() As Integer
        Get
            Return UBound(myValues) + 1
        End Get
    End Property

    Public ReadOnly Property N() As Integer
        Get
            Return UBound(myValues, 2) + 1
        End Get
    End Property


    Public Sub New(ByVal iValues As Double(,))

        myValues = iValues

    End Sub

    Public Sub New(ByVal m As Integer, ByVal n As Integer)

        ReDim myValues(m - 1, n - 1)

    End Sub

    Public Sub New(ByVal iVector As Vector)

        If iVector.Orientation = Vector.OrientationType.horizontal Then
            ReDim myValues(0, iVector.Count - 1)
            For j As Integer = 0 To iVector.Count - 1
                myValues(0, j) = iVector.Item(j)
            Next
        Else
            ReDim myValues(iVector.Count - 1, 0)
            For i As Integer = 0 To iVector.Count - 1
                myValues(i, 0) = iVector.Item(i)
            Next
        End If

    End Sub

    Public Shared Operator +(ByVal iA As Matrix, ByVal iB As Matrix) As Matrix

        Try
            If (iA.N <> iB.N) Or (iA.M <> iB.M) Then
                Throw New Exception("Matrix dimensions don't match for addition")
            End If

            Dim result(iA.M - 1, iA.N - 1) As Double
            For i As Integer = 0 To iA.M - 1
                For j As Integer = 0 To iA.N - 1
                    result(i, j) = iA.Value(i, j) + iB.Value(i, j)
                Next
            Next
            Return New Matrix(result)

            
        Catch ex As Exception

        End Try

        Return Nothing

    End Operator

    Public Shared Operator *(ByVal iA As Matrix, ByVal c As Double) As Matrix

        Dim result(iA.M - 1, iA.N - 1) As Double
        For i As Integer = 0 To iA.M - 1
            For j As Integer = 0 To iA.N - 1
                result(i, j) = iA.Value(i, j) * c
            Next
        Next

        Return New Matrix(result)

    End Operator

    Public Shared Operator *(ByVal iA As Matrix, ByVal iB As Matrix) As Matrix

        Try
            If iA.N <> iB.M Then
                Throw New Exception("Matrix dimensions don't match for multiplication")
            End If

            Dim result(iA.M - 1, iB.N - 1) As Double

            For m As Integer = 0 To iA.M - 1
                For n As Integer = 0 To iB.N - 1
                    Dim t As Double = 0
                    For i As Integer = 0 To iA.N - 1
                        t += iA.Value(m, i) * iB.Value(i, n)
                    Next
                    result(m, n) = t
                Next
            Next

            Return New Matrix(result)

        Catch ex As Exception

        End Try

        Return Nothing

    End Operator

    Public Shared Operator *(ByVal iA As Matrix, ByVal iB As Vector) As Vector

        Try
            If iA.N <> iB.Count Then
                Throw New Exception("Matrix and Vector dimensions don't match for multiplication")
            ElseIf iB.Orientation = Vector.OrientationType.horizontal Then
                Throw New Exception("Vector dimensions don't match for multiplication")
            End If

            Dim vals As New GenericList(Of Double)

            For m As Integer = 0 To iA.M - 1
                Dim t As Double = 0
                For i As Integer = 0 To iA.N - 1
                    t += iA.Value(m, i) * iB.Item(i)
                Next
                vals.Add(t)
            Next

            Return New Vector(vals, Vector.OrientationType.vertical)

        Catch ex As Exception

        End Try

        Return Nothing

    End Operator


    Public Shared Function Inverse(ByVal iA As Matrix) As Matrix

        Try

            If iA.M <> iA.N Then
                Throw New Exception("Not square matrix")
            End If

            Dim n As Integer = iA.N


            Dim x(n - 1, n - 1) As Double


            Dim p(n) As Integer

            Dim i, j, k, l As Integer
            Dim t, q As Double

            Dim ufl As Double = Double.Epsilon

            Dim g As Double = 4
            g = g / 3
            g = g - 1
            Dim eps As Double = Abs(((g + g) - 1) + g)

            g = 1
            For j = 0 To n - 1
                p(j) = 0
                For i = 0 To n - 1
                    t = iA.Value(i, j)
                    x(i, j) = t
                    t = Math.Abs(t)

                    If t > p(j) Then p(j) = t
                Next
            Next

            For k = 0 To n - 1
                q = 0
                j = k
                For i = k To n - 1
                    t = Abs(x(i, k))
                    If t > q Then
                        q = t
                        j = i
                    End If
                Next

                If q = 0 Then
                    q = eps * p(k) + ufl
                    x(k, k) = q
                End If
                If p(k) > 0 Then
                    q = q / p(k)
                    If q > g Then g = q
                End If

                If g > 8 * (k + 1) Then
                    Throw New Exception("Growth factor g = " & g & " exceeds " & (8 * (k + 1)))
                End If

                p(k) = j
                If k <> j Then
                    For l = 0 To n - 1
                        q = x(j, l)
                        x(j, l) = x(k, l)
                        x(k, l) = q
                    Next
                End If

                q = x(k, k)
                x(k, k) = 1
                For j = 0 To n - 1
                    x(k, j) = x(k, j) / q
                Next

                For i = 0 To n - 1
                    If i <> k Then
                        q = x(i, k)
                        x(i, k) = 0
                        For j = 0 To n - 1
                            x(i, j) = x(i, j) - x(k, j) * q
                        Next
                    End If
                Next
            Next


            For k = n - 2 To 0 Step -1
                j = p(k)
                If j <> k Then
                    For i = 0 To n - 1
                        q = x(i, k)
                        x(i, k) = x(i, j)
                        x(i, j) = q
                    Next
                End If
            Next

            Return New Matrix(x)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return Nothing


    End Function


    Public Shared Function Transpose(ByVal iA As Matrix) As Matrix

        Dim vals(iA.N - 1, iA.M - 1) As Double

        For i As Integer = 0 To iA.M - 1
            For j As Integer = 0 To iA.N - 1
                vals(j, i) = iA.Value(i, j)
            Next
        Next

        Return New Matrix(vals)

    End Function


    Public Shared Function Identity(ByVal order As Integer) As Matrix

        Dim vals(order - 1, order - 1) As Double

        For i As Integer = 0 To order - 1
            For j As Integer = 0 To order - 1
                If i = j Then
                    vals(i, j) = 1
                Else
                    vals(i, j) = 0
                End If
            Next
        Next

        Return New Matrix(vals)

    End Function

    Public Shared Function Eigenvalues(ByVal Mat As Matrix) As Vector

        Dim iMat(,) As Double = Mat.Values

        Dim evMat(,) As Double = MatEigenvalue_QR(iMat)

        Dim evVec As New GenericList(Of Double)
        For i As Integer = 0 To UBound(evMat, 1)
            evVec.Add(evMat(i, 0))
        Next

        Return New Vector(evVec, Vector.OrientationType.vertical)

    End Function

#Region "Imported EigenValue Functions"

    Shared Function HQR(ByVal n, ByVal low, ByVal igh, ByVal h, ByVal wr, ByVal wi) As Integer
        '
        '     THIS SUBROUTINE IS A TRANSLATION OF THE ALGOL PROCEDURE HQR,
        '     NUM. MATH. 14, 219-231(1970) BY MARTIN, PETERS, AND WILKINSON.
        '     HANDBOOK FOR AUTO. COMP., VOL.II-LINEAR ALGEBRA, 359-371(1971).
        '
        '     THIS SUBROUTINE FINDS THE EIGENVALUES OF A REAL
        '     UPPER HESSENBERG MATRIX BY THE QR METHOD.
        '
        '     ON INPUT
        '
        '        NM MUST BE SET TO THE ROW DIMENSION OF TWO-DIMENSIONAL
        '          ARRAY PARAMETERS AS DECLARED IN THE CALLING PROGRAM
        '          DIMENSION STATEMENT.
        '
        '        N IS THE ORDER OF THE MATRIX.
        '
        '        LOW AND IGH ARE INTEGERS DETERMINED BY THE BALANCING
        '          SUBROUTINE  BALANC.  IF  BALANC  HAS NOT BEEN USED,
        '          SET LOW=1, IGH=N.
        '
        '        H CONTAINS THE UPPER HESSENBERG MATRIX.  INFORMATION ABOUT
        '          THE TRANSFORMATIONS USED IN THE REDUCTION TO HESSENBERG
        '          FORM BY  ELMHES  OR  ORTHES, IF PERFORMED, IS STORED
        '          IN THE REMAINING TRIANGLE UNDER THE HESSENBERG MATRIX.
        '
        '     ON OUTPUT
        '
        '        H HAS BEEN DESTROYED.  THEREFORE, IT MUST BE SAVED
        '          BEFORE CALLING  HQR  IF SUBSEQUENT CALCULATION AND
        '          BACK TRANSFORMATION OF EIGENVECTORS IS TO BE PERFORMED.
        '
        '        WR AND WI CONTAIN THE REAL AND IMAGINARY PARTS,
        '          RESPECTIVELY, OF THE EIGENVALUES.  THE EIGENVALUES
        '          ARE UNORDERED EXCEPT THAT COMPLEX CONJUGATE PAIRS
        '          OF VALUES APPEAR CONSECUTIVELY WITH THE EIGENVALUE
        '          HAVING THE POSITIVE IMAGINARY PART FIRST.  IF AN
        '          ERROR EXIT IS MADE, THE EIGENVALUES SHOULD BE CORRECT
        '          FOR INDICES IERR+1,...,N.
        '
        '        IERR IS SET TO
        '          ZERO       FOR NORMAL RETURN,
        '          J          IF THE LIMIT OF 30*N ITERATIONS IS EXHAUSTED
        '                     WHILE THE J-TH EIGENVALUE IS BEING SOUGHT.
        '
        '     ------------------------------------------------------------------

        Dim i&, j&, k&, L&, m&, en&, ll&, mm&, na&, itn&, its&, MP2&, ENM2&
        Dim p#, q#, r#, s#, t#, w#, x#, y#, ZZ#, tst1#, tst2#
        Dim NOTLAS As Boolean
        '
        Dim Ierr As Integer
        Ierr = 0
        k = 1

Lab50:
        '
        en = igh
        t = 0.0#
        itn = 30 * n
        '     .......... SEARCH FOR NEXT EIGENVALUES ..........
Lab60:
        If (en < low) Then GoTo Lab1001
        its = 0
        na = en - 1
        ENM2 = na - 1
        '     .......... LOOK FOR SINGLE SMALL SUB-DIAGONAL ELEMENT
        '                FOR L=EN STEP -1 UNTIL LOW DO -- ..........
Lab70:
        For ll = low To en
            L = en + low - ll
            If (L = low) Then GoTo Lab100
            s = Abs(h(L - 1, L - 1)) + Abs(h(L, L))
            If (s = 0) Then s = 1 's = norm  ' fix bug 2.11.05 VL
            tst1 = s
            tst2 = tst1 + Abs(h(L, L - 1))
            If (tst2 = tst1) And Abs(h(L, L - 1)) < 1 Then GoTo Lab100
        Next ll
        '     .......... FORM SHIFT ..........
Lab100:
        x = h(en, en)
        If (L = en) Then GoTo Lab270
        y = h(na, na)
        w = h(en, na) * h(na, en)
        If (L = na) Then GoTo Lab280
        If (itn = 0) Then GoTo Lab1000
        If ((its <> 10) And (its <> 20)) Then GoTo Lab130
        '     .......... FORM EXCEPTIONAL SHIFT ..........
        t = t + x
        '
        For i = low To en
            h(i, i) = h(i, i) - x
        Next i
        '
        s = Abs(h(en, na)) + Abs(h(na, ENM2))
        x = 0.75 * s
        y = x
        w = -0.4375 * s * s
Lab130:
        its = its + 1
        itn = itn - 1
        '     .......... LOOK FOR TWO CONSECUTIVE SMALL
        '                SUB-DIAGONAL ELEMENTS.
        '                FOR M=EN-2 STEP -1 UNTIL L DO -- ..........
        For mm = L To ENM2
            m = ENM2 + L - mm
            ZZ = h(m, m)
            r = x - ZZ
            s = y - ZZ
            p = (r * s - w) / h(m + 1, m) + h(m, m + 1)
            q = h(m + 1, m + 1) - ZZ - r - s
            r = h(m + 2, m + 1)
            s = Abs(p) + Abs(q) + Abs(r)
            p = p / s
            q = q / s
            r = r / s
            If (m = L) Then GoTo Lab150
            tst1 = Abs(p) * (Abs(h(m - 1, m - 1)) + Abs(ZZ) + Abs(h(m + 1, m + 1)))
            tst2 = tst1 + Abs(h(m, m - 1)) * (Abs(q) + Abs(r))
            If (tst2 = tst1) Then GoTo Lab150
        Next mm
        '
Lab150:
        MP2 = m + 2
        '
        For i = MP2 To en
            h(i, i - 2) = 0.0#
            If (i <> MP2) Then h(i, i - 3) = 0.0#
        Next i
        '     .......... DOUBLE QR STEP INVOLVING ROWS L TO EN AND
        '                COLUMNS M TO EN ..........
        For k = m To na
            NOTLAS = k <> na
            If (k = m) Then GoTo Lab170
            p = h(k, k - 1)
            q = h(k + 1, k - 1)
            r = 0.0#
            If (NOTLAS) Then r = h(k + 2, k - 1)
            x = Abs(p) + Abs(q) + Abs(r)
            If (x = 0.0#) Then GoTo Lab260
            p = p / x
            q = q / x
            r = r / x
Lab170:
            s = dsign(Sqrt(p * p + q * q + r * r), p)
            If (k = m) Then GoTo Lab180
            h(k, k - 1) = -s * x
            GoTo Lab190
Lab180:
            If (L <> m) Then h(k, k - 1) = -h(k, k - 1)
Lab190:
            p = p + s
            x = p / s
            y = q / s
            ZZ = r / s
            q = q / p
            r = r / p
            If (NOTLAS) Then GoTo Lab225
            '     .......... ROW MODIFICATION ..........
            For j = k To n
                p = h(k, j) + q * h(k + 1, j)
                h(k, j) = h(k, j) - p * x
                h(k + 1, j) = h(k + 1, j) - p * y
            Next j
            '
            j = Min(en, k + 3)
            '     .......... COLUMN MODIFICATION ..........
            For i = 0 To j
                p = x * h(i, k) + y * h(i, k + 1)
                h(i, k) = h(i, k) - p
                h(i, k + 1) = h(i, k + 1) - p * q
            Next i
            GoTo Lab255
Lab225:
            '     .......... ROW MODIFICATION ..........
            For j = k To n
                p = h(k, j) + q * h(k + 1, j) + r * h(k + 2, j)
                h(k, j) = h(k, j) - p * x
                h(k + 1, j) = h(k + 1, j) - p * y
                h(k + 2, j) = h(k + 2, j) - p * ZZ
            Next j
            '
            j = Min(en, k + 3)
            '     .......... COLUMN MODIFICATION ..........
            For i = 0 To j
                p = x * h(i, k) + y * h(i, k + 1) + ZZ * h(i, k + 2)
                h(i, k) = h(i, k) - p
                h(i, k + 1) = h(i, k + 1) - p * q
                h(i, k + 2) = h(i, k + 2) - p * r
            Next i
Lab255:
            '
        Next k
Lab260:
        '
        GoTo Lab70
        '     .......... ONE ROOT FOUND ..........
Lab270:
        wr(en) = x + t
        wi(en) = 0.0#
        en = na
        GoTo Lab60
        '     .......... TWO ROOTS FOUND ..........
Lab280:
        p = (y - x) / 2.0#
        q = p * p + w
        ZZ = Sqrt(Abs(q))
        x = x + t
        If (q < 0.0#) Then GoTo Lab320
        '     .......... REAL PAIR ..........
        ZZ = p + dsign(ZZ, p)
        wr(na) = x + ZZ
        wr(en) = wr(na)
        If (ZZ <> 0.0#) Then wr(en) = x - w / ZZ
        wi(na) = 0.0#
        wi(en) = 0.0#
        GoTo Lab330
        '     .......... COMPLEX PAIR ..........
Lab320:
        wr(na) = x + p
        wr(en) = x + p
        wi(na) = ZZ
        wi(en) = -ZZ
Lab330:
        en = ENM2
        GoTo Lab60
        '     .......... SET ERROR -- ALL EIGENVALUES HAVE NOT
        '                CONVERGED AFTER 30*N ITERATIONS ..........
Lab1000:
        Ierr = en
Lab1001:
        Return Ierr
    End Function


    Private Shared Function dsign(ByVal x, ByVal y)
        If y >= 0 Then
            Return Abs(x)
        Else
            Return -Abs(x)
        End If
    End Function


    Shared Sub ELMHES0(ByVal n, ByVal Mat)
        '  sources by Martin, R. S. and Wilkinson, J. H., see [MART70].   *
        '
        Dim k As Long, i As Long, j As Long, x As Double, y As Double

        For k = 1 To n - 1
            i = k
            x = 0
            For j = k To n
                If (Abs(Mat(j, k - 1)) > Abs(x)) Then
                    x = Mat(j, k - 1)
                    i = j
                End If
            Next j
            If (i <> k) Then
                '           SWAP0 rows and columns of MAT
                For j = k - 1 To n
                    Call SWAP0(Mat(i, j), Mat(k, j))
                Next j
                For j = 0 To n
                    Call SWAP0(Mat(j, i), Mat(j, k))
                Next j
            End If
            If (x <> 0) Then
                For i = k + 1 To n
                    y = Mat(i, k - 1)
                    If (y <> 0) Then
                        y = y / x
                        Mat(i, k - 1) = y
                        For j = k To n
                            Mat(i, j) = Mat(i, j) - y * Mat(k, j)
                        Next j
                        For j = 0 To n
                            Mat(j, k) = Mat(j, k) + y * Mat(j, i)
                        Next j
                    End If
                Next i
            End If
        Next k
    End Sub

    Private Shared Sub SWAP0(ByRef x, ByRef y)
        Dim temp As Double
        temp = x
        x = y
        y = temp
    End Sub


    Shared Function MatEigenvalue_QR(ByVal Mat) As Double(,)
        'Find real and complex eigenvalues with the iterative QR method
        Dim A, wr#(), wi#(), tiny#
        Dim b(,) As Double
        tiny = 2 * 10 ^ -15
        A = Mat
        Dim nn As Integer = UBound(A, 1)

        ReDim wr(nn), wi(nn)

        ELMHES0(nn, A)

        Dim Ierr As Integer = HQR(nn, 0, nn, A, wr, wi)

        ReDim b(nn, 2 - 1)

        For i As Integer = 0 To nn
            If i >= Ierr Then
                b(i, 0) = wr(i)
                b(i, 1) = wi(i)
            End If
        Next
        b = MatMopUp(b, tiny)
        Call MatrixSort(b, "A")
        Return b
    End Function


    Shared Function MatMopUp(ByVal Mat, ByVal ErrMin)
        'eliminates values too small
        Dim A
        A = Mat

        For i As Integer = 1 To UBound(A, 1)
            For j As Integer = 1 To UBound(A, 2)
                If IsNumeric(A(i, j)) Then
                    If Abs(A(i, j)) < ErrMin Then A(i, j) = 0
                End If
            Next j
        Next i
        MatMopUp = A
    End Function


    Shared Sub MatrixSort(ByVal A, ByVal Order)
        '
        'Sorting Routine with swapping algorithm
        'A() may be matrix (N x M) or vector (N)
        'Sort is always based on the first column
        'Order = A (Ascending), D (Descending)
        'Note: it's simple but slow. Use only in non critical part
        '
        Dim c As Double
        Dim flag_exchanged As Boolean
        Dim i_min&, i_max&, j_min&, j_max&, i&, k&, j&

        i_min = LBound(A, 1)
        i_max = UBound(A, 1)
        j_min = LBound(A, 2)
        j_max = UBound(A, 2)

        'Sorting algortithm begin
        Do
            flag_exchanged = False
            For i = i_min To i_max Step 2
                k = i + 1
                If k > i_max Then Exit For
                If (A(i, j_min) > A(k, j_min) And Order = "A") Or _
                   (A(i, j_min) < A(k, j_min) And Order = "D") Then
                    'swap rows
                    For j = j_min To j_max
                        c = A(k, j)
                        A(k, j) = A(i, j)
                        A(i, j) = c
                    Next j
                    flag_exchanged = True
                End If
            Next
            If i_min = LBound(A, 1) Then
                i_min = LBound(A, 1) + 1
            Else
                i_min = LBound(A, 1)
            End If
        Loop Until flag_exchanged = False And i_min = LBound(A, 1)

    End Sub

#End Region

End Class


