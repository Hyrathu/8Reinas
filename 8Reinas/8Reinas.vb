Public Class Form1
    Dim Tablero(7, 7) As Integer
    Dim damas As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i, j As Integer

        For i = 0 To 7
            For j = 0 To 7
                Dim pb As New PictureBox
                pb.Size = New System.Drawing.Size(50, 50)
                pb.Location = New System.Drawing.Point(20 + (50 * j), 20 + (50 * i))
                pb.Name = "pbCelda" & i & j

                If (((i + j) Mod 2) = 0) Then
                    pb.Tag = "b"
                    pb.BackColor = Color.White
                Else
                    pb.Tag = "n"
                    pb.BackColor = Color.Black
                End If

                Me.Controls.Add(pb)

                AddHandler pb.MouseClick, AddressOf PictureBox_MouseClick

            Next
        Next
    End Sub

    Private Sub PictureBox_MouseClick(sender As Object, e As EventArgs)
        Dim nombre As String = sender.name
        Dim f, c As Integer
        Dim EsPosible As Integer = 1

        f = nombre.Substring(7, 1)
        c = nombre.Substring(8, 1)
        'MsgBox("Fila=" & f & vbCrLf & "Columna=" & c)

        If (Tablero(f, c) = 1) Then
            Tablero(f, c) = 0
            damas -= 1
            If (sender.tag = "b") Then
                sender.backcolor = Color.White
            Else
                sender.backcolor = Color.Black
            End If
        Else
            EsPosible *= ComprobarFila(f)
            EsPosible *= ComprobarColumna(c)
            EsPosible *= ComprobarDiagonal1(f, c)
            EsPosible *= ComprobarDiagonal2(f, c)

            If EsPosible = 1 Then
                sender.backcolor = Color.Red
                Tablero(f, c) = 1
                damas += 1
            Else
                MsgBox("No es posible poner la Dama")
            End If
        End If

        If damas = 3 Then
            Dim valor As Integer
            valor = MsgBox("Has ganado!!!!!. ¿Desea reiniciar el juego?", vbYesNo)

            If (valor.Equals(6)) Then
                Application.Restart()

            End If
        End If

    End Sub

    Private Function ComprobarFila(fila As Integer)
        Dim i As Integer

        For i = 0 To 7
            If (Tablero(fila, i) = 1) Then
                Return 0
            End If
        Next

        Return 1
    End Function

    Private Function ComprobarColumna(columna As Integer)
        Dim i As Integer

        For i = 0 To 7
            If (Tablero(i, columna) = 1) Then
                Return 0
            End If
        Next

        Return 1
    End Function

    Private Function ComprobarDiagonal1(fila As Integer, columna As Integer)
        Dim Menor As Integer
        Dim saltos, i As Integer

        Menor = Math.Min(fila, columna)
        fila -= Menor
        columna -= Menor

        If (fila >= columna) Then
            saltos = 7 - fila + columna
        End If

        For i = 0 To saltos
            If (Tablero(fila + i, columna + i) = 1) Then
                Return 0
            End If
        Next
        Return 1
    End Function

    Private Function ComprobarDiagonal2(fila As Integer, columna As Integer)
        Dim i, dif As Integer
        columna += fila
        fila = 0
        If (columna > 7) Then
            dif = columna - 7
            fila = dif
            columna = 7
        End If

        For i = 0 To (columna - fila)
            If (Tablero(fila + i, columna - i) = 1) Then
                Return 0
            End If
        Next

        Return 1
    End Function


End Class
