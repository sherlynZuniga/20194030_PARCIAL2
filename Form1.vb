Public Class Form1
    Private Sub ProcesarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcesarToolStripMenuItem.Click

        'verificando 
        If ComboBox1.Text = “ “ Or ComboBox2.Text = “” Or ComboBox3.Text = “” Or TextBox1.Text = “” Or TextBox2.Text = “” Then
            MsgBox(“Los campos están vacios”)
            Exit Sub
        End If



        'asignando a variables
        nombre(fila) = TextBox1.Text
        edad(fila) = Val(TextBox2.Text)
        genero(fila) = ComboBox1.SelectedItem
        laboratorio(fila) = ComboBox2.SelectedItem
        examenes_seleccionados = CheckBox1.Checked = True Or CheckBox2.Checked = True Or CheckBox3.Checked = True
        tipoenvio(fila) = ComboBox3.SelectedItem



        'envio de valores
        If (fila <= 5) Then

            'valores en vectores
            valortotalexamenes(fila) = funciontotalexamenes(laboratorio(fila), examenes_seleccionados(fila))
            valorenvio(fila) = (funcionenvio(valorenvio(fila))) * valortotalexamenes(fila)
            total(fila) = valortotalexamenes(fila) + valorenvio(fila)

            'mostrar valores en listbox
            ListBox1.Items.Add(nombre(fila))
            ListBox2.Items.Add(edad(fila))
            ListBox3.Items.Add(genero(fila))
            ListBox4.Items.Add(laboratorio(fila))
            ListBox5.Items.Add(valortotalexamenes(fila))
            ListBox6.Items.Add(valorenvio(fila))
            ListBox7.Items.Add(total(fila))

        Else
            MsgBox("VECTORES LLENOS")
        End If

    End Sub

    Private Sub LimpiarEntradasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarEntradasToolStripMenuItem.Click

        'limpiar entradas 

        TextBox1.Clear()
        TextBox2.Clear()

        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()


    End Sub

    Private Sub LimpiarVectoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarVectoresToolStripMenuItem.Click

        'limpiar vectores

        fila = 0

        TextBox1.Clear()
        TextBox2.Clear()

        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()



    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click

        'formulario de salida
        Form2.Show()

    End Sub
End Class
