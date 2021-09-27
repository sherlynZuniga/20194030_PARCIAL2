Module Module1

    'variables
    Public nombre(5) As String
    Public edad(5) As Integer
    Public genero(5)
    Public laboratorio(5)
    Public valortotalexamenes(5) As Double
    Public tipoenvio(5)
    Public valorenvio(5) As Double
    Public total(5) As Double
    Public laboratorio_seleccionado As String
    Public examenes_seleccionados As String
    Public examenes

    Public fila As Byte = 0

    'variables para funciones
    Public totalexamenes As Double = 0
    Public envio As Double = 0

    'funcion para total de exámenes

    Function funciontotalexamenes(laboratorio_Seleccionado As String, examenes_seleccionados As String) As Double

        Select Case Form1.ComboBox2.SelectedIndex
            'majadas
            Case 0
                If Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = False And Form1.CheckBox3.Checked = False Then
                    totalexamenes = 200.99
                ElseIf Form1.CheckBox2.Checked = True And Form1.CheckBox1.Checked = False And Form1.CheckBox3.Checked = False Then
                    totalexamenes = 150.5
                ElseIf Form1.CheckBox3.Checked = True And Form1.CheckBox1.Checked = False And Form1.CheckBox2.Checked = False Then
                    totalexamenes = 100.99
                ElseIf Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = True And Form1.CheckBox3.Checked = True Then
                    totalexamenes = 200.99 + 150.5 + 100.99
                ElseIf Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = True And Form1.CheckBox3.Checked = False Then
                    totalexamenes = 200.99 + 150.5
                ElseIf Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = False And Form1.CheckBox3.Checked = True Then
                    totalexamenes = 200.99 + 100.99
                ElseIf Form1.CheckBox1.Checked = False And Form1.CheckBox2.Checked = True And Form1.CheckBox3.Checked = True Then
                    totalexamenes = 150.5 + 100.99
                End If

            'multimedica
            Case 1
                If Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = False And Form1.CheckBox3.Checked = False Then
                    totalexamenes = 210.9
                ElseIf Form1.CheckBox2.Checked = True And Form1.CheckBox1.Checked = False And Form1.CheckBox3.Checked = False Then
                    totalexamenes = 145.5
                ElseIf Form1.CheckBox3.Checked = True And Form1.CheckBox1.Checked = False And Form1.CheckBox2.Checked = False Then
                    totalexamenes = 110.99
                ElseIf Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = True And Form1.CheckBox3.Checked = True Then
                    totalexamenes = 2010.9 + 145.5 + 110.99
                ElseIf Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = True And Form1.CheckBox3.Checked = False Then
                    totalexamenes = 2010.9 + 145.5
                ElseIf Form1.CheckBox1.Checked = True And Form1.CheckBox2.Checked = False And Form1.CheckBox3.Checked = True Then
                    totalexamenes = 2010.9 + 110.99
                ElseIf Form1.CheckBox1.Checked = False And Form1.CheckBox2.Checked = True And Form1.CheckBox3.Checked = True Then
                    totalexamenes = 145.5 + 110.99
                End If
        End Select
        Return totalexamenes
    End Function

    'funcion para tipo de envio
    Function funcionenvio(valorenvio As String) As Double
        Select Case Form1.ComboBox3.SelectedIndex
            'Correo electrónico
            Case 0
                envio = 5 / 100

                'mensajeria
            Case 1
                envio = 6 / 100

                'personal
            Case 2
                envio = -25 / 100
        End Select
        Return funcionenvio
    End Function




End Module
