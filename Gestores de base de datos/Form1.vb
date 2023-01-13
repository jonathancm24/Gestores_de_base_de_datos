Imports System.Runtime.InteropServices
Imports Npgsql
Imports System.Data.OracleClient
Imports Oracle.ManagedDataAccess.Client
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports ADODB
Imports OracleInternal.Json
Imports System.Data.SqlClient
Imports MySqlConnector.MySqlConnection
Imports MySqlConnector
Imports MySql.Data.MySqlClient
'Botones de postgres
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conn As New NpgsqlConnection("Server=localhost;Database=BDPRODUCTO;Username=postgres;Password=123456")

        Try
            ' Abre la conexión
            conn.Open()

            ' Muestra un mensaje indicando que la conexión es válida
            MessageBox.Show("Conexión válida")
        Catch ex As Exception
            ' Muestra un mensaje de error si la conexión no es válida
            MessageBox.Show("Error de conexión: " & ex.Message)
        Finally
            ' Cierra la conexión
            conn.Close()
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim conn As New NpgsqlConnection("Server=localhost;Database=BDPRODUCTO;Username=postgres;Password=123456")

        ' Abre la conexión
        conn.Open()

        ' Obtiene la hora actual
        Dim startTime As DateTime = DateTime.Now

        ' Obtiene los datos a insertar del cuadro de texto
        Dim datos1 As String = DatoID.Text
        Dim datos2 As String = Datodes.Text
        Dim datos3 As Integer = DatoCosto.Text
        Dim datos4 As Integer = DatoPrecio.Text
        ' Inicializa el contador de inserciones
        Dim inserciones As Integer = 0

        ' Bucle que se ejecuta durante un minuto
        While (DateTime.Now - startTime).TotalSeconds < 60
            ' Crea un nuevo objeto Command y establece el comando de inserción
            Dim cmd As New NpgsqlCommand("INSERT INTO PRODUCTO (ID_PRODUCTO,DESCRIPCION,COSTO,PRECIO) VALUES (@datos1,@datos2,@datos3,@datos4)", conn)

            ' Agrega los parámetros de la consulta
            cmd.Parameters.AddWithValue("@datos1", datos1)
            cmd.Parameters.AddWithValue("@datos2", datos2)
            cmd.Parameters.AddWithValue("@datos3", datos3)
            cmd.Parameters.AddWithValue("@datos4", datos4)
            ' Ejecuta el comando de inserción
            cmd.ExecuteNonQuery()

            ' Incrementa el contador de inserciones
            inserciones += 1
        End While

        ' Muestra el número de inserciones realizadas en el cuadro de texto
        InsercionesPLMD.Text = "Se insertaron " + inserciones.ToString() + " registros."

        ' Cierra la conexión
        conn.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Crea una nueva conexión a la base de datos
        Dim conn As New NpgsqlConnection("Server=localhost;Database=BDPRODUCTO;Username=postgres;Password=123456")

        ' Abre la conexión
        conn.Open()

        ' Crea un nuevo objeto Command y establece el comando de eliminación de datos
        Dim cmd As New NpgsqlCommand("DELETE FROM PRODUCTO", conn)

        ' Ejecuta el comando de eliminación
        cmd.ExecuteNonQuery()
        ' Muestra un mensaje de confirmación
        MessageBox.Show("Los registros han sido eliminados correctamente.")

        ' Cierra la conexión
        conn.Close()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ' Crea una nueva conexión a la base de datos
        Dim conn As New NpgsqlConnection("Server=localhost;Database=BDPRODUCTO;Username=postgres;Password=123456")

        ' Abre la conexión
        conn.Open()
        ' Obtiene la hora actual
        Dim startTime As DateTime = DateTime.Now

        ' Obtiene los datos a insertar del cuadro de texto
        Dim datos1 As String = DatoID.Text
        Dim datos2 As String = Datodes.Text
        Dim datos3 As Integer = DatoCosto.Text
        Dim datos4 As Integer = DatoPrecio.Text
        ' Inicializa el contador de inserciones
        Dim inserciones As Integer = 0

        ' Bucle que se ejecuta durante un minuto
        While (DateTime.Now - startTime).TotalSeconds < 60
            ' Crea un nuevo objeto Command y establece el nombre del SP
            Dim cmd As New NpgsqlCommand("PA_INSERTARPRODUCTO", conn)
            cmd.CommandType = CommandType.StoredProcedure

            ' Agrega los parámetros de la consulta
            cmd.Parameters.AddWithValue("pidproducto", datos1)
            cmd.Parameters.AddWithValue("pdescripcion", datos2)
            cmd.Parameters.AddWithValue("pcosto", datos3)
            cmd.Parameters.AddWithValue("pprecio", datos4)
            ' Ejecuta el comando de inserción
            cmd.ExecuteNonQuery()

            ' Incrementa el contador de inserciones
            inserciones += 1
        End While

        ' Muestra el número de inserciones realizadas en el cuadro de texto
        InsercionesPSP.Text = "Se insertaron " + inserciones.ToString() + " registros."

        ' Cierra la conexión
        conn.Close()

    End Sub
    '************************************Botones de Oracle************************************************
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim conn As New OracleConnection("Data Source=localhost;Persist Security Info=True;User ID=system;Password=123456")

        Try
            ' Abre la conexión
            conn.Open()

            ' Muestra un mensaje indicando que la conexión es válida
            MessageBox.Show("Conexión válida")
        Catch ex As Exception
            ' Muestra un mensaje de error si la conexión no es válida
            MessageBox.Show("Error de conexión: " & ex.Message)
        Finally
            ' Cierra la conexión
            conn.Close()
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim conn As New OracleConnection("Data Source=localhost;Persist Security Info=True;User ID=system;Password=123456")
        ' Abre la conexión
        conn.Open()

        ' Obtiene la hora actual
        Dim startTime As DateTime = DateTime.Now

        ' Inicializa el contador de inserciones
        Dim inserciones As Integer = 0

        While (DateTime.Now - startTime).TotalSeconds < 60
            Dim cmd As New OracleCommand("INSERT INTO PRODUCTO (ID_PRODUCTO, DESCRIPCION , COSTO, PRECIO) VALUES (:ID_PRODUCTO, :DESCRIPCION, :COSTO, :PRECIO)", conn)

            cmd.Parameters.Add(New OracleParameter(":ID_PRODUCTO", OracleDbType.Varchar2, 20)).Value = DatoID.Text
            cmd.Parameters.Add(New OracleParameter(":DESCRIPCION", OracleDbType.Varchar2, 20)).Value = Datodes.Text
            cmd.Parameters.Add(New OracleParameter(":COSTO", OracleDbType.Int32)).Value = CInt(DatoCosto.Text)
            cmd.Parameters.Add(New OracleParameter(":PRECIO", OracleDbType.Int32)).Value = CInt(DatoPrecio.Text)
            cmd.ExecuteNonQuery()

            ' Incrementa el contador de inserciones
            inserciones += 1
        End While

        ' Muestra el número de inserciones realizadas en el cuadro de texto
        InsercionesOLMD.Text = "Se insertaron " + inserciones.ToString() + " registros."

        conn.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim conn As New OracleConnection("Data Source=localhost;Persist Security Info=True;User ID=system;Password=123456")
        conn.Open()

        Dim cmd As New OracleCommand("DELETE FROM PRODUCTO", conn)
        cmd.ExecuteNonQuery()
        ' Muestra un mensaje de confirmación
        MessageBox.Show("Los registros han sido eliminados correctamente.")
        conn.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim conn As New OracleConnection("Data Source=localhost;Persist Security Info=True;User ID=system;Password=123456")
        ' Abre la conexión
        conn.Open()

        ' Obtiene la hora actual
        Dim startTime As DateTime = DateTime.Now

        ' Inicializa el contador de inserciones
        Dim inserciones As Integer = 0

        While (DateTime.Now - startTime).TotalSeconds < 60
            Dim cmd As New OracleCommand("PA_INSERTARPRODUCTO", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add(New OracleParameter(":PIDPRODUCTO", OracleDbType.Varchar2, 20)).Value = DatoID.Text
            cmd.Parameters.Add(New OracleParameter(":PDESCRIPCION", OracleDbType.Varchar2, 20)).Value = Datodes.Text
            cmd.Parameters.Add(New OracleParameter(":PCOSTO", OracleDbType.Int32)).Value = CInt(DatoCosto.Text)
            cmd.Parameters.Add(New OracleParameter(":PPRECIO", OracleDbType.Int32)).Value = CInt(DatoPrecio.Text)
            cmd.ExecuteNonQuery()
            ' Incrementa el contador de inserciones
            inserciones += 1
        End While
        ' Muestra el número de inserciones realizadas en el cuadro de texto
        InsercionesOSP.Text = "Se insertaron " + inserciones.ToString() + " registros."

        conn.Close()
    End Sub

    '***********************************************Botones de SQL Server****************************************
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Declare variables
        Dim connectionString As String
        Dim conn As New ADODB.Connection
        connectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-3SUCHO1\SQLEXPRESS;Initial Catalog=BDPRODUCTO;Integrated Security=SSPI;"
        Try
            ' Abre la conexión
            conn.Open(connectionString)

            ' Muestra un mensaje indicando que la conexión es válida
            MessageBox.Show("Conexión válida")
        Catch ex As Exception
            ' Muestra un mensaje de error si la conexión no es válida
            MessageBox.Show("Error de conexión: " & ex.Message)
        Finally
            ' Cierra la conexión
            conn.Close()
        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Declare variables
        Dim connectionString As String
        Dim conn As New ADODB.Connection
        Dim sql As String
        connectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-3SUCHO1\SQLEXPRESS;Initial Catalog=BDPRODUCTO;Integrated Security=SSPI;"
        conn.Open(connectionString)
        ' Obtiene la hora actual
        Dim startTime As DateTime = DateTime.Now

        ' Inicializa el contador de inserciones
        Dim inserciones As Integer = 0
        ' Obtiene los datos a insertar del cuadro de texto
        Dim datos1 As String = DatoID.Text
        Dim datos2 As String = Datodes.Text
        Dim datos3 As Integer = DatoCosto.Text
        Dim datos4 As Integer = DatoPrecio.Text
        While (DateTime.Now - startTime).TotalSeconds < 60
            'Set the SQL query
            sql = "INSERT INTO PRODUCTO (ID_PRODUCTO, DESCRIPCION, COSTO, PRECIO) VALUES ('" & datos1 & "', '" & datos2 & "','" & datos3 & "','" & datos4 & "')"

            'Execute the query
            conn.Execute(sql)

            ' Incrementa el contador de inserciones
            inserciones += 1
        End While
        ' Muestra el número de inserciones realizadas en el cuadro de texto
        InsersionSSLMD.Text = "Se insertaron " + inserciones.ToString() + " registros."

        conn.Close()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'Declare variables
        Dim connectionString As String
        Dim conn As New ADODB.Connection
        Dim sql As String
        connectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-3SUCHO1\SQLEXPRESS;Initial Catalog=BDPRODUCTO;Integrated Security=SSPI;"
        conn.Open(connectionString)

        sql = "DELETE FROM PRODUCTO"
        conn.Execute(sql)
        ' Muestra un mensaje de confirmación
        MessageBox.Show("Los registros han sido eliminados correctamente.")
        conn.Close()


    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim datos1 As String = DatoID.Text
        Dim datos2 As String = Datodes.Text
        Dim datos3 As Integer
        Dim datos4 As Integer
        Dim inserciones As Integer = 0
        ' Valida los datos de entrada 
        If Not Integer.TryParse(DatoCosto.Text, datos3) Or Not Integer.TryParse(DatoPrecio.Text, datos4) Then
            MessageBox.Show("Ingrese solo valores numéricos para el costo y precio!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        ' Cadena de conexión a la base de datos
        Dim conexion As New SqlConnection("Data Source=DESKTOP-3SUCHO1\SQLEXPRESS;Initial Catalog=BDPRODUCTO;Integrated Security=True")
        Try
            ' Abrir la conexión
            conexion.Open()
            ' Obtiene la hora actual
            Dim startTime As DateTime = DateTime.Now
            While (DateTime.Now - startTime).TotalSeconds < 60
                ' llamar al sp y sus propiedades
                Using cmd As New SqlCommand("PA_INSERTARPRODUCTO", conexion)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@PIDPRODUCTO", datos1)
                    cmd.Parameters.AddWithValue("@PDESCRIPCION", datos2)
                    cmd.Parameters.AddWithValue("@PCOSTO", datos3)
                    cmd.Parameters.AddWithValue("@PPRECIO", datos4)
                    ' Ejecutar el procedimiento almacenado
                    cmd.ExecuteNonQuery()
                End Using
                ' Incrementa el contador de inserciones
                inserciones += 1
            End While
            ' Mostrar mensaje de éxito al usuario
            InsersionesSSSP.Text = "Se insertaron " + inserciones.ToString() + " registros."
            MessageBox.Show("Inserción realizada exitosamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As SqlException
            ' Manejar excepción
            MessageBox.Show("Error al insertar el registro: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Cerrar la conexión
            conexion.Close()
        End Try
    End Sub


    '*******************************************Botones de MYSQL*************************************************************
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim cadenaConexion = "Server=127.0.0.1;User ID=root;Password=123456;Database=BDPRODUCTO"
        Dim connection As New MySql.Data.MySqlClient.MySqlConnection(cadenaConexion)

        Dim cmd As New MySql.Data.MySqlClient.MySqlCommand
        Try

            ' SI SE ABRE LA CONEXIÓN MANDAMOS MENSAJE
            connection.Open()
            MsgBox("Conexion Correcta")

        Catch ex As Exception
            ' SI NO SE ABRE LA CONEXIÓN MANDAMOS MENSAJE DE ERROR
            MsgBox("ERROR " & ex.Message)

            Exit Sub
        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        Dim cmd As New MySql.Data.MySqlClient.MySqlCommand

        Dim connection As New MySql.Data.MySqlClient.MySqlConnection("Server=127.0.0.1;User ID=root;Password=123456;Database=BDPRODUCTO")



        Dim datos1 As String = DatoID.Text
        Dim datos2 As String = Datodes.Text
        Dim datos3 As Integer
        Dim datos4 As Integer
        Dim inserciones As Integer = 0
        ' Valida los datos de entrada 
        If Not Integer.TryParse(DatoCosto.Text, datos3) Or Not Integer.TryParse(DatoPrecio.Text, datos4) Then
            MessageBox.Show("Ingrese solo valores numéricos para el costo y precio!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Try

            ' Abre la conexión
            connection.Open()
            cmd.Connection = connection
            ' Obtiene la hora actual
            Dim startTime As DateTime = DateTime.Now
            While (DateTime.Now - startTime).TotalSeconds < 60

                ' Obtiene la hora actual

                ' Agregando parametros
                cmd.CommandText = "INSERT INTO producto (ID_PRODUCTO,DESCRIPCION ,COSTO ,PRECIO) VALUES (@PIDPRODUCTO, @PDESCRIPCION, @PCOSTO, @PPRECIO)"
                cmd.Parameters.AddWithValue("@PIDPRODUCTO", datos1)
                cmd.Parameters.AddWithValue("@PDESCRIPCION", datos2)
                cmd.Parameters.AddWithValue("@PCOSTO", datos3)
                cmd.Parameters.AddWithValue("@PPRECIO", datos4)
                cmd.ExecuteNonQuery()
                ' Incrementa el contador de inserciones
                inserciones += 1
                cmd.Parameters.Clear()
            End While
            ' Mostrar mensaje de éxito al usuario
            InsercionesMSLMD.Text = "Se insertaron " + inserciones.ToString() + " registros."
            MessageBox.Show("Inserción realizada exitosamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As MySql.Data.MySqlClient.MySqlException
            ' Manejar excepción
            MessageBox.Show("Error al insertar el registro: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Cerrar la conexión
            Connection.Close()
        End Try
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim connection As New MySql.Data.MySqlClient.MySqlConnection("Server=127.0.0.1;User ID=root;Password=123456;Database=BDPRODUCTO")
        Dim cmd As New MySql.Data.MySqlClient.MySqlCommand()
        cmd.Connection = connection
        Try
            connection.Open()
            cmd.CommandText = "DELETE FROM producto"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Datos eliminados correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As MySql.Data.MySqlClient.MySqlException
            ' Manejar excepción
            MessageBox.Show("Error al eliminar los datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            connection.Close()
        End Try
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        Dim cmd As New MySql.Data.MySqlClient.MySqlCommand

        Dim connection As New MySql.Data.MySqlClient.MySqlConnection("Server=127.0.0.1;User ID=root;Password=123456;Database=BDPRODUCTO;CharSet=utf8mb4")
        cmd.Connection = connection
        Dim datos1 As String = DatoID.Text
        Dim datos2 As String = Datodes.Text
        Dim datos3 As Integer
        Dim datos4 As Integer
        Dim inserciones As Integer = 0
        ' Valida los datos de entrada 
        If Not Integer.TryParse(DatoCosto.Text, datos3) Or Not Integer.TryParse(DatoPrecio.Text, datos4) Then
            MessageBox.Show("Ingrese solo valores numéricos para el costo y precio!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Try

            ' Abre la conexión
            connection.Open()
            cmd.Connection = connection
            ' Obtiene la hora actual
            Dim startTime As DateTime = DateTime.Now
            While (DateTime.Now - startTime).TotalSeconds < 60
                cmd.Parameters.Clear()
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandText = "PA_INSERTARPRODUCTO"
                cmd.Parameters.AddWithValue("@PIDPRODUCTO", datos1)
                cmd.Parameters.AddWithValue("@PDESCRIPCION", datos2)
                cmd.Parameters.AddWithValue("@PCOSTO", datos3)
                cmd.Parameters.AddWithValue("@PPRECIO", datos4)
                cmd.ExecuteNonQuery()

                ' Incrementa el contador de inserciones
                inserciones += 1
            End While
            ' Mostrar mensaje de éxito al usuario
            InsercionesMYSP.Text = "Se insertaron " + inserciones.ToString() + " registros."
            MessageBox.Show("Inserción realizada exitosamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As MySql.Data.MySqlClient.MySqlException
            ' Manejar excepción
            MessageBox.Show("Error al insertar el registro: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Cerrar la conexión
            Connection.Close()
        End Try
    End Sub
End Class
