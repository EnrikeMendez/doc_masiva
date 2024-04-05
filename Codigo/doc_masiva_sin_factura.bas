Attribute VB_Name = "doc_masiva_sin_factura"
Option Explicit
Option Base 0

'se declaran variables a utilizar
Private n_client As Long

'Private Sub init_var()
'End Sub

Sub carga_archivo(Archivo As String, num_client As String, correo_electronico As String, tipo_carga As String, n_client As String)

'Call init_var

Call log_SQL("carga_archivo", "inicio", num_client)

n_client = Trim(n_client)
num_client = Trim(num_client)
tipo_carga = Trim(UCase(tipo_carga))


Dim My_excel As Excel.Application
Set My_excel = New Excel.Application

Call log_SQL("carga_archivo", "excel abierto con carga " & tipo_carga, num_client)


If tipo_carga = "SIN_FACTURA" Then
    HDR = "Yes"
    Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
End If

End Sub

