Attribute VB_Name = "web_carga_archivo"
Option Explicit
Option Base 0



Private SQL As String
Private rs As New ADODB.Recordset
Private rs2 As New ADODB.Recordset
Private rs3 As New ADODB.Recordset
Private oRS As New ADODB.Recordset
Private oRS2 As New ADODB.Recordset
Private oRS3 As New ADODB.Recordset


Private jmail As New jmail.Message

Dim sData As String
Private oConn As ADODB.Connection
Private colSheets As New Collection

Private id_factura As Integer
Private num_factura As String

Private id_pedido_cliente As Integer
Private num_pedido_cliente As String

Private id_orden_compra As Integer
Private num_orden_compra As String

Private id_fecha_factura As Integer
Private num_fecha_factura As String

Private id_cil_num_dir As Integer

Private id_dec_dir As Integer

Private id_cdad_bultos As Integer
Private num_cdad_bultos As Long
Private num_cdad_tarimas As Long
Private num_cdad_cajas_tarima As Long

Private id_cclclave As Integer
Private num_cclclave As Long
Private num_direccion_cliente As String
Private my_cclclave As String

Private id_dieclave As Integer
Private num_dieclave As Long

Private num_dieclave_entrega As Long

Private id_importe As Integer
Private num_importe As String

Private id_allclave_ori As Integer
Private num_allclave_ori As Integer

Private id_allclave_dest As Integer
Private num_allclave_dest As Integer

Private id_peso As Integer
Private num_peso As Double
Private id_volumen As Integer
Private num_volumen As Double

Private id_collect_prepaid As Integer
Private num_collect_prepaid As String

Private id_ref As Integer
Private num_ref As String

Private id_eirclef_cdad As Long
Private num_eirclef_cdad As Long
Private num_eirclef As Long
Private num_disclef As Long

'Johnson & Johnson ------------------>
Private id_fecha_cita_prog As Integer
Private num_fecha_cita_prog As String
'Johnson & Johnson ------------------>

'20180606 -- > Carga GSK
Private id_tarimas As Integer
Private num_cdad_empaques_x_bulto As String
Private num_tpaclave As String
Private num_bulto_tpaclave As String
'20181019 -- >
Private num_wcd_cajas_tarimas As Long
Private num_wcd_cdad_tarimas As Long
Private num_wcd_cdad_cajas As Long
'20181019 < --
'20180606 < --

'COMPLEMENTO
Private id_fac_complemento As Integer
Private num_fac_complemento As String


Private num_observacion As String

Private pestana_encabezados, pestana_detalles As String

Private error_msg As String
Private facturas_insertadas As String
Private nuis_insertados As String
Private clef As Long
Private status As String
Private HDR As String

Private lineas_factura As Integer

Private i As Integer
Private j As Integer


'20180614 -- >
Private talones_actualizar() As String
Private actualiza As Integer
Private posicion As Integer
Private facturas_actualizadas As String
Private facturas_error As String
Private facturas_error2 As String
'20180614 < --

Private lstFactura As String
Private lstOrdenC As String

Private Sub init_var()
                  
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenForwardOnly
    rs2.LockType = adLockBatchOptimistic
    rs2.ActiveConnection = Db_link_orfeo
    
    Set rs3 = New ADODB.Recordset
    rs3.CursorLocation = adUseClient
    rs3.CursorType = adOpenForwardOnly
    rs3.LockType = adLockBatchOptimistic
    rs3.ActiveConnection = Db_link_orfeo
        
        actualiza = 0
        num_cdad_tarimas = 0
        num_cdad_cajas_tarima = 0
        clef = 0
        nuis_insertados = ""
        my_cclclave = ""
        j = 0
        lineas_factura = 1
                
        'COMPLEMENTO
        num_fac_complemento = "N"
    
End Sub

Sub carga_archivo(Archivo As String, cliente As String, correo_electronico As String, tipo_carga As String, mi_disclef As String)
    '<<<CC-DESA-31102023-01: agrego funcionalidad para carga GSK de tipo ex-cross_dock.
    If tipo_carga = "UNICA" Then        '---------(DOCUMENTACION MASIVA DE NUIS CON DOCUMENTO FUENTE)---------
        Call carga_ltl_doc_fte(Archivo, cliente, correo_electronico, tipo_carga, mi_disclef)
        Exit Sub
    '<-- CARGA DE DOC FUENTES       ---------(DOCUMENTACION MASIVA DE FACTURAS)---------
    ElseIf tipo_carga = "FUENTE" Then
        Call carga_facturas_ltl(Archivo, cliente, correo_electronico, tipo_carga, mi_disclef)
        Exit Sub
    'CARGA DE DOC FUENTES -->
    '<-- DOCUMENTACION DE NUIS SIN FACTURA
    ElseIf tipo_carga = "SIN_FACTURA" Then
        Call doc_masiva_sin_fact(Archivo, cliente, correo_electronico, mi_disclef)
        Exit Sub
    'DOCUMENTACION DE NUIS SIN FACTURA -->
	End If
    '   CC-DESA-31102023-01>>>

'archivo viene compuesto de lo siguiente:
'1er archivo con la letra de Drive del servidor en este caso D:
'| seguido por la IP
'| y 2o archivo si hay un otro

Dim numero_factura As String

Call init_var

Call log_SQL("carga_archivo", "inicio", cliente)

num_disclef = Trim(mi_disclef)
cliente = Trim(cliente)
tipo_carga = Trim(UCase(tipo_carga))


Dim My_excel As Excel.Application
Set My_excel = New Excel.Application

Call log_SQL("carga_archivo", "excel abierto con carga " & tipo_carga, cliente)

'preparacion de archivos
Select Case cliente
    Case "3882"
        'Cerraduras Phillips
        My_excel.Workbooks.OpenText filename:="""\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "") & """", Origin:= _
                xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
                , Comma:=False, Space:=False, Other:=True, OtherChar:="|", FieldInfo _
                :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2))
        My_excel.ActiveSheet.Name = "encabezados"
        
        My_excel.Workbooks.OpenText filename:="""\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(2), Split(Split(Archivo, "|")(2), "\")(0), "") & """", Origin:= _
                xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
                , Comma:=False, Space:=False, Other:=True, OtherChar:="|", FieldInfo _
                :=Array(Array(1, 2), Array(2, 2), Array(3, 2))
        My_excel.ActiveSheet.Name = "detalles"
        
        My_excel.Workbooks(2).Worksheets(1).Move , My_excel.Workbooks(1).Worksheets(1)
        
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "") & ".xls"
        My_excel.ActiveWorkbook.SaveAs Archivo, xlNormal, , , , , , xlLocalSessionChanges
    
        My_excel.Quit

        HDR = "No"
        
    Case "3624"
        HDR = "No"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
        My_excel.Workbooks.Open Archivo
        My_excel.DisplayAlerts = False
        My_excel.ActiveWorkbook.SaveAs Archivo, xlNormal, , , , , , True
        My_excel.Quit
        
    Case "3885"
        HDR = "No"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
        
    'Johnson & Johnson 15178 ------------------------------------------->
    Case "15178"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    'Johnson & Johnson 15178 ------------------------------------------->

    Case "3081"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")

    Case "13128"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")

    Case "17873"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    
    '20180606 -- > Carga GSK
    Case "20341"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    
    Case "20305"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    '   20181001 -- > Carga GSK
    Case "20501"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    
    Case "20502"
        HDR = "Yes"
        Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    '   20181001  -- Carga GSK
    '20180606 < -- Carga GSK

    ''<< 20220525: Cliente de pruebas
    'Case "20123"
    '    HDR = "Yes"
    '    Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
    ''   20220525 >>
End Select

If tipo_carga = "SIN_FACTURA" Then
    HDR = "Yes"
    Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
End If

If tipo_carga = "CON_FACTURA" Then
    HDR = "Yes"
    Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
End If

Set oConn = New ADODB.Connection
oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Archivo & ";" & _
           "Extended Properties=""Excel 8.0;HDR=" & HDR & ";IMEX=1;"""
           'HDR=YES propiedad que permite poner leer la 1ra linea como nombre de los campos
           'IMEX=1 propiedad que permite leer correctamente la celda en formato texto


Call log_SQL("carga_archivo", "ocbd excel abierto", cliente)

'obtenemos el nombre de la hoja de trabajo(la tabla)
Set colSheets = GetAllXLSheetNames(Archivo, False)

'verificacion del remitente
status = get_remitente(cliente, mi_disclef, num_allclave_ori)
If status <> "ok" Then
    'hubo un error al recuperar la direccion
    error_msg = error_msg & status & vbCrLf & vbCrLf
End If
   

'reset campos:
id_factura = -1
id_pedido_cliente = -1
id_orden_compra = -1
id_fecha_factura = -1
id_dec_dir = -1
id_cil_num_dir = -1
id_cdad_bultos = -1
id_cclclave = -1
id_dieclave = -1
id_importe = -1
id_allclave_ori = -1
id_allclave_dest = -1
id_peso = -1
id_volumen = -1
id_collect_prepaid = -1
id_ref = -1
'Johnson & Johnson 15178 ------------------------------------------->
id_fecha_cita_prog = -1
'Johnson & Johnson 15178 ------------------------------------------->

'20180606 -- > Carga GSK
id_tarimas = -1
'20180606 < --

'COMPLEMENTO
id_fac_complemento = -1


'configuracion de campos
Select Case cliente
    Case "3624"
        'posiciones de campos para Random:
        id_factura = 4
        id_pedido_cliente = 1
        id_fecha_factura = 2
        id_orden_compra = 3
        id_dec_dir = 5
        id_cdad_bultos = 8
        id_importe = 11
        
        pestana_encabezados = colSheets.Item(1)
        

    Case "3882"
        'posiciones de campos para Phillips:
        id_cclclave = 4
        id_factura = 2
        id_cdad_bultos = 7
        id_peso = 8
        id_collect_prepaid = 9
        
        'detalles:
        id_ref = 2
        id_eirclef_cdad = 3
        pestana_encabezados = "encabezados$"
        pestana_detalles = "detalles$"
    
    Case "3885"
        id_factura = 1
        id_fecha_factura = 3
        id_dec_dir = 5
        id_peso = 9
        id_volumen = 10
        id_orden_compra = 14    'sirve para crear los vendedores CVECLAVE
        
        'detalle
        id_ref = 1
        id_eirclef_cdad = 3
        id_cdad_bultos = 4
        
        Dim sSheetName
        For Each sSheetName In colSheets
            If LCase(sSheetName) = "detalle$" Then
                pestana_detalles = "detalle$"
            ElseIf InStr(LCase(sSheetName), "corte") Then
                pestana_encabezados = sSheetName
            End If
        Next
        'Johnson & Johnson 15178 ------------------------------------------->
    Case "15178"
        id_orden_compra = 0
        id_factura = 2
        id_dec_dir = 4
        id_cdad_bultos = 8
        id_fecha_cita_prog = 9
        'pestana_encabezados = colSheets.Item(1)
        pestana_encabezados = "CRUCE1$"
        'Johnson & Johnson 15178 ------------------------------------------->
    
    Case "3081"
        id_factura = 1
        id_pedido_cliente = 0
        id_orden_compra = 3
        id_fecha_factura = 5
        id_dec_dir = 10
        id_cdad_bultos = 13
        id_importe = 16
        pestana_encabezados = "Hoja1$"
        
    Case "13128"
        id_factura = 0
        id_cil_num_dir = 1
        id_dec_dir = 2
        id_cdad_bultos = 5
        id_peso = 6
        id_fecha_factura = 9
        pestana_encabezados = colSheets.Item(1)
        
    Case "17873"
        id_factura = 0
        id_fecha_factura = 1
        id_dec_dir = 3
        id_cdad_bultos = 4
        pestana_encabezados = colSheets.Item(1)
    
    Case "20341"
        id_fecha_factura = 1
        id_orden_compra = 2
        id_factura = 3
        id_pedido_cliente = 4
        id_dec_dir = 5
        id_cdad_bultos = 10
        id_tarimas = 11
        id_fecha_cita_prog = 13
        pestana_encabezados = colSheets.Item(1)
        
    Case "20305"
        id_fecha_factura = 1
        id_orden_compra = 2
        id_factura = 3
        id_pedido_cliente = 4
        id_dec_dir = 5
        id_cdad_bultos = 10
        id_tarimas = 11
        id_fecha_cita_prog = 13
        pestana_encabezados = colSheets.Item(1)
        
    '   20181001 -- > Carga GSK
    Case "20501"
        id_fecha_factura = 1
        id_orden_compra = 2
        id_factura = 3
        id_pedido_cliente = 4
        id_dec_dir = 5
        id_cdad_bultos = 10
        id_tarimas = 11
        id_fecha_cita_prog = 13
        pestana_encabezados = colSheets.Item(1)
        
    Case "20502"
        id_fecha_factura = 1
        id_orden_compra = 2
        id_factura = 3
        id_pedido_cliente = 4
        id_dec_dir = 5
        id_cdad_bultos = 10
        id_tarimas = 11
        id_fecha_cita_prog = 13
        pestana_encabezados = colSheets.Item(1)
    '   20181001 < -- Carga GSK
    
    ''<<20220525: Cliente de pruebas
    'Case "20123"
    '    id_fecha_factura = 1
    '    id_orden_compra = 2
    '    id_factura = 3
    '    id_pedido_cliente = 4
    '    id_dec_dir = 5
    '    id_cdad_bultos = 10
    '    id_tarimas = 11
    '    id_fecha_cita_prog = 13
    '    pestana_encabezados = colSheets.Item(1)
    '    'LTL
    '    id_cclclave = 0
    ''  20220525>>
End Select

If UCase(tipo_carga) = "SIN_FACTURA" Then
   id_fecha_factura = -1
   id_factura = 0
   id_dec_dir = 1
   id_cdad_bultos = 2
   id_tarimas = 3
   id_fecha_cita_prog = 4
   
   ''''LTL
   id_cclclave = -1
   
   pestana_encabezados = colSheets.Item(1)
   
   '''id_orden_compra = 2
   '''id_pedido_cliente = 4
End If

If UCase(tipo_carga) = "CON_FACTURA" Then
   id_fecha_factura = -1
   id_orden_compra = 0
   id_pedido_cliente = 1
   id_factura = 2
   
   'COMPLEMENTO
   id_fac_complemento = 3
   
   id_dec_dir = 4
   id_cdad_bultos = 5
   id_tarimas = 6
   id_fecha_cita_prog = 7
   
   ''''LTL
   id_cclclave = -1
   
   pestana_encabezados = colSheets.Item(1)
   
End If

Call log_SQL("carga_archivo", "campos inicializados", cliente)

'query para recuperar los datos principales
oRS.Open "Select * from [" & pestana_encabezados & "] ", oConn, adOpenStatic, adLockOptimistic


If UCase(tipo_carga) <> "SIN_FACTURA" And UCase(tipo_carga) <> "CON_FACTURA" Then
        Select Case cliente
                Case "3624"
                        'avanzamos de 4 lineas para llegar a los primeros datos:
                        oRS.MoveNext
                        oRS.MoveNext
                        oRS.MoveNext
                        oRS.MoveNext
                
                Case "3885"
                        Do While Not IsNumeric(NVL(oRS.Fields(0)))
                                oRS.MoveNext
                        Loop
                        
                        'Johnson & Johnson 15178 ------------------------------------------->
                Case "15178"
                        'avanza hasta donde haya factura
                        Do While NVL(oRS.Fields(id_factura)) = ""
                                oRS.MoveNext
                        Loop
                        'Johnson & Johnson 15178 ------------------------------------------->

        End Select
Else
    status = ValidarLayOut(UCase(tipo_carga))
End If

'oRS.MoveFirst
'Do While Not oRS.EOF
'    oRS.MoveNext
'Loop

If status = "" Or status = "ok" Then
    Do While Not oRS.EOF
        j = j + 1
    
            If UCase(tipo_carga) <> "SIN_FACTURA" And UCase(tipo_carga) <> "CON_FACTURA" Then
                    If cliente = "3885" Or cliente = "3081" Or cliente = "13128" Then
                            If NVL(oRS.Fields(0)) = "" Then Exit Do
                    End If
                    
                    'Johnson & Johnson 15178 ------------------------------------------->
                    If cliente = "15178" Then
                            If NVL(oRS.Fields(id_factura)) = "" Then Exit Do
                    End If
                    'Johnson & Johnson 15178 ------------------------------------------->
                    
                    '20171122 -- >
                    '20180606 -- > Carga GSK
                    'If cliente = "17873" Then
                    '   20181001 -- > Carga GSK se agregan 20501 y 20502
                    If cliente = "17873" Or cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                    '   20181001 < -- Carga GSK se agregan 20501 y 20502
                    '20180606 < --
                            If NVL(oRS.Fields(id_factura)) = "" Then Exit Do
                    End If
                    '20171122 < --
                    
                    Call log_SQL("carga_archivo", "init verificacion registro", cliente)
                    
                    '20180614 -- >
                    'Si es GSK y carga CD verifica si ya esta registrado el delivery
                    actualiza = 0
                    '   20181001 -- > Carga GSK se agregan 20501 y 20502
                    If (cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502") And tipo_carga = "CD" And NVL(oRS.Fields(id_pedido_cliente)) <> "" Then
                    '   20181001 < -- Carga GSK se agregan 20501 y 20502
                            SQL = "SELECT COUNT(0) " & vbCrLf
                            SQL = SQL & " FROM WCROSS_DOCK WCD " & vbCrLf
                            SQL = SQL & " WHERE WCD.WCDFACTURA = '" & NVL(oRS.Fields(id_factura)) & "'" & vbCrLf
                            SQL = SQL & " AND WCD.WCD_CLICLEF = " & cliente & vbCrLf
                            SQL = SQL & " AND WCD.WCD_PEDIDO_CLIENTE IS NULL " & vbCrLf
                            SQL = SQL & " AND WCD.WCDSTATUS IN (1,2) " & vbCrLf
    
                            rs.Open SQL
                            If Not rs.EOF Then
                                    If rs.Fields(0) = "1" Then
                                            ReDim Preserve talones_actualizar(1, UBoundCheck(talones_actualizar, 2) + 1)
                                            talones_actualizar(0, UBound(talones_actualizar, 2)) = NVL(oRS.Fields(id_factura))
                                            talones_actualizar(1, UBound(talones_actualizar, 2)) = NVL(oRS.Fields(id_pedido_cliente))
                                            actualiza = 1
                                    End If
                            End If
                            rs.Close
                    End If
                    '20180614 < --
            ElseIf UCase(tipo_carga) = "CON_FACTURA" Then
                            If NVL(oRS.Fields(id_factura)) = "" Then
                                    'status = "- No es posible documentar NUI�s si no se indica la factura."
                            Else
                            'COMPLEMENTO
                                If (obtiene_valor_complemento_pedido(NVL(oRS.Fields(id_fac_complemento))) = "N") Then
                                    If existe_factura_cliente(cliente, NVL(oRS.Fields(id_factura))) = True Then
                                        status = "- La factura ya se encuentra asignada a otro NUI."
                                    Else
                                        status = "ok"
                                    End If
                                End If
                                                            
                                    If (obtiene_valor_complemento_pedido(NVL(oRS.Fields(id_fac_complemento))) = "") Then
                                            error_msg = error_msg & "- No se indico si la factura es complemento en la linea " & oRS.AbsolutePosition + 1 & vbCrLf & vbCrLf
                                    End If
                            End If
                    End If
        '<<20240213
    If status = "ok" Then
        status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
        If status <> "ok" Then
            error_msg = error_msg & vbCrLf & status
        End If
    End If
    '  20240213>>
        '20180614 -- >
        ''''''''''''''''''''''''''''''''''''''''''''''''''If status <> "ok" Then
            If actualiza = 0 Then
                    If UCase(tipo_carga) = "SIN_FACTURA" Or UCase(tipo_carga) = "CON_FACTURA" Then
                        If status = "ok" Or status = "" Then
                            my_cclclave = get_cclclave(oRS.Fields(id_dec_dir), cliente)
                            
                            If my_cclclave <> "" Then
                                '''<<20240214: se reutiliza la funcionalidad de CrossDock
                                    'status = get_direccion_entrega_ltl(my_cclclave, num_cclclave, num_allclave_dest)
                                    status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
                                '''  20240214>>
                                If status <> "ok" Then
                                    error_msg = error_msg & vbCrLf & status
                                End If
                            Else
                                status = "- direccion inexistente, o destino INSEGURO, INVALIDO o TIPO DE ENTREGA no autorizado"
                            End If
                        End If
                    End If
                    
            'verificacion de los datos de direccion de clientes
            If tipo_carga = "CD" Then
                If cliente = "13128" Then
                    num_direccion_cliente = oRS.Fields(2) & " " & oRS.Fields(3) & vbCrLf & oRS.Fields(4)
                    status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente, oRS.Fields(id_cil_num_dir))
                ElseIf (cliente <> "3081" And cliente <> "17873") Then
                    'Johnson & Johnson 15178 ------------------------------------------->
                    If cliente = "15178" Then
                        ' Para Johnson no viene la descripcon de la direccion del cliente en el archivo
                        num_direccion_cliente = "N/A"
                    'Johnson & Johnson 15178 ------------------------------------------->
                    '20180606 -- > Carga GSK
                    '   20181001 -- > Carga GSK se agregan 20501 y 20502
                    ElseIf cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                    '   20181001 < -- Carga GSK se agregan 20501 y 20502
                        num_direccion_cliente = oRS.Fields(6) & ", " & oRS.Fields(8) & ", " & oRS.Fields(9)
                    '20180606 < --
                    Else
                        'para los clientes 3885 y 3624 la direccion del cliente esta en los campos 6 y 7
                        num_direccion_cliente = oRS.Fields(6) & vbCrLf & oRS.Fields(7)
                    End If
                    status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
                    
                    If status <> "ok" And status <> "" Then
                        error_msg = error_msg & vbCrLf & status
                    End If
                Else
                    status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
                    
                    If status <> "ok" And status <> "" Then
                        error_msg = error_msg & vbCrLf & status
                    End If
                End If
            Else
                'status = get_direccion_entrega_ltl(oRS.Fields(id_cclclave), num_cclclave, num_allclave_dest)
            End If
            
            Call log_SQL("carga_archivo", "direccion entrega verificada", cliente)
        'End If
        
            If status <> "ok" And status <> "" Then
                'hubo un error al recuperar la direccion
                If UCase(tipo_carga) = "CON_FACTURA" Then
                    error_msg = error_msg & status & "." & "Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                    status = ""
                End If
            End If
        
            
            If UCase(tipo_carga) <> "SIN_FACTURA" Then
                'verificamos que los campos obligatorios esten en el archivo
                'factura:
                If NVL(oRS.Fields(id_factura)) = "" Then
                    error_msg = error_msg & "- No hay numero de factura en la linea " & oRS.AbsolutePosition + 1 & vbCrLf & vbCrLf
                End If
            End If
            
            If NVL(oRS.Fields(id_cdad_bultos)) = "" Then
                error_msg = error_msg & "No hay cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf
            ElseIf Not IsNumeric(NVL(oRS.Fields(id_cdad_bultos))) Then
                error_msg = error_msg & "La cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser numerica. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
            ElseIf CDbl(NVL(oRS.Fields(id_cdad_bultos))) <= 0 Then
                error_msg = error_msg & "La cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser positiva. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
            End If
                    
            '20180606 -- > Carga GSK
            'If cliente <> "3885" Then
            '   20181001 -- > Carga GSK se agregan 20501 y 20502
            If UCase(tipo_carga) <> "SIN_FACTURA" Then
    '                        If cliente <> "3885" And cliente <> "20341" And cliente <> "20305" And cliente <> "20501" And cliente <> "20502" Then
    '                        '   20181001 < -- Carga GSK se agregan 20501 y 20502
    '                        '20180606 < --
    '                                'para este cliente, la cdad de bultos esta en el detalle
    '                                If NVL(oRS.Fields(id_cdad_bultos)) = "" Then
    '                                        error_msg = error_msg & "No hay cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf
    '                                ElseIf Not IsNumeric(NVL(oRS.Fields(id_cdad_bultos))) Then
    '                                        error_msg = error_msg & "La cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser numerica. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf
    '                                ElseIf CDbl(NVL(oRS.Fields(id_cdad_bultos))) <= 0 Then
    '                                        error_msg = error_msg & "La cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser positiva. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf
    '                                End If
    '                        End If
    
                            '20180606 -- > Carga GSK
                            '   20181001 -- > Carga GSK se agregan 20501 y 20502
                            If cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Or cliente = "23488" Or cliente = "23489" Then
                            '   20181001 < -- Carga GSK se agregan 20501 y 20502
                                    If NVL(oRS.Fields(id_cdad_bultos)) = "" Then
                                            error_msg = error_msg & "No hay cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    ElseIf Not IsNumeric(NVL(oRS.Fields(id_cdad_bultos))) Then
                                            error_msg = error_msg & "La cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser numerica. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    ElseIf CDbl(NVL(oRS.Fields(id_cdad_bultos))) <= 0 Then
                                            error_msg = error_msg & "La cdad de bultos en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser positiva. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    End If
                                    
                                    'Valida tarimas
                                    If id_tarimas > -1 Then
                                            If NVL(oRS.Fields(id_tarimas)) <> "" And IsNumeric(NVL(oRS.Fields(id_tarimas))) Then
                                                    If CDbl(NVL(oRS.Fields(id_tarimas))) <= 0 Then
                                                            error_msg = error_msg & "La cdad de tarimas en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser positiva. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                                    End If
                                            End If
                                    End If
                            End If
                            '20180606 < --
            End If
                    
            Call log_SQL("carga_archivo", "bultos verificados", cliente)
        
            If status = "ok" Then
                If UCase(tipo_carga) <> "SIN_FACTURA" Then
                    If tipo_carga = "CD" Then
                        numero_factura = NVL(oRS.Fields(id_factura))
                        status = get_factura_duplicada(NVL(oRS.Fields(id_factura)), cliente)
                        Call log_SQL("carga_archivo", "factura duplicada verificada " & numero_factura, cliente)
                    ElseIf UCase(tipo_carga) = "CON_FACTURA" Then
                        numero_factura = NVL(oRS.Fields(id_factura))
                                            
                        'COMPLEMENTO
                        num_fac_complemento = obtiene_valor_complemento_pedido(NVL(oRS.Fields(id_fac_complemento)))
                                            
                    End If
                End If
                
                If status <> "ok" Then
                    'hubo un error al recuperar las facturas duplicadas
                    error_msg = error_msg & status & vbCrLf & vbCrLf
                End If
            End If
            
            
            
            
            
                    If UCase(tipo_carga) <> "SIN_FACTURA" Then
                            If id_peso > -1 Then
                                    'verificacion del peso
                                    If NVL(oRS.Fields(id_peso)) = "" Then
                                            error_msg = error_msg & "No hay peso en la linea " & oRS.AbsolutePosition + 1 & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    ElseIf Not IsNumeric(NVL(oRS.Fields(id_peso))) Then
                                            error_msg = error_msg & "El peso en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser numerico. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    ElseIf CDbl(NVL(oRS.Fields(id_peso))) <= 0 Then
                                            error_msg = error_msg & "El peso en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser positivo. " & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    End If
                            Call log_SQL("carga_archivo", "peso verificado " & numero_factura, cliente)
                                                    End If
                    End If
        
                    If UCase(tipo_carga) = "SIN_FACTURA" Then
                            If cliente = "3882" Then
                                    'verificacion del collect_prepaid
                                    If NVL(oRS.Fields(id_collect_prepaid)) = "" Then
                                            error_msg = error_msg & "No hay tipo de LTL en la linea " & oRS.AbsolutePosition + 1 & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    ElseIf NVL(oRS.Fields(id_collect_prepaid)) <> "FLETE PAGADO" And NVL(oRS.Fields(id_collect_prepaid)) <> "FLETE POR COBRAR" Then
                                            error_msg = error_msg & "El tipo de LTL en la linea " & oRS.AbsolutePosition + 1 & " tiene que ser FLETE PAGADO o FLETE POR COBRAR." & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    End If
                                    
                                    If num_cclclave > 0 And NVL(oRS.Fields(id_collect_prepaid)) = "FLETE POR COBRAR" Then
                                            'verificacion de las LTLs por cobrar
                                            
                                            '<<<--CHG-DESA-27022024-01
                                                                                        ''' SQL = "SELECT COUNT(0) " & vbCrLf
                                            ''' SQL = SQL & " FROM WEB_CLIENT_CLIENTE WCCL " & vbCrLf
                                            ''' SQL = SQL & " WHERE WCCL.WCCLCLAVE = " & num_cclclave & " " & vbCrLf
                                            ''' SQL = SQL & " AND NOT EXISTS ( " & vbCrLf
                                            ''' SQL = SQL & " SELECT NULL FROM ECLIENT WHERE CLIRFC = WCCL_RFC) " & vbCrLf
                                                                                        
                                                                                        
                                            SQL = " SELECT COUNT(0) " & vbCrLf
                                            SQL = SQL & " FROM ECLIENT_CLIENTE CCL " & vbCrLf
                                            SQL = SQL & " WHERE 1=1 " & vbCrLf
                                            SQL = SQL & " AND CCL.CCLCLAVE = " & num_cclclave & "  " & vbCrLf
                                            SQL = SQL & " AND NOT EXISTS ( SELECT NULL FROM ECLIENT WHERE CLIRFC = CCL_RFC) " & vbCrLf
                                            'CHG-DESA-27022024-01-->>>
                                                                                        
                            
                                            rs.Open SQL
                                            If Not rs.EOF Then
                                                    If rs.Fields(0) <> "0" Then
                                                            error_msg = error_msg & "El talon POR COBRAR no se puede crear porque el destinatario de id " & oRS.Fields(4) & " y RFC " & oRS.Fields(5) & " no esta dado de alta como cliente Orfeo!!" & vbCrLf & vbCrLf
                                                    End If
                                            End If
                                            rs.Close
                            
                                                                                
                                            '<<<--CHG-DESA-27022024-01
                                            'estamos en POR COBRAR, checamos que el destinatario no este ni en el DF ni en el estado de mexico 2
                                            ''' SQL = "SELECT COUNT(0) " & vbCrLf
                                            ''' SQL = SQL & " FROM WEB_CLIENT_CLIENTE WCCL " & vbCrLf
                                            ''' SQL = SQL & " WHERE WCCL.WCCLCLAVE = " & num_cclclave & " " & vbCrLf
                                            ''' SQL = SQL & " AND EXISTS ( " & vbCrLf
                                            ''' SQL = SQL & " SELECT NULL FROM ECIUDADES, EESTADOS WHERE VILCLEF = WCCL_VILLE AND ESTESTADO = VIL_ESTESTADO " & vbCrLf
                                            ''' SQL = SQL & " AND ESTESTADO IN (1129, 1444) ) " & vbCrLf
                                            ''' SQL = SQL & " AND NOT EXISTS ( " & vbCrLf
                                            ''' SQL = SQL & " SELECT NULL FROM ECLIENT, ECREDIT WHERE CLIRFC = WCCL_RFC " & vbCrLf
                                            ''' SQL = SQL & " AND CRECLIENT = CLICLEF AND CREREGIME = 2 ) " & vbCrLf
                                    
                                                                        
                                            SQL = " SELECT COUNT(0) " & vbCrLf
                                            SQL = SQL & " FROM ECLIENT_CLIENTE CCL " & vbCrLf
                                            SQL = SQL & " WHERE 1=1" & vbCrLf
                                            SQL = SQL & " AND CCL.CCLCLAVE =  " & num_cclclave & " " & vbCrLf
                                            SQL = SQL & " AND EXISTS ( " & vbCrLf
                                            SQL = SQL & " SELECT NULL FROM ECIUDADES, EESTADOS WHERE VILCLEF = CCL_VILLE AND ESTESTADO = VIL_ESTESTADO " & vbCrLf
                                            SQL = SQL & " AND ESTESTADO IN (1129, 1444) ) " & vbCrLf
                                            SQL = SQL & " AND NOT EXISTS ( " & vbCrLf
                                            SQL = SQL & " SELECT NULL FROM ECLIENT, ECREDIT WHERE CLIRFC = CCL_RFC " & vbCrLf
                                            SQL = SQL & " AND CRECLIENT = CLICLEF AND CREREGIME = 2 ) " & vbCrLf
                                            'CHG-DESA-27022024-01-->>>

                                                                        
                                            rs.Open SQL
                                            If Not rs.EOF Then
                                                    If rs.Fields(0) <> "0" Then
                                                            error_msg = error_msg & "El talon POR COBRAR no se puede crear porque el destinatario de id " & oRS.Fields(4) & " y RFC " & oRS.Fields(5) & " es del Distrito Federal o sus cercanias!!" & vbCrLf & vbCrLf
                                                    End If
                                            End If
                                            rs.Close
                                    End If
                                    Call log_SQL("carga_archivo", "tipo talon verificado " & numero_factura, cliente)
                            End If
                    
            
                            If cliente = "3882" Or cliente = "3885" Then
                                    'para Phillips verificamos el detalle de bultos
                                    If cliente = "3882" Then
                                            oRS2.Open "Select * from [" & pestana_detalles & "] where F1 =""" & oRS.Fields(id_factura) & """ and F2=""" & oRS.Fields(id_cclclave) & """", oConn, adOpenStatic, adLockOptimistic
                                    Else
                                            oRS2.Open "Select * from [" & pestana_detalles & "] where F1 =""" & oRS.Fields(id_factura) & """", oConn, adOpenStatic, adLockOptimistic
                                    End If
                                    
                                    If oRS2.EOF Then
                                            error_msg = error_msg & "No hay detalle de referencias para la factura " & oRS.Fields(id_factura) & ". " & vbCrLf & vbCrLf
                                    End If
                                    
                                    Do While Not oRS2.EOF
                                            status = get_referencia(cliente, oRS2.Fields(id_ref), num_eirclef)
                                            
                                            If cliente = "3885" And status <> "ok" Then
                                                    SQL = "INSERT INTO EINVENTARIO_REFERENCIA " & vbCrLf
                                                    SQL = SQL & " (EIRCLEF, EIR_CLICLEF, EIRREFERENCIA, EIRNOM, EIR_STATUS, CREATED_BY, DATE_CREATED) " & vbCrLf
                                                    SQL = SQL & " VALUES " & vbCrLf
                                                    SQL = SQL & " (SEQ_EINVENTARIO_REFERENCIA.nextval, 3885, SUBSTR('" & Replace(oRS2.Fields(id_ref), "'", "''") & "', 1, 20), 'PENDIENTE', 'ACTIVA', USER || '-CG_WEB', SYSDATE) "
                                                    rs.Open SQL
                                                    
                                            ElseIf status <> "ok" Then
                                                    'hubo un error al recuperar las referencias
                                                    error_msg = error_msg & status & vbCrLf & vbCrLf
                                            End If
                                            
                                            If NVL(oRS2.Fields(id_eirclef_cdad)) = "" Then
                                                    error_msg = error_msg & "No hay cdad de referencia en la linea " & oRS2.AbsolutePosition & " de la factura " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                            ElseIf Not IsNumeric(NVL(oRS2.Fields(id_eirclef_cdad))) Then
                                                    error_msg = error_msg & "La cdad de referencia en la linea " & oRS2.AbsolutePosition & " de la factura " & NVL(oRS.Fields(id_factura)) & " tiene que ser numerica. " & vbCrLf & vbCrLf
                                            ElseIf CDbl(NVL(oRS2.Fields(id_eirclef_cdad))) <= 0 Then
                                                    error_msg = error_msg & "cdad de referencia en la linea " & oRS2.AbsolutePosition & " de la factura " & NVL(oRS.Fields(id_factura)) & " tiene que ser positiva. " & vbCrLf & vbCrLf
                                            End If
                                            
                                            oRS2.MoveNext
                                    Loop
                                    oRS2.Close
                                    
                                    Call log_SQL("carga_archivo", "referencias verificadas " & numero_factura, cliente)
                            End If
                    End If
        
            'verificacion de los FORANEO 5, no aplican para L'Oreal
            'If cliente <> "3110" Then
            If num_dieclave > 0 Or num_dieclave_entrega > 0 Then  'entonces existe una direccion de entrega
                SQL = ""
                            If UCase(tipo_carga) = "SIN_FACTURA" Then
                                                        
                    '<<<--CHG-DESA-27022024-01
                    ''' SQL = "SELECT 1 " & vbCrLf
                    ''' SQL = SQL & " FROM WEB_CLIENT_CLIENTE " & vbCrLf
                    ''' SQL = SQL & " , EALMACENES_LOGIS EAL " & vbCrLf
                    ''' SQL = SQL & " , EDESTINOS_POR_RUTA " & vbCrLf
                    ''' SQL = SQL & " WHERE WCCLCLAVE = " & num_cclclave & vbCrLf
                    ''' SQL = SQL & "   AND ALLCLAVE = " & num_allclave_ori & vbCrLf
                    ''' SQL = SQL & "   AND DER_VILCLEF(+) = WCCL_VILLE " & vbCrLf
                    ''' SQL = SQL & "   AND (NVL(DER_TIPO_ENTREGA, 'FORANEO 5') <> 'FORANEO 5' " & vbCrLf
                    ''' SQL = SQL & "   OR EXISTS " & vbCrLf
                    ''' SQL = SQL & "   ( " & vbCrLf
                    ''' SQL = SQL & "       SELECT /*+ ORDERED USE_NL(CCO) */ NULL " & vbCrLf
                    ''' SQL = SQL & "       FROM EBASES_POR_CONCEPT BPC, " & vbCrLf
                    ''' SQL = SQL & "       ECLIENT_APLICA_CONCEPTOS CCO, " & vbCrLf
                    ''' SQL = SQL & "       EPARAMETRO_RESTRICT PAR, " & vbCrLf
                    ''' SQL = SQL & "       ECONCEPTOSHOJA " & vbCrLf
                    ''' SQL = SQL & "       WHERE BPC_CHOCLAVE = CHOCLAVE " & vbCrLf
                    ''' SQL = SQL & "       AND CHONUMERO IN (172)  " & vbCrLf
                    ''' SQL = SQL & "       AND CHOTIPOIE = 'I'  " & vbCrLf
                    ''' SQL = SQL & "       AND CCO_BPCCLAVE = BPCCLAVE " & vbCrLf
                    ''' SQL = SQL & "       AND CCO_CLICLEF = " & cliente & vbCrLf
                    ''' SQL = SQL & "       AND PARCLAVE = BPC_PARCLAVE " & vbCrLf
                    ''' SQL = SQL & "       AND PAR_VILCLEF_ORI = EAL.ALL_VILCLEF " & vbCrLf
                    ''' SQL = SQL & "       AND PAR_VILCLEF_DEST = WCCL_VILLE " & vbCrLf
                    ''' SQL = SQL & "       AND ROWNUM = 1 " & vbCrLf
                    ''' SQL = SQL & "   ) " & vbCrLf
                    ''' SQL = SQL & "   )  "
                                        
                                        
                    SQL = " SELECT 1 " & vbCrLf
                    SQL = SQL & " FROM ECLIENT_CLIENTE " & vbCrLf
                    SQL = SQL & " , EALMACENES_LOGIS EAL  " & vbCrLf
                    SQL = SQL & " , EDESTINOS_POR_RUTA  " & vbCrLf
                    SQL = SQL & " WHERE 1=1" & vbCrLf
                    SQL = SQL & "    AND CCLCLAVE = " & num_cclclave & vbCrLf
                    SQL = SQL & "    AND ALLCLAVE = " & num_allclave_ori & vbCrLf
                    SQL = SQL & "   AND DER_VILCLEF(+) = CCL_VILLE  " & vbCrLf
                    SQL = SQL & "   AND (NVL(DER_TIPO_ENTREGA, 'FORANEO 5') <> 'FORANEO 5'  " & vbCrLf
                    SQL = SQL & "   OR EXISTS  " & vbCrLf
                    SQL = SQL & "   (  " & vbCrLf
                    SQL = SQL & "     SELECT /*+ ORDERED USE_NL(CCO) */ NULL  " & vbCrLf
                    SQL = SQL & "     FROM EBASES_POR_CONCEPT BPC,  " & vbCrLf
                    SQL = SQL & "     ECLIENT_APLICA_CONCEPTOS CCO,  " & vbCrLf
                    SQL = SQL & "     EPARAMETRO_RESTRICT PAR,  " & vbCrLf
                    SQL = SQL & "     ECONCEPTOSHOJA  " & vbCrLf
                    SQL = SQL & "     WHERE BPC_CHOCLAVE = CHOCLAVE  " & vbCrLf
                    SQL = SQL & "     AND CHONUMERO IN (172)   " & vbCrLf
                    SQL = SQL & "     AND CHOTIPOIE = 'I'   " & vbCrLf
                    SQL = SQL & "     AND CCO_BPCCLAVE = BPCCLAVE  " & vbCrLf
                    SQL = SQL & "     AND CCO_CLICLEF = " & cliente & vbCrLf
                    SQL = SQL & "     AND PARCLAVE = BPC_PARCLAVE  " & vbCrLf
                    SQL = SQL & "     AND PAR_VILCLEF_ORI = EAL.ALL_VILCLEF  " & vbCrLf
                    SQL = SQL & "     AND PAR_VILCLEF_DEST = CCL_VILLE  " & vbCrLf
                    SQL = SQL & "     AND ROWNUM = 1  " & vbCrLf
                    SQL = SQL & "   )  " & vbCrLf
                    SQL = SQL & "   )" & vbCrLf
                    'CHG-DESA-27022024-01-->>>
                                        
                                        
                            Else
                                    If tipo_carga = "CD" Or Trim(UCase(tipo_carga)) = "CON_FACTURA" Then
                                                    SQL = "SELECT 1 " & vbCrLf
                                                    SQL = SQL & " FROM EDIRECCIONES_ENTREGA " & vbCrLf
                                                    SQL = SQL & " , EALMACENES_LOGIS EAL " & vbCrLf
                                                    SQL = SQL & " , EDESTINOS_POR_RUTA " & vbCrLf
                                                    SQL = SQL & " WHERE DIECLAVE = " & IIf(num_dieclave_entrega > 0, num_dieclave_entrega, num_dieclave) & vbCrLf
                                                    SQL = SQL & "   AND ALLCLAVE = " & num_allclave_ori & vbCrLf
                                                    SQL = SQL & "   AND DER_VILCLEF(+) = DIEVILLE " & vbCrLf
                                                    SQL = SQL & "   AND (NVL(DER_TIPO_ENTREGA, 'FORANEO 5') <> 'FORANEO 5' " & vbCrLf
                                                    SQL = SQL & "   OR EXISTS " & vbCrLf
                                                    SQL = SQL & "   ( " & vbCrLf
                                                    SQL = SQL & "       SELECT /*+ ORDERED USE_NL(CCO) */ NULL " & vbCrLf
                                                    SQL = SQL & "       FROM EBASES_POR_CONCEPT BPC, " & vbCrLf
                                                    SQL = SQL & "       ECLIENT_APLICA_CONCEPTOS CCO, " & vbCrLf
                                                    SQL = SQL & "       EPARAMETRO_RESTRICT PAR, " & vbCrLf
                                                    SQL = SQL & "       ECONCEPTOSHOJA " & vbCrLf
                                                    SQL = SQL & "       WHERE BPC_CHOCLAVE = CHOCLAVE " & vbCrLf
                                                    SQL = SQL & "       AND CHONUMERO IN (40, 240)  " & vbCrLf
                                                    SQL = SQL & "       AND CHOTIPOIE = 'I'  " & vbCrLf
                                                    SQL = SQL & "       AND CCO_BPCCLAVE = BPCCLAVE " & vbCrLf
                                                    SQL = SQL & "       AND CCO_CLICLEF = " & cliente & vbCrLf
                                                    SQL = SQL & "       AND PARCLAVE = BPC_PARCLAVE " & vbCrLf
                                                    SQL = SQL & "       AND PAR_VILCLEF_ORI = EAL.ALL_VILCLEF " & vbCrLf
                                                    SQL = SQL & "       AND PAR_VILCLEF_DEST = DIEVILLE " & vbCrLf
                                                    SQL = SQL & "       AND ROWNUM = 1 " & vbCrLf
                                                    SQL = SQL & "   ) " & vbCrLf
                                                    SQL = SQL & "   ) "
                                    Else
                                            '<<20240214:
                                            'If num_cclclave <> -1 And num_allclave_ori <> -1 Then
                                            '  20240214>>
                                                                                        
                                                                                                        '<<<--CHG-DESA-27022024-01
                                                    ''' SQL = "SELECT 1 " & vbCrLf
                                                    ''' SQL = SQL & " FROM WEB_CLIENT_CLIENTE " & vbCrLf
                                                    ''' SQL = SQL & " , EALMACENES_LOGIS EAL " & vbCrLf
                                                    ''' SQL = SQL & " , EDESTINOS_POR_RUTA " & vbCrLf
                                                    ''' SQL = SQL & " WHERE WCCLCLAVE = " & num_cclclave & vbCrLf
                                                    ''' SQL = SQL & "   AND ALLCLAVE = " & num_allclave_ori & vbCrLf
                                                    ''' SQL = SQL & "   AND DER_VILCLEF(+) = WCCL_VILLE " & vbCrLf
                                                    ''' SQL = SQL & "   AND (NVL(DER_TIPO_ENTREGA, 'FORANEO 5') <> 'FORANEO 5' " & vbCrLf
                                                    ''' SQL = SQL & "   OR EXISTS " & vbCrLf
                                                    ''' SQL = SQL & "   ( " & vbCrLf
                                                    ''' SQL = SQL & "       SELECT /*+ ORDERED USE_NL(CCO) */ NULL " & vbCrLf
                                                    ''' SQL = SQL & "       FROM EBASES_POR_CONCEPT BPC, " & vbCrLf
                                                    ''' SQL = SQL & "       ECLIENT_APLICA_CONCEPTOS CCO, " & vbCrLf
                                                    ''' SQL = SQL & "       EPARAMETRO_RESTRICT PAR, " & vbCrLf
                                                    ''' SQL = SQL & "       ECONCEPTOSHOJA " & vbCrLf
                                                    ''' SQL = SQL & "       WHERE BPC_CHOCLAVE = CHOCLAVE " & vbCrLf
                                                    ''' SQL = SQL & "       AND CHONUMERO IN (172)  " & vbCrLf
                                                    ''' SQL = SQL & "       AND CHOTIPOIE = 'I'  " & vbCrLf
                                                    ''' SQL = SQL & "       AND CCO_BPCCLAVE = BPCCLAVE " & vbCrLf
                                                    ''' SQL = SQL & "       AND CCO_CLICLEF = " & cliente & vbCrLf
                                                    ''' SQL = SQL & "       AND PARCLAVE = BPC_PARCLAVE " & vbCrLf
                                                    ''' SQL = SQL & "       AND PAR_VILCLEF_ORI = EAL.ALL_VILCLEF " & vbCrLf
                                                    ''' SQL = SQL & "       AND PAR_VILCLEF_DEST = WCCL_VILLE " & vbCrLf
                                                    ''' SQL = SQL & "       AND ROWNUM = 1 " & vbCrLf
                                                    ''' SQL = SQL & "   ) " & vbCrLf
                                                    ''' SQL = SQL & "   )  "
                                                                                                        
                                                                                                        
                                                                                                        
                                                    SQL = " SELECT 1  " & vbCrLf
                                                    SQL = SQL & " FROM ECLIENT_CLIENTE  " & vbCrLf
                                                    SQL = SQL & " , EALMACENES_LOGIS EAL  " & vbCrLf
                                                    SQL = SQL & " , EDESTINOS_POR_RUTA  " & vbCrLf
                                                    SQL = SQL & " WHERE 1=1" & vbCrLf
                                                    SQL = SQL & "  AND WCCLCLAVE = " & num_cclclave & vbCrLf
                                                    SQL = SQL & "  AND ALLCLAVE = " & num_allclave_ori & vbCrLf
                                                    SQL = SQL & "  AND DER_VILCLEF(+) = CCL_VILLE  " & vbCrLf
                                                    SQL = SQL & "  AND (NVL(DER_TIPO_ENTREGA, 'FORANEO 5') <> 'FORANEO 5'  " & vbCrLf
                                                    SQL = SQL & "  OR EXISTS  " & vbCrLf
                                                    SQL = SQL & "  (  " & vbCrLf
                                                    SQL = SQL & "    SELECT /*+ ORDERED USE_NL(CCO) */ NULL  " & vbCrLf
                                                    SQL = SQL & "    FROM EBASES_POR_CONCEPT BPC,  " & vbCrLf
                                                    SQL = SQL & "    ECLIENT_APLICA_CONCEPTOS CCO,  " & vbCrLf
                                                    SQL = SQL & "    EPARAMETRO_RESTRICT PAR,  " & vbCrLf
                                                    SQL = SQL & "    ECONCEPTOSHOJA  " & vbCrLf
                                                    SQL = SQL & "    WHERE BPC_CHOCLAVE = CHOCLAVE  " & vbCrLf
                                                    SQL = SQL & "    AND CHONUMERO IN (172)   " & vbCrLf
                                                    SQL = SQL & "    AND CHOTIPOIE = 'I'   " & vbCrLf
                                                    SQL = SQL & "    AND CCO_BPCCLAVE = BPCCLAVE  " & vbCrLf
                                                    SQL = SQL & "    AND CCO_CLICLEF = " & cliente & vbCrLf
                                                    SQL = SQL & "    AND PARCLAVE = BPC_PARCLAVE  " & vbCrLf
                                                    SQL = SQL & "    AND PAR_VILCLEF_ORI = EAL.ALL_VILCLEF  " & vbCrLf
                                                    SQL = SQL & "    AND PAR_VILCLEF_DEST = CCL_VILLE  " & vbCrLf
                                                    SQL = SQL & "    AND ROWNUM = 1  " & vbCrLf
                                                    SQL = SQL & "  )  " & vbCrLf
                                                    SQL = SQL & "  )  " & vbCrLf
                                                    'CHG-DESA-27022024-01-->>>
                                                                                                        
                                            '<<20240214:
                                            'End If
                                            '  20240214>>
                                    End If
                            End If
                            
                            If SQL <> "" Then
                                    rs.Open SQL
                                    If rs.EOF Then
                                            error_msg = error_msg & "- la direccion de entrega de la linea " & oRS.AbsolutePosition + 1 & " es un FORANEO 5 y no tiene tarifa asociada." & " Factura: " & NVL(oRS.Fields(id_factura)) & vbCrLf & vbCrLf
                                    End If
                                    rs.Close
                            Call log_SQL("carga_archivo", "tarifa verificada " & numero_factura, cliente)
                            End If
                    End If
            'End If
            oRS.MoveNext
            
            Call log_SQL("carga_archivo", "registro verificado " & numero_factura, cliente)
        Else
            Call log_SQL("carga_archivo", "registro agregado para actualizacion " & NVL(oRS.Fields(id_factura)), cliente)
            oRS.MoveNext
        End If
        '20180614 < --
    Loop 'TERMINAN VALIDACIONES
End If

If status <> "ok" Then
    error_msg = error_msg & vbCrLf & status
End If
     
numero_factura = ""
Call log_SQL("carga_archivo", "registros verificados", cliente)
     
If error_msg <> "" Then
    'hubo error, mandamos el correo de notificacion
    jmail.From = mail_From
    jmail.FromName = mail_FromName
    jmail.ClearRecipients
    
    'para debug, estoy en los contactos ;)
    jmail.AddRecipientBCC mail_grupo_error(0)
    
    For i = 0 To UBound(Split(Replace(correo_electronico, ",", ";"), ";"))
        jmail.AddRecipient Trim(Split(Replace(correo_electronico, ",", ";"), ";")(i))
    Next
    
    
    '<---- CC-CHG-DESA-13032024-01: Se integra el grupo cargamasiva_smo@logis.com.mx a peticion del usuario
    jmail.AddRecipient "cargamasiva_smo@logis.com.mx"
    'CC-CHG-DESA-13032024-01 ---->
    
    
    jmail.subject = "Error carga de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
    jmail.body = "Hola, hubo un error al cargar el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", favor de revisar los errores y volver a cargarlo." & vbCrLf & vbCrLf & error_msg
    
        If FSO.FileExists(Archivo) Then
        jmail.AddAttachment Archivo
        End If
        
        jmail.Send mail_server
    Exit Sub
End If

oRS.MoveFirst
'avanzamos de 4 lineas para llegar a los primeros datos:
'configuracion de campos
Select Case cliente
    Case "3624"
        'avanzamos de 4 lineas para llegar a los primeros datos:
        oRS.MoveNext
        oRS.MoveNext
        oRS.MoveNext
        oRS.MoveNext
        
    Case "3885"
        Do While Not IsNumeric(NVL(oRS.Fields(0)))
            oRS.MoveNext
        Loop

        'Johnson & Johnson 15178 ------------------------------------------->
    Case "15178"
        'Avanza hasta que haya factura
        Do While NVL(oRS.Fields(id_factura)) = ""
            oRS.MoveNext
        Loop
        'Johnson & Johnson 15178 ------------------------------------------->
End Select


Call log_SQL("carga_archivo", "listo para inserciones", cliente)

'Johnson & Johnson 15178 ------------------------------------------->
num_orden_compra = ""
'Johnson & Johnson 15178 ------------------------------------------->

Do While Not oRS.EOF
        
        If UCase(tipo_carga) <> "SIN_FACTURA" Then
                If cliente = "3885" Or cliente = "3081" Or cliente = "13128" Then
                        If NVL(oRS.Fields(0)) = "" Then Exit Do
                End If
                'Johnson & Johnson 15178 ------------------------------------------->
                'If cliente = "15178" Then
                '   20181001 -- > Carga GSK se agregan 20501 y 20502
                                If cliente = "15178" Or cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                '   20181001 < -- Carga GSK se agregan 20501 y 20502
                        If NVL(oRS.Fields(id_factura)) = "" Then Exit Do
                End If
                'Johnson & Johnson 15178 ------------------------------------------->
                
                '20171122 -- >
                If cliente = "17873" Then
                        If NVL(oRS.Fields(id_factura)) = "" Then Exit Do
                End If
                '20171122 < --
        End If

    'reinicializacion de campos:
    num_factura = ""
    num_pedido_cliente = ""
    num_fecha_factura = "NULL"
    'Johnson & Johnson 15178 ------------------------------------------->
    If cliente <> "15178" Then
        num_orden_compra = ""
    End If
    'Johnson & Johnson 15178 ------------------------------------------->
    num_cdad_bultos = -1
    num_importe = ""
    num_cclclave = -1
    num_dieclave = -1
    num_dieclave_entrega = -1
    num_allclave_ori = -1
    num_allclave_dest = -1
    num_peso = 0
    num_volumen = 0
    num_collect_prepaid = ""
    num_ref = ""
    num_eirclef = -1
    num_observacion = ""
    'Johnson & Johnson 15178 ------------------------------------------->
    num_fecha_cita_prog = ""
    'Johnson & Johnson 15178 ------------------------------------------->
    '20180606 -- > GSK
    num_cdad_empaques_x_bulto = ""
    num_tpaclave = ""
    num_bulto_tpaclave = ""
    '20181019 -- >
    num_wcd_cajas_tarimas = 0
    num_wcd_cdad_tarimas = 0
    num_wcd_cdad_cajas = 0
    '20181019 < --
    '20180606 < --
                
                
        'COMPLEMENTO
        If id_fac_complemento <> -1 Then
            num_fac_complemento = obtiene_valor_complemento_pedido(NVL(oRS.Fields(id_fac_complemento)))
        End If
        
    
        If UCase(tipo_carga) = "SIN_FACTURA" Then
                If NVL(oRS.Fields(id_factura)) <> "" Then
                    posicion = IsInArray_Multi(NVL(oRS.Fields(id_factura)), talones_actualizar, 0)
                    If posicion = -1 Then
                        Call log_SQL("carga_archivo", "inicio nuevo registro", cliente)
                        Call get_remitente(cliente, mi_disclef, num_allclave_ori)
                        Call log_SQL("carga_archivo", "remitente listo " & mi_disclef, cliente)
                    End If
                End If
                
                Call log_SQL("carga_archivo", "inicio nuevo registro", cliente)
                Call get_remitente(cliente, mi_disclef, num_allclave_ori)
                Call log_SQL("carga_archivo", "remitente listo " & mi_disclef, cliente)
                
                If id_factura > -1 Then num_factura = Replace(NVL(oRS.Fields(id_factura)), "'", "''")
                If id_cdad_bultos > -1 And cliente <> "3885" Then num_cdad_bultos = NVL(oRS.Fields(id_cdad_bultos))
                If cliente_con_seguro(cliente) = True Then
                        num_importe = 1
                End If
                num_collect_prepaid = "PREPAGADO"
                num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & "', 'DD/MM/YYYY')"
                If id_tarimas > -1 Then
                        If IsNumeric(NVL(oRS.Fields(id_tarimas))) Then
                                If CLng(NVL(oRS.Fields(id_tarimas))) >= 1 Then
                                        num_cdad_tarimas = CLng(NVL(oRS.Fields(id_tarimas)))
                                        num_cdad_cajas_tarima = NVL(oRS.Fields(id_cdad_bultos))
                                Else
                                        num_cdad_bultos = NVL(oRS.Fields(id_cdad_bultos))
                                End If
                        End If
                Else
                        num_cdad_bultos = NVL(oRS.Fields(id_cdad_bultos))
                End If
                
                my_cclclave = get_cclclave(oRS.Fields(id_dec_dir), cliente)
                                
                                '<<CHG-DESA-27022024-01: Se obtienen la CCLCLAVE y la DIECLAVE a partir del dato del Layout.
                                Call get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest)
                                If Obtiene_DIECLAVE_CCLCLAVE(oRS.Fields(id_dec_dir), cliente, num_dieclave, num_cclclave) = False Then
                                        status = "- direccion inexistente"
                                Else
                                        If Valida_CCLCLAVE(num_cclclave) = True Then
                                        '    'ORP
                                        '    'num_dieclave = -1
                                        '    If num_dieclave = "" Then
                                        '        num_dieclave = -1
                                        '    End If
                                        'Else
                                        '        num_cclclave = -1
                                        End If
                                End If
                                
                '       If my_cclclave <> "" Then
                '           '''<<20240214: se reutiliza la funcionalidad de CrossDock
                '               'Call get_direccion_entrega_ltl(my_cclclave, num_cclclave, num_allclave_dest)
                '               status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
                '           '''  20240214>>
                '        Else
                '           status = "- direccion inexistente, o destino INSEGURO, INVALIDO o TIPO DE ENTREGA no autorizado"
                '       End If
                                '  CHG-DESA-27022024-01>>
                
                If cliente = "3882" Then
                        'para Phillips tenemos que guardar las facturas en el campo de observacion
                        num_observacion = "Factura(s): " & Replace(oRS.Fields(2), "'", "''") & vbCrLf & "CONTENIDO FERRETERIA"
                        
                        'agregamos el destinatario a la factura
                        num_factura = num_factura & "-" & oRS.Fields(id_cclclave)
                End If
                
                Call log_SQL("carga_archivo", "preparacion insercion talon", cliente)
                
                If status <> "ok" Then
                    If num_factura <> "" Then
                        status = " - " & num_factura & " " & status
                    End If
                    Exit Do
                End If
                'Primero busco el siguiente NUI disponible:
                SQL = "SELECT MIN(WELCLAVE) FROM WEB_LTL WHERE WELSTATUS = 3 AND WEL_CLICLEF = '" & cliente & "'" & vbCrLf
                rs.Open SQL
                clef = rs.Fields(0)
                rs.Close
                
                Debug.Print "NUI > " & clef
                
                'Ahora se actualiza la informaci�n recibida:
                SQL = " UPDATE   WEB_LTL " & vbCrLf
                SQL = SQL & "   SET  WEL_ALLCLAVE_ORI       =   '" & num_allclave_ori & "' " & vbCrLf
                SQL = SQL & "       ,WEL_DISCLEF            =   '" & num_disclef & "' " & vbCrLf
                SQL = SQL & "       ,WEL_ALLCLAVE_DEST      =   '" & num_allclave_dest & "' " & vbCrLf
                If num_cclclave <> -1 Then
                    '<<<--CHG-DESA-27022024-01
                    'SQL = SQL & "       ,WEL_WCCLCLAVE          =   '" & num_cclclave & "' " & vbCrLf
                    SQL = SQL & "       ,WEL_CCLCLAVE          =   '" & num_cclclave & "' " & vbCrLf
                    'CHG-DESA-27022024-01-->>>
                End If
                '<<CHG-DESA-27022024-01
                If num_dieclave <> -1 Then
                    SQL = SQL & "       ,WEL_DIECLAVE          =   '" & num_dieclave & "' " & vbCrLf
                End If
                'CHG-DESA-27022024-01>>
                If num_fecha_cita_prog = "" Then
                    SQL = SQL & "       ,WEL_FECHA_RECOLECCION  =    " & "NULL" & " " & vbCrLf
                    SQL = SQL & "       ,WELRECOL_DOMICILIO     =   '" & "N" & "' " & vbCrLf
                Else
                    SQL = SQL & "       ,WEL_FECHA_RECOLECCION  =    " & num_fecha_cita_prog & " " & vbCrLf
                    SQL = SQL & "       ,WELRECOL_DOMICILIO     =   '" & "S" & "' " & vbCrLf
                End If
                If num_factura <> "" Then
                        SQL = SQL & "       ,WELFACTURA             =   SUBSTR('" & num_factura & "',1,99) " & vbCrLf
                Else
                        SQL = SQL & "       ,WELFACTURA             =   '_PENDIENTE_' " & vbCrLf
                End If
                
                If num_pedido_cliente <> "" Then
                        SQL = SQL & "   ,WEL_ORDEN_COMPRA       =   SUBSTR('" & num_pedido_cliente & "'1,49) " & vbCrLf
                End If
                If num_importe <> "" Then
                        SQL = SQL & "       ,WELIMPORTE             =   '" & num_importe & "' " & vbCrLf
                End If
                
                If num_cdad_tarimas > 0 Then
                        SQL = SQL & "       ,WEL_CDAD_BULTOS        =   '" & num_cdad_tarimas & "' " & vbCrLf
                        SQL = SQL & "       ,WEL_CDAD_TARIMAS       =   '" & num_cdad_tarimas & "' " & vbCrLf
                        SQL = SQL & "       ,WEL_CAJAS_TARIMAS      =   '" & num_cdad_cajas_tarima & "' " & vbCrLf
                ElseIf num_cdad_tarimas = 0 Then
                        SQL = SQL & "       ,WEL_CDAD_BULTOS        =   '" & num_cdad_bultos & "' " & vbCrLf
                End If
                
                If num_observacion <> "" Then
                        SQL = SQL & "       ,WELOBSERVACION         =   SUBSTR('" & num_observacion & "',1,1999) " & vbCrLf
                Else
                        SQL = SQL & "       ,WELOBSERVACION         =   '_PENDIENTE_' " & vbCrLf
                End If
                
                
                SQL = SQL & "       ,WELENTREGA_DOMICILIO   =   '" & "S" & "' " & vbCrLf
                SQL = SQL & "       ,WELPESO                =   '" & num_peso & "' " & vbCrLf
                SQL = SQL & "       ,WELVOLUMEN             =   '" & num_volumen & "' " & vbCrLf
                SQL = SQL & "       ,DATE_CREATED           =    " & "SYSDATE" & " " & vbCrLf
                SQL = SQL & "       ,CREATED_BY             =    " & "SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 29) " & vbCrLf
                SQL = SQL & "       ,WEL_COLLECT_PREPAID    =   SUBSTR('" & num_collect_prepaid & "',1,9) " & vbCrLf
                SQL = SQL & "       ,WELSTATUS              =   1 " & vbCrLf
                SQL = SQL & " WHERE  WELCLAVE   =   '" & clef & "' " & vbCrLf
                rs.Open SQL
                
                SQL = " UPDATE WEB_TRACKING_STAGE     " & vbCrLf
                SQL = SQL & " SET FECHA_DOCUMENTACION = SYSDATE     " & vbCrLf
                SQL = SQL & "       ,USR_DOC        =       " & "SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 29) " & vbCrLf
                SQL = SQL & " WHERE NUI = '" & clef & "'    " & vbCrLf
                
                rs.Open SQL
                
                If nuis_insertados = "" Then
                        nuis_insertados = clef
                Else
                        nuis_insertados = nuis_insertados & ", " & clef
                End If
                
                Call log_SQL("carga_archivo", "preparacion etiquetas ", cliente)
                
                
                'insercion de etiquetas
                SQL = ""
                If num_cdad_tarimas > 0 Then
                        SQL = "INSERT INTO ETRANS_ETIQUETAS_BULTOS ( " & vbCrLf
                        SQL = SQL & "    TEBCLAVE, TEB_WELCLAVE, TEBCONS_ETIQ,  " & vbCrLf
                        SQL = SQL & "    TEBTOT_ETIQ, CREATED_BY, DATE_CREATED)  " & vbCrLf
                        SQL = SQL & " SELECT SEQ_ETRANS_ETIQUETAS_BULTOS.nextval, " & clef & ", rownum  " & vbCrLf
                        SQL = SQL & " , to_number(" & num_cdad_tarimas & "), SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE " & vbCrLf
                        SQL = SQL & " FROM WEB_LTL " & vbCrLf
                        SQL = SQL & " WHERE rownum <= to_number(" & num_cdad_tarimas & ") "
                ElseIf num_cdad_bultos > 0 Then
                        SQL = "INSERT INTO ETRANS_ETIQUETAS_BULTOS ( " & vbCrLf
                        SQL = SQL & "    TEBCLAVE, TEB_WELCLAVE, TEBCONS_ETIQ,  " & vbCrLf
                        SQL = SQL & "    TEBTOT_ETIQ, CREATED_BY, DATE_CREATED)  " & vbCrLf
                        SQL = SQL & " SELECT SEQ_ETRANS_ETIQUETAS_BULTOS.nextval, " & clef & ", rownum  " & vbCrLf
                        SQL = SQL & " , to_number(" & num_cdad_bultos & "), SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE " & vbCrLf
                        SQL = SQL & " FROM WEB_LTL " & vbCrLf
                        SQL = SQL & " WHERE rownum <= to_number(" & num_cdad_bultos & ") "
                End If
                
                If SQL <> "" Then
                        rs.Open SQL
                        Call log_SQL("carga_archivo", "etiquetas listas ", cliente)
                End If
                        

                'concepto de recoleccion a domicilio
                SQL = "select NVL(logis.facturacion_TRAD.GET_IMPORTE_DEL_CONCEPTO('WELCLAVE=" & clef & ";CLIENTE=" & cliente & ";DIV=MXN;CHOCLAVE=1684;EMP=10'), 0) from dual"
                rs.Open SQL
                If Not rs.EOF Then
                        If rs.Fields(0) <> "0" Then
                                SQL = "INSERT INTO WEB_LTL_CONCEPTOS ( "
                                SQL = SQL & " WLCCLAVE, WLC_WELCLAVE, WLC_CHOCLAVE  "
                                SQL = SQL & " , WLC_IMPORTE, CREATED_BY, DATE_CREATED) "
                                SQL = SQL & " VALUES ( SEQ_WEB_LTL_CONCEPTOS.nextval, " & clef & ", 1684"
                                SQL = SQL & "     , " & rs.Fields(0) & " , SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE ) "
                                rs2.Open SQL
                        End If
                End If
                rs.Close
                
                
        Else
                'rs.ActiveConnection.BeginTrans
                'verificamos que no haya una linea vacia buscando el num de factura
                If NVL(oRS.Fields(id_factura)) <> "" Then
                        '20180614 -- >
                        posicion = IsInArray_Multi(NVL(oRS.Fields(id_factura)), talones_actualizar, 0)
                        '20180614 < --
                        
                        If posicion = -1 Then
                                Call log_SQL("carga_archivo", "inicio nuevo registro", cliente)
                                
                                Call get_remitente(cliente, mi_disclef, num_allclave_ori)
                        
                                Call log_SQL("carga_archivo", "remitente listo " & mi_disclef, cliente)
                        
                                If id_factura > -1 Then num_factura = Replace(NVL(oRS.Fields(id_factura)), "'", "''")
                                If id_pedido_cliente > -1 Then num_pedido_cliente = Replace(NVL(oRS.Fields(id_pedido_cliente)), "'", "''")
                                'Johnson & Johnson 15178 ------------------------------------------->
                                If cliente <> "15178" Then
                                        If id_orden_compra > -1 Then num_orden_compra = Replace(NVL(oRS.Fields(id_orden_compra)), "'", "''")
                                Else
                                        If id_orden_compra > -1 Then
                                                If Replace(NVL(oRS.Fields(id_orden_compra)), "'", "''") <> "" Then
                                                        num_orden_compra = Replace(NVL(oRS.Fields(id_orden_compra)), "'", "''")
                                                End If
                                        End If
                                End If
                                'Johnson & Johnson 15178 ------------------------------------------->
                                If id_cdad_bultos > -1 And cliente <> "3885" Then num_cdad_bultos = NVL(oRS.Fields(id_cdad_bultos))
                                If id_importe > -1 Then num_importe = Replace(NVL(oRS.Fields(id_importe)), "'", "''")
                                If id_peso > -1 Then num_peso = Replace(NVL(oRS.Fields(id_peso)), "'", "''")
                                If id_volumen > -1 Then num_volumen = Replace(NVL(oRS.Fields(id_volumen)), "'", "''")
                                If id_collect_prepaid > -1 Then
                                        If cliente = "3882" Then
                                                If NVL(oRS.Fields(id_collect_prepaid)) = "FLETE POR COBRAR" Then
                                                        num_collect_prepaid = "POR COBRAR"
                                                Else
                                                        num_collect_prepaid = "PREPAGADO"
                                                End If
                                        End If
                                End If
                                'Johnson & Johnson 15178 ------------------------------------------->
                                If id_fecha_cita_prog > -1 Then
                                        '20180606 -- > Carga GSK
                                        '   20181001 -- > Carga GSK se agregan 20501 y 20502
                                        If cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                                        '   20181001 < -- Carga GSK se agregan 20501 y 20502
                                                If Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") <> "" Then
                                                        If Replace(NVL(oRS.Fields(id_fecha_cita_prog + 1)), "'", "''") <> "" Then
                                                                num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & " " & Replace(Mid(NVL(oRS.Fields(id_fecha_cita_prog + 1)), 1, 5), "'", "''") & "', 'DD/MM/YYYY hh24:mi')"
                                                        Else
                                                                num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & "', 'DD/MM/YYYY')"
                                                        End If
                                                Else
                                                        num_fecha_cita_prog = "NULL"
                                                End If
                                        ElseIf Trim(UCase(tipo_carga)) = "CON_FACTURA" Then
                                            num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & "', 'DD/MM/YYYY')"
                                        Else
                                                num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & "', 'DD/MM/YYYY')"
                                        End If
                                        '20180606 < --
                                'Johnson & Johnson 15178 ------------------------------------------->
                                End If
                                
                                If cliente = "3885" Then
                                        'vamos a actualizar la cdad de bultos despues con el detalle
                                        num_cdad_bultos = 0
                                        num_fecha_factura = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") & "', 'MM/DD/YYYY')"
                                ElseIf cliente = "3624" Then
                                        num_fecha_factura = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") & "', 'DD.MM.YYYY')"
                                ElseIf cliente = "3081" Then
                                        num_fecha_factura = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") & "', 'DD/MM/YYYY')"
                                ElseIf cliente = "13128" Then
                                        num_fecha_factura = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") & "', 'DD.MM.YYYY')"
                                ElseIf cliente = "17873" Then
                                        num_fecha_factura = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") & "', 'DD/MM/YYYY')"
                                '20180606 -- > GSK
                                '   20181001 -- > Carga GSK se agregan 20501 y 20502
                                ElseIf cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                                '   20181001 < -- Carga GSK se agregan 20501 y 20502
                                        If Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") <> "" Then
                                                num_fecha_factura = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_factura)), "'", "''") & "', 'DD/MM/YYYY')"
                                        End If
                                        
                                        If id_tarimas > -1 Then
                                                'If NVL(oRS.Fields(id_tarimas)) <> "" And NVL(oRS.Fields(id_tarimas)) <> "0" And IsNumeric(NVL(oRS.Fields(id_tarimas))) Then
                                                If IsNumeric(NVL(oRS.Fields(id_tarimas))) Then
                                                        If CLng(NVL(oRS.Fields(id_tarimas))) >= 1 Then
                                                                num_cdad_empaques_x_bulto = num_cdad_bultos
                                                                '20181019 -- >
                                                                num_wcd_cajas_tarimas = num_cdad_bultos
                                                                '20181019 < --
                                                                num_cdad_bultos = NVL(oRS.Fields(id_tarimas))
                                                                '20181019 -- >
                                                                num_wcd_cdad_tarimas = CLng(NVL(oRS.Fields(id_tarimas)))
                                                                '20181019 < --
                                                                num_tpaclave = "1"
                                                                num_bulto_tpaclave = "9"
                                                        Else
                                                                num_cdad_empaques_x_bulto = "NULL"
                                                                num_tpaclave = "9"
                                                                num_bulto_tpaclave = "NULL"
                                                        End If
                                                Else
                                                        num_cdad_empaques_x_bulto = "NULL"
                                                        num_tpaclave = "9"
                                                        num_bulto_tpaclave = "NULL"
                                                End If
                                        End If
                                '20180606 < --
                                End If
                                
                                If tipo_carga = "CD" Then
                                        Call log_SQL("carga_archivo", "busqueda direccion entrega " & num_factura, cliente)
                                        
                                        Call get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest)
                                        
                                        Call log_SQL("carga_archivo", "direccion entrega lista " & num_factura, cliente)
                                        
                                        '< CHG-20220208: Se modifica la forma de obtener el siguiente Identity:
'                    SQL = "SELECT SEQ_WCROSS_DOCK.NEXTVAL FROM DUAL"
                                        SQL = "SELECT NVL(MAX(WCDCLAVE),0) + 1 FROM WCROSS_DOCK"
                                        ' CHG-20220208 >
                                        rs.Open SQL
                                        clef = rs.Fields(0)
                                        rs.Close
                
                                        
                                        'insercion
'<<C    C-CHG-20220525: Se modifica la l�gica para que ahora se obtenga un NUI reservado y ese ser� documentado.
                                        'SQL = " INSERT INTO WCROSS_DOCK ( " & vbCrLf
                                        'SQL = SQL & "    WCDCLAVE                           , WCDFACTURA  "
                                        'SQL = SQL & "    , WCD_PEDIDO_CLIENTE               , WCD_ORDEN_COMPRA  "
                                        'SQL = SQL & "    , WCDVOLUMEN                       , WCDPESO  "
                                        'SQL = SQL & "    , WCDIMPORTE                       , WCD_CDAD_BULTOS  "
                                        'SQL = SQL & "    , WCD_CCLCLAVE                     , WCD_DISCLEF  "
                                        'SQL = SQL & "    , WCD_CLICLEF                      , WCD_ALLCLAVE_ORI  "
                                        'SQL = SQL & "    , WCD_ALLCLAVE_DEST                , WCD_DIECLAVE, WCD_DIECLAVE_ENTREGA  "
                                        'SQL = SQL & "    , DATE_CREATED                     , CREATED_BY  "
                                        ''Johnson & Johnson 15178 ------------------------------------------->
                                        ''20180606 -- > Carga GSK
                                        ''If cliente = "15178" Then
                                        ''   20181001 -- > Carga GSK se agregan 20501 y 20502
                                        'If Cliente = "15178" Or Cliente = "20341" Or Cliente = "20305" Or Cliente = "20501" Or Cliente = "20502" Then
                                        ''   20181001 < -- Carga GSK se agregan 20501 y 20502
                                        ''20180606 < --
                                        '    '20181019 -- >
                                        '    'SQL = SQL & "   , WCD_FEC_CITA_PROGRAMADA  "
                                        '    SQL = SQL & "   , WCD_FEC_CITA_PROGRAMADA , WCD_CDAD_TARIMAS , WCD_CAJAS_TARIMAS , WCDCDAD_CAJAS "
                                        '    '20181019 < --
                                        'End If
                                        ''Johnson & Johnson 15178 ------------------------------------------->
                                        'SQL = SQL & "    , WCD_FECHA_FACTURA)  "
                                        'SQL = SQL & " VALUES ( " & clef & "                 , '" & num_factura & "' "
                                        'SQL = SQL & "     , '" & num_pedido_cliente & "'    , '" & num_orden_compra & "' "
                                        'SQL = SQL & "     , '" & num_volumen & "'           , '" & num_peso & "' "
                                        'SQL = SQL & "     , '" & num_importe & "'           , '" & num_cdad_bultos & "' "
                                        'SQL = SQL & "     , " & num_cclclave & "            , " & num_disclef & " "
                                        'SQL = SQL & "     , " & Cliente & "                 , " & num_allclave_ori & " "
                                        'SQL = SQL & "     , " & num_allclave_dest & "       , " & num_dieclave & ", " & IIf(num_dieclave_entrega > 0, num_dieclave_entrega, num_dieclave) & " "
                                        'SQL = SQL & "     , SYSDATE                         , SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30) "
                                        ''Johnson & Johnson 15178 ------------------------------------------->
                                        ''20180606 -- > Carga GSK
                                        ''If cliente = "15178" Then
                                        ''   20181001 -- > Carga GSK se agregan 20501 y 20502
                                        'If Cliente = "15178" Or Cliente = "20341" Or Cliente = "20305" Or Cliente = "20501" Or Cliente = "20502" Then
                                        ''   20181001 -- > Carga GSK se agregan 20501 y 20502
                                        ''20180606 < --
                                        '    '20181019 -- >
                                        '    'SQL = SQL & "   , " & num_fecha_cita_prog & " "
                                        '    SQL = SQL & "   , " & num_fecha_cita_prog & " , " & num_wcd_cdad_tarimas & " , " & num_wcd_cajas_tarimas & " , " & num_wcd_cdad_cajas
                                        '    '20181019 < --
                                        'End If
                                        ''Johnson & Johnson 15178 ------------------------------------------->
                                        'SQL = SQL & "     , " & num_fecha_factura & ") "
                                        
                                        'Primero busco el siguiente NUI disponible:
                                        SQL = "SELECT MIN(WCDCLAVE) FROM WCROSS_DOCK WHERE WCDSTATUS = 3 AND WCD_CLICLEF = '" & cliente & "'" & vbCrLf
                                        rs.Open SQL
                                                clef = rs.Fields(0)
                                        rs.Close
                                        
                                        Debug.Print clef
                                        
                                        'Ahora se actualiza la informaci�n recibida:
                                        SQL = " UPDATE   WCROSS_DOCK " & vbCrLf
                                        SQL = SQL & "   SET  WCDFACTURA             =   '" & num_factura & "' " & vbCrLf
                                        SQL = SQL & "       ,WCD_PEDIDO_CLIENTE     =   '" & num_pedido_cliente & "' " & vbCrLf
                                        SQL = SQL & "       ,WCD_ORDEN_COMPRA       =   '" & num_orden_compra & "' " & vbCrLf
                                        SQL = SQL & "       ,WCDVOLUMEN             =   '" & num_volumen & "' " & vbCrLf
                                        SQL = SQL & "       ,WCDPESO                =   '" & num_peso & "' " & vbCrLf
                                        If num_importe <> "" Then
                                                SQL = SQL & "       ,WCDIMPORTE             =   '" & num_importe & "' " & vbCrLf
                                        End If
                                        SQL = SQL & "       ,WCD_CDAD_BULTOS        =   '" & num_cdad_bultos & "' " & vbCrLf
                                        If num_cclclave <> -1 Then
                                                SQL = SQL & "       ,WCD_CCLCLAVE           =    " & num_cclclave & " " & vbCrLf
                                        End If
                                        SQL = SQL & "       ,WCD_DISCLEF            =    " & num_disclef & " " & vbCrLf
                                        SQL = SQL & "       ,WCD_ALLCLAVE_ORI       =    " & num_allclave_ori & " " & vbCrLf
                                        If num_allclave_dest <> -1 Then
                                            SQL = SQL & "       ,WCD_ALLCLAVE_DEST      =    " & num_allclave_dest & " " & vbCrLf
                                        End If
                                        If num_dieclave <> -1 Then
                                                SQL = SQL & "       ,WCD_DIECLAVE           =    " & num_dieclave & " " & vbCrLf
                                                SQL = SQL & "       ,WCD_DIECLAVE_ENTREGA   =    " & IIf(num_dieclave_entrega > 0, num_dieclave_entrega, num_dieclave) & " " & vbCrLf
                                        End If
                                        SQL = SQL & "       ,DATE_CREATED           =    " & "SYSDATE" & " " & vbCrLf
                                        SQL = SQL & "       ,CREATED_BY             =   SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30) " & vbCrLf
                                        
                                        If cliente = "15178" Or cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                                                SQL = SQL & "       ,WCD_FEC_CITA_PROGRAMADA    =   " & num_fecha_cita_prog & " " & vbCrLf
                                                SQL = SQL & "       ,WCD_CDAD_TARIMAS           =   " & num_wcd_cdad_tarimas & " " & vbCrLf
                                                SQL = SQL & "       ,WCD_CAJAS_TARIMAS          =   " & num_wcd_cajas_tarimas & " " & vbCrLf
                                                SQL = SQL & "       ,WCDCDAD_CAJAS              =   " & num_wcd_cdad_cajas & " " & vbCrLf
                                        End If
                                        
                                        SQL = SQL & "       ,WCD_FECHA_FACTURA      =   " & num_fecha_factura & " " & vbCrLf
                                        SQL = SQL & "       ,WCDSTATUS      =   1 " & vbCrLf
                                        SQL = SQL & " WHERE  WCDCLAVE   =   '" & clef & "' " & vbCrLf
' CC    -CHG-20220525>>
                                        'Debug.Print SQL
                                        
                                        rs.Open SQL
                                        
                                        Call log_SQL("carga_archivo", "insercion cross dock lista " & num_factura & "-" & clef, cliente)
                                
                                        Call CHECK_VALID_CD(clef)
                                
                                        If facturas_insertadas <> "" Then facturas_insertadas = facturas_insertadas & ", "
                                        facturas_insertadas = facturas_insertadas & num_factura
                                        
                                        If cliente = "3110" Then
                                                SQL = "SELECT 1 " & vbCrLf
                                                SQL = SQL & " FROM EDIRECCIONES_ENTREGA " & vbCrLf
                                                SQL = SQL & " , EALMACENES_LOGIS EAL " & vbCrLf
                                                SQL = SQL & " , EDESTINOS_POR_RUTA " & vbCrLf
                                                SQL = SQL & " WHERE DIECLAVE = " & IIf(num_dieclave_entrega > 0, num_dieclave_entrega, num_dieclave) & vbCrLf
                                                SQL = SQL & "   AND ALLCLAVE = " & num_allclave_ori & vbCrLf
                                                SQL = SQL & "   AND DER_VILCLEF(+) = DIEVILLE " & vbCrLf
                                                SQL = SQL & "   AND (NVL(DER_TIPO_ENTREGA, 'FORANEO 5') <> 'FORANEO 5' " & vbCrLf
                                                SQL = SQL & "   OR EXISTS " & vbCrLf
                                                SQL = SQL & "   ( " & vbCrLf
                                                SQL = SQL & "       SELECT /*+ ORDERED USE_NL(CCO) */ NULL " & vbCrLf
                                                SQL = SQL & "       FROM EBASES_POR_CONCEPT BPC, " & vbCrLf
                                                SQL = SQL & "       ECLIENT_APLICA_CONCEPTOS CCO, " & vbCrLf
                                                SQL = SQL & "       EPARAMETRO_RESTRICT PAR " & vbCrLf
                                                SQL = SQL & "       WHERE BPC_CHOCLAVE IN (1580, 1838) " & vbCrLf
                                                SQL = SQL & "       AND CCO_BPCCLAVE = BPCCLAVE " & vbCrLf
                                                SQL = SQL & "       AND CCO_CLICLEF = " & cliente & vbCrLf
                                                SQL = SQL & "       AND PARCLAVE = BPC_PARCLAVE " & vbCrLf
                                                SQL = SQL & "       AND PAR_VILCLEF_ORI = EAL.ALL_VILCLEF " & vbCrLf
                                                SQL = SQL & "       AND PAR_VILCLEF_DEST = DIEVILLE " & vbCrLf
                                                SQL = SQL & "       AND ROWNUM = 1 " & vbCrLf
                                                SQL = SQL & "   ) " & vbCrLf
                                                SQL = SQL & "   ) "
                                                rs.Open SQL
                                                If rs.EOF Then
                                                        facturas_insertadas = facturas_insertadas & " (OJO: esta factura es un FORANEO 5) "
                                                End If
                                                rs.Close
                                        End If
                
                                        
                                        If cliente = "3885" Then
                                                'consultamos el detalle de cajas
                                                oRS2.Open "Select * from [" & pestana_detalles & "] where F1 =""" & oRS.Fields(id_factura) & """", oConn, adOpenStatic, adLockOptimistic
                                                
                                                Do While Not oRS2.EOF
                                                        Call log_SQL("carga_archivo", "inicio busqueda referencia " & num_factura, cliente)
                                                        
                                                        status = get_referencia(cliente, oRS2.Fields(id_ref), num_eirclef)
                                                        
                                                        Call log_SQL("carga_archivo", "busqueda referencia listo " & num_factura, cliente)
                                                        
                                                        If status <> "ok" Then facturas_insertadas = facturas_insertadas & " " & status
                                                        
                                                        'INSERT REFERENCIA
                                                        SQL = " INSERT INTO WCREFERENCIA (WCRCLAVE" & vbCrLf
                                                        SQL = SQL & " ,WCR_CANTIDAD_REF " & vbCrLf
                                                        SQL = SQL & " ,WCR_EIRCLEF " & vbCrLf
                                                        SQL = SQL & " ,WCR_WCDCLAVE " & vbCrLf
                                                        SQL = SQL & " ,CREATED_BY " & vbCrLf
                                                        SQL = SQL & " ,DATE_CREATED)" & vbCrLf
                                                        SQL = SQL & " VALUES(SEQ_WCREFERENCIA.nextval" & vbCrLf
                                                        SQL = SQL & ",'" & oRS2.Fields(id_eirclef_cdad) & "'" & vbCrLf
                                                        SQL = SQL & ",'" & num_eirclef & "'" & vbCrLf
                                                        SQL = SQL & "," & clef & vbCrLf
                                                        SQL = SQL & ", SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE) "   'CREATED_BY,DATE_CREATED
                                                        rs.Open SQL
                                                        
                                                        Call log_SQL("carga_archivo", "insercion referencia lista " & num_factura, cliente)
                                                        
                                                        num_cdad_bultos = CDbl(num_cdad_bultos) + oRS2.Fields(4)
                                                        
                                                        oRS2.MoveNext
                                                Loop
                                                oRS2.Close
                                                
                                                Call log_SQL("carga_archivo", "inicio alta bultos " & num_factura, cliente)
                                                
                                                'Insercion de los bultos
                                                SQL = " INSERT INTO WCBULTOS ( "
                                                SQL = SQL & "   WCBCLAVE, WCB_WCDCLAVE, WCB_CANTIDAD "
                                                SQL = SQL & "   , WCBLARGO, WCBANCHO, WCBALTO "
                                                SQL = SQL & "   , WCB_TPACLAVE, DATE_CREATED, CREATED_BY)"
                                                SQL = SQL & " VALUES ( SEQ_WCBULTOS.nextval, " & clef & ", " & num_cdad_bultos
                                                SQL = SQL & "   , 0, 0, 0 "
                                                SQL = SQL & "   , 9, SYSDATE, SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30))"
                                                rs.Open SQL
                                                
                                                Call log_SQL("carga_archivo", "alta bultos lista " & num_factura, cliente)
                                                
                                                'actualizamos la cdad de bultos
                                                SQL = " UPDATE WCROSS_DOCK " & vbCrLf
                                                SQL = SQL & " SET WCD_CDAD_BULTOS = " & num_cdad_bultos & vbCrLf
                                                SQL = SQL & " WHERE WCDCLAVE = " & clef
                                                rs.Open SQL
                                                
                                                Call log_SQL("carga_archivo", "update bultos listo " & num_factura, cliente)
                                        End If
                                        
                                        If cliente = "3081" Then
                                                Call log_SQL("carga_archivo", "inicio alta bultos " & num_factura, cliente)
                                                
                                                'Insercion de los bultos
                                                SQL = " INSERT INTO WCBULTOS ( "
                                                SQL = SQL & "   WCBCLAVE, WCB_WCDCLAVE, WCB_CANTIDAD "
                                                SQL = SQL & "   , WCBLARGO, WCBANCHO, WCBALTO "
                                                SQL = SQL & "   , WCB_TPACLAVE, DATE_CREATED, CREATED_BY)"
                                                SQL = SQL & " VALUES ( SEQ_WCBULTOS.nextval, " & clef & ", " & num_cdad_bultos
                                                SQL = SQL & "   , 0, 0, 0 "
                                                SQL = SQL & "   , 9, SYSDATE, SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30))"
                                                rs.Open SQL
                                        End If
                                        
                                        '20180606 -- > Carga GSK
                                        '20181016 -- > No registra bultos para GSK
                                        'If cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                                        '    Call log_SQL("carga_archivo", "inicio alta bultos " & num_factura, cliente)
                                        '    'Insercion de los bultos
                                        '    SQL = " INSERT INTO WCBULTOS ( "
                                        '    SQL = SQL & "   WCBCLAVE, WCB_WCDCLAVE, WCB_CANTIDAD "
                                        '    SQL = SQL & "   , WCBLARGO, WCBANCHO, WCBALTO "
                                        '    SQL = SQL & "   , WCB_CDAD_EMPAQUES_X_BULTO "
                                        '    SQL = SQL & "   , WCB_BULTO_TPACLAVE "
                                        '    SQL = SQL & "   , WCB_TPACLAVE, DATE_CREATED, CREATED_BY)"
                                        '    SQL = SQL & " VALUES ( SEQ_WCBULTOS.nextval, " & clef & ", " & num_cdad_bultos
                                        '    SQL = SQL & "   , 0, 0, 0 "
                                        '    SQL = SQL & "   , " & num_cdad_empaques_x_bulto
                                        '    SQL = SQL & "   , " & num_bulto_tpaclave
                                        '    SQL = SQL & "   , " & num_tpaclave & ", SYSDATE, SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30))"
                                        '    rs.Open SQL
                                        'End If
                                        '20181016 < --
                                        '20180606 < --
                                        Call log_SQL("carga_archivo", "preparacion etiquetas " & num_factura, cliente)
                                        
                                        'insercion de etiquetas
                                        SQL = "INSERT INTO ETRANS_ETIQUETAS_BULTOS ( " & vbCrLf
                                        SQL = SQL & "    TEBCLAVE, TEB_WCDCLAVE, TEBCONS_ETIQ,  " & vbCrLf
                                        SQL = SQL & "    TEBTOT_ETIQ, CREATED_BY, DATE_CREATED)  " & vbCrLf
                                        SQL = SQL & " SELECT SEQ_ETRANS_ETIQUETAS_BULTOS.nextval, " & clef & ", rownum  " & vbCrLf
                                        SQL = SQL & " , to_number(" & num_cdad_bultos & "), SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE " & vbCrLf
                                        SQL = SQL & " FROM WCROSS_DOCK " & vbCrLf
                                        SQL = SQL & " WHERE rownum <= to_number(" & num_cdad_bultos & ") "
                                        
                                        rs.Open SQL
                                        
                                        Call log_SQL("carga_archivo", "etiquetas listas " & num_factura, cliente)
                                
                                                                ElseIf UCase(tipo_carga) = "CON_FACTURA" Then
                                                                        
                                                                        If NVL(oRS.Fields(id_factura)) <> "" Then
                                                                                Call log_SQL("carga_archivo", "inicio nuevo registro", cliente)
                                                                                Call get_remitente(cliente, mi_disclef, num_allclave_ori)
                                                                                Call log_SQL("carga_archivo", "remitente listo " & mi_disclef, cliente)
                                                                        End If
                                                                        
                                                                        If id_orden_compra > -1 Then num_orden_compra = Replace(NVL(oRS.Fields(id_orden_compra)), "'", "''")
                                                                        If id_pedido_cliente > -1 Then num_pedido_cliente = Replace(NVL(oRS.Fields(id_pedido_cliente)), "'", "''")
                                                                        If id_factura > -1 Then num_factura = Replace(NVL(oRS.Fields(id_factura)), "'", "''")
                                                                        If id_tarimas > -1 Then
                                                                                If IsNumeric(NVL(oRS.Fields(id_tarimas))) Then
                                                                                        If CLng(NVL(oRS.Fields(id_tarimas))) >= 1 Then
                                                                                                num_cdad_tarimas = CLng(NVL(oRS.Fields(id_tarimas)))
                                                                                                num_cdad_cajas_tarima = NVL(oRS.Fields(id_cdad_bultos))
                                                                                        Else
                                                                                                num_cdad_bultos = NVL(oRS.Fields(id_cdad_bultos))
                                                                                        End If
                                                                                End If
                                                                        Else
                                                                                num_cdad_bultos = NVL(oRS.Fields(id_cdad_bultos))
                                                                        End If
                                                                        num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & "', 'DD/MM/YYYY')"
                                                                        If cliente_con_seguro(cliente) = True Then
                                                                                num_importe = 1
                                                                        End If
                                                                        num_collect_prepaid = "PREPAGADO"
                                                                                                                
                                                                        my_cclclave = get_cclclave(oRS.Fields(id_dec_dir), cliente)
                                                                                                                                                
                                                                        '<<CHG-DESA-27022024-01: Se obtienen la CCLCLAVE y la DIECLAVE a partir del dato del Layout.
                                                                        Call get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest)
                                                                        
                                                                            If Obtiene_DIECLAVE_CCLCLAVE(oRS.Fields(id_dec_dir), cliente, num_dieclave, num_cclclave) = False Then
                                                                                status = "- direccion inexistente"
                                                                            Else
                                                                                If Valida_CCLCLAVE(num_cclclave) = True Then
                                                                                '    'ORP
                                                                                '    'num_dieclave = -1
                                                                                '    If num_dieclave <> "" Then
                                                                                '        num_dieclave = -1
                                                                                '    End If
                                                                                'Else
                                                                                '    num_cclclave = -1
                                                                                End If
                                                                            End If
                                                                                                                                                
                                                                        ''If my_cclclave <> "" Then
                                                                        ''    '''<<20240214: se reutiliza la funcionalidad de CrossDock
                                                                        ''    'Call get_direccion_entrega_ltl(my_cclclave, num_cclclave, num_allclave_dest)
                                                                        ''    status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
                                                                        ''    '''  20240214>>
                                                                        ''Else
                                                                        ''        status = "- direccion inexistente, o destino INSEGURO, INVALIDO o TIPO DE ENTREGA no autorizado"
                                                                        ''End If
                                                                        'CHG-DESA-27022024-01>>>
                                                                        
                                                                        Call log_SQL("carga_archivo", "preparacion insercion talon", cliente)
                                                                        
                                                                        'Primero busco el siguiente NUI disponible:
                                                                        SQL = "SELECT MIN(WELCLAVE) FROM WEB_LTL WHERE WELSTATUS = 3 AND WEL_CLICLEF = '" & cliente & "'" & vbCrLf
                                                                        rs.Open SQL
                                                                        clef = rs.Fields(0)
                                                                        rs.Close
                                                                        
                                                                        
                                                                        'If lstFactura = "" Then
                                                                        '       lstFactura = num_factura
                                                                        'Else
                                                                        '       If num_factura <> "" Then
                                                                        '               lstFactura = lstFactura & "," & num_factura
                                                                        '       End If
                                                                        'End If
                                                                        '
                                                                        '
                                                                        'If lstOrdenC = "" Then
                                                                        '       lstOrdenC = num_orden_compra
                                                                        'Else
                                                                        '       If num_orden_compra <> "" Then
                                                                        '               lstOrdenC = lstOrdenC & "," & num_orden_compra
                                                                        '       End If
                                                                        'End If
                                                                        
                                                                        Debug.Print "NUI > " & clef
                                                                        
                                                                        'Ahora se actualiza la informaci�n recibida:
                                                                        SQL = " UPDATE   WEB_LTL " & vbCrLf
                                                                        SQL = SQL & "       SET WELFACTURA             =   SUBSTR('" & num_factura & "',1,99) " & vbCrLf
                                                                        SQL = SQL & "           ,WEL_ALLCLAVE_ORI       =   '" & num_allclave_ori & "' " & vbCrLf
                                                                        SQL = SQL & "       ,WEL_DISCLEF            =   '" & num_disclef & "' " & vbCrLf
                                                                        SQL = SQL & "       ,WEL_ALLCLAVE_DEST      =   '" & num_allclave_dest & "' " & vbCrLf
                                                                        If num_cclclave <> -1 Then
                                                                        '<<<--CHG-DESA-27022024-01:
                                                                            'SQL = SQL & "   ,WEL_WCCLCLAVE          =   '" & num_cclclave & "' " & vbCrLf
                                                                            SQL = SQL & "   ,WEL_CCLCLAVE          =   '" & num_cclclave & "' " & vbCrLf
                                                                        'CHG-DESA-27022024-01-->>>
                                                                        End If
                                                                        
                                                                        '<<CHG-DESA-27022024-01
                                                                            If num_dieclave <> -1 Then
                                                                                SQL = SQL & "   ,WEL_DIECLAVE          =   '" & num_dieclave & "' " & vbCrLf
                                                                            End If
                                                                        'CHG-DESA-27022024-01>>
                                                                                                                                                
                                                                        
                                                                        If num_fecha_cita_prog = "" Then
                                                                                SQL = SQL & "       ,WEL_FECHA_RECOLECCION  =    " & "NULL" & " " & vbCrLf
                                                                                SQL = SQL & "       ,WELRECOL_DOMICILIO     =   '" & "N" & "' " & vbCrLf
                                                                        Else
                                                                                SQL = SQL & "       ,WEL_FECHA_RECOLECCION  =    " & num_fecha_cita_prog & " " & vbCrLf
                                                                                SQL = SQL & "       ,WELRECOL_DOMICILIO     =   '" & "S" & "' " & vbCrLf
                                                                        End If
                                                                        
                                                                        If num_orden_compra <> "" Then
                                                                                        SQL = SQL & "   ,WEL_ORDEN_COMPRA       =   SUBSTR('" & num_orden_compra & "',1,49) " & vbCrLf
                                                                        End If
                                                                        
                                                                        If num_importe <> "" Then
                                                                                        SQL = SQL & "       ,WELIMPORTE             =   '" & num_importe & "' " & vbCrLf
                                                                        End If
                                                                        
                                                                        If num_cdad_tarimas > 0 Then
                                                                                        SQL = SQL & "       ,WEL_CDAD_BULTOS        =   '" & num_cdad_tarimas & "' " & vbCrLf
                                                                                        SQL = SQL & "       ,WEL_CDAD_TARIMAS       =   '" & num_cdad_tarimas & "' " & vbCrLf
                                                                                        SQL = SQL & "       ,WEL_CAJAS_TARIMAS      =   '" & num_cdad_cajas_tarima & "' " & vbCrLf
                                                                        ElseIf num_cdad_tarimas = 0 Then
                                                                                        SQL = SQL & "       ,WEL_CDAD_BULTOS        =   '" & num_cdad_bultos & "' " & vbCrLf
                                                                        End If
                                                                        
                                                                        If num_observacion <> "" Then
                                                                                        SQL = SQL & "       ,WELOBSERVACION         =   SUBSTR('" & num_observacion & "',1,1999) " & vbCrLf
                                                                        Else
                                                                                        SQL = SQL & "       ,WELOBSERVACION         =   '_PENDIENTE_' " & vbCrLf
                                                                        End If
                                                                        
                                                                        SQL = SQL & "       ,WELENTREGA_DOMICILIO   =   '" & "S" & "' " & vbCrLf
                                                                        SQL = SQL & "       ,WELPESO                =   '" & num_peso & "' " & vbCrLf
                                                                        SQL = SQL & "       ,WELVOLUMEN             =   '" & num_volumen & "' " & vbCrLf
                                                                        SQL = SQL & "       ,DATE_CREATED           =    " & "SYSDATE" & " " & vbCrLf
                                                                        SQL = SQL & "       ,CREATED_BY             =    " & "SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 29) " & vbCrLf
                                                                        SQL = SQL & "       ,WEL_COLLECT_PREPAID    =   SUBSTR('" & num_collect_prepaid & "',1,9) " & vbCrLf
                                                                        SQL = SQL & "       ,WELSTATUS              =   1 " & vbCrLf
                                                                        SQL = SQL & " WHERE  WELCLAVE   =   '" & clef & "' " & vbCrLf
                                                                        
                                                                        rs.Open SQL
                                                                        
                                                                        
                                                                        
                                                                        
                                                                        SQL = " UPDATE WEB_TRACKING_STAGE     " & vbCrLf
                                                                        SQL = SQL & " SET FECHA_DOCUMENTACION = SYSDATE     " & vbCrLf
                                                                        SQL = SQL & "       ,USR_DOC        =       " & "SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 29) " & vbCrLf
                                                                        SQL = SQL & " WHERE NUI = '" & clef & "'    " & vbCrLf
                                                                        
                                                                        rs.Open SQL
                                                                        
                                                                        
                                                                        
                                                                        SQL = " INSERT INTO EFACTURAS_DOC ( " & vbCrLf
                                                                        SQL = SQL & "            ID_FACTURA_DOC " & vbCrLf
                                                                        SQL = SQL & "           ,NUI " & vbCrLf
                                                                        SQL = SQL & "           ,NO_FACTURA " & vbCrLf
                                                                        SQL = SQL & "           ,VALOR " & vbCrLf
                                                                        SQL = SQL & "           ,LINEAS_FACTURA " & vbCrLf
                                                                        SQL = SQL & "           ,NO_ORDEN " & vbCrLf
                                                                        SQL = SQL & "           ,PEDIDO " & vbCrLf
                                                                        SQL = SQL & "           ,DATE_CREATED " & vbCrLf
                                                                        SQL = SQL & "           ,CREATED_BY " & vbCrLf
                                                                        'COMPLEMENTO
                                                                        SQL = SQL & "           ,TIENE_COMPLEMENTO " & vbCrLf
                                                                                                                                                
                                                                        SQL = SQL & "   ) " & vbCrLf
                                                                        SQL = SQL & "   VALUES ( " & vbCrLf
                                                                        'ID_FACTURA_DOC
                                                                        SQL = SQL & "            (SELECT MAX(ID_FACTURA_DOC)+1 FROM EFACTURAS_DOC) " & vbCrLf
                                                                        'NUI
                                                                        SQL = SQL & "           ,'" & clef & "' " & vbCrLf
                                                                        'NO_FACTURA
                                                                        SQL = SQL & "           ,'" & num_factura & "' " & vbCrLf
                                                                        'VALOR
                                                                        If num_importe <> "" Then
                                                                                SQL = SQL & "       , " & num_importe & "" & vbCrLf
                                                                        Else
                                                                            SQL = SQL & "       , 0" & vbCrLf
                                                                        End If
                                                                        'LINEAS_FACTURA
                                                                        SQL = SQL & "       , " & lineas_factura & "" & vbCrLf
                                                                        'NO_ORDEN
                                                                        If num_orden_compra <> "" Then
                                                                                SQL = SQL & "       ,'" & num_orden_compra & "' " & vbCrLf
                                                                        Else
                                                                            SQL = SQL & "       ,'' " & vbCrLf
                                                                        End If
                                                                        'PEDIDO
                                                                        If num_pedido_cliente <> "" Then
                                                                                SQL = SQL & "       ," & num_pedido_cliente & " " & vbCrLf
                                                                        Else
                                                                            SQL = SQL & "       ,'' " & vbCrLf
                                                                        End If
                                                                        'DATE_CREATED
                                                                        SQL = SQL & "           ,SYSDATE " & vbCrLf
                                                                        'CREATED_BY
                                                                        SQL = SQL & "           ,SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30) " & vbCrLf
                                                                        'COMPLEMENTO
                                                                        SQL = SQL & "           ,SUBSTR('" & num_fac_complemento & "',1,2) " & vbCrLf
                                                                        SQL = SQL & "   ) " & vbCrLf
                                                                        
                                                                        rs.Open SQL
                                                                        
                                                                        
                                                                        'If nuis_insertados = "" Then
                                                                        '       nuis_insertados = clef
                                                                        'Else
                                                                        '       nuis_insertados = nuis_insertados & ", " & clef
                                                                        'End If
                                                                        
                                                                        'COMPLEMENTO
                                                                        If facturas_insertadas <> "" Then
                                                                                facturas_insertadas = facturas_insertadas & ", " & num_factura & IIf(num_fac_complemento = "S", " (complemento)", "")
                                                                        Else
                                                                                facturas_insertadas = facturas_insertadas & num_factura & IIf(num_fac_complemento = "S", " (complemento)", "")
                                                                        End If
                                     
                                                                        
                                                                        
                                                                        Call log_SQL("carga_archivo", "preparacion etiquetas ", cliente)
                                                                        
                                                                                        'insercion de etiquetas
                                                                        SQL = ""
                                                                        If num_cdad_tarimas > 0 Then
                                                                                        SQL = "INSERT INTO ETRANS_ETIQUETAS_BULTOS ( " & vbCrLf
                                                                                        SQL = SQL & "    TEBCLAVE, TEB_WELCLAVE, TEBCONS_ETIQ,  " & vbCrLf
                                                                                        SQL = SQL & "    TEBTOT_ETIQ, CREATED_BY, DATE_CREATED)  " & vbCrLf
                                                                                        SQL = SQL & " SELECT SEQ_ETRANS_ETIQUETAS_BULTOS.nextval, " & clef & ", rownum  " & vbCrLf
                                                                                        SQL = SQL & " , to_number(" & num_cdad_tarimas & "), SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE " & vbCrLf
                                                                                        SQL = SQL & " FROM WEB_LTL " & vbCrLf
                                                                                        SQL = SQL & " WHERE rownum <= to_number(" & num_cdad_tarimas & ") "
                                                                        ElseIf num_cdad_bultos > 0 Then
                                                                                        SQL = "INSERT INTO ETRANS_ETIQUETAS_BULTOS ( " & vbCrLf
                                                                                        SQL = SQL & "    TEBCLAVE, TEB_WELCLAVE, TEBCONS_ETIQ,  " & vbCrLf
                                                                                        SQL = SQL & "    TEBTOT_ETIQ, CREATED_BY, DATE_CREATED)  " & vbCrLf
                                                                                        SQL = SQL & " SELECT SEQ_ETRANS_ETIQUETAS_BULTOS.nextval, " & clef & ", rownum  " & vbCrLf
                                                                                        SQL = SQL & " , to_number(" & num_cdad_bultos & "), SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE " & vbCrLf
                                                                                        SQL = SQL & " FROM WEB_LTL " & vbCrLf
                                                                                        SQL = SQL & " WHERE rownum <= to_number(" & num_cdad_bultos & ") "
                                                                        End If
                                                                        
                                                                        If SQL <> "" Then
                                                                                        rs.Open SQL
                                                                                        Call log_SQL("carga_archivo", "etiquetas listas ", cliente)
                                                                        End If
                                                                                        
                                        
                                                                        'concepto de recoleccion a domicilio
                                                                        SQL = "select NVL(logis.facturacion_TRAD.GET_IMPORTE_DEL_CONCEPTO('WELCLAVE=" & clef & ";CLIENTE=" & cliente & ";DIV=MXN;CHOCLAVE=1684;EMP=10'), 0) from dual"
                                                                        rs.Open SQL
                                                                        If Not rs.EOF Then
                                                                                        If rs.Fields(0) <> "0" Then
                                                                                                        SQL = "INSERT INTO WEB_LTL_CONCEPTOS ( "
                                                                                                        SQL = SQL & " WLCCLAVE, WLC_WELCLAVE, WLC_CHOCLAVE  "
                                                                                                        SQL = SQL & " , WLC_IMPORTE, CREATED_BY, DATE_CREATED) "
                                                                                                        SQL = SQL & " VALUES ( SEQ_WEB_LTL_CONCEPTOS.nextval, " & clef & ", 1684"
                                                                                                        SQL = SQL & "     , " & rs.Fields(0) & " , SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE ) "
                                                                                                        rs2.Open SQL
                                                                                        End If
                                                                        End If
                                                                        rs.Close
                                                                        
                                                                        
                                                                Else
                                                                '<<CHG-DESA-27022024-01: Se obtienen la CCLCLAVE y la DIECLAVE a partir del dato del Layout.
                                                                    '''''<<20240214: se reutiliza la funcionalidad de CrossDock
                                                                        'Call get_direccion_entrega_ltl(oRS.Fields(id_cclclave), num_cclclave, num_allclave_dest)
                                                                    ''    status = get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest, num_direccion_cliente)
                                                                    '''''  20240214>>
                                                                    Call get_direccion_entrega_cd(cliente, oRS.Fields(id_dec_dir), num_cclclave, num_dieclave, num_dieclave_entrega, num_allclave_dest)
                                                                    If Obtiene_DIECLAVE_CCLCLAVE(oRS.Fields(id_dec_dir), cliente, num_dieclave, num_cclclave) = False Then
                                                                        status = " - direccion inexistente"
                                                                    Else
                                                                        If Valida_CCLCLAVE(num_cclclave) = True Then
                                                                        '    'ORP
                                                                        '    'num_dieclave = -1
                                                                        '    If num_dieclave = "" Then
                                                                        '        num_dieclave = -1
                                                                        '    End If
                                                                        'Else
                                                                        '    num_cclclave = -1
                                                                        End If
                                                                    End If
                                                                'CHG-DESA-27022024-01>>>>
                                        '< CHG-20220208: Se modifica la forma de obtener el siguiente Identity:
                    '                   SQL = "SELECT SEQ_WEB_LTL.NEXTVAL FROM DUAL"
                                        SQL = "SELECT NVL(MAX(WELCLAVE),0) + 1 FROM WEB_LTL"
                                        ' CHG-20220208>
                                        rs.Open SQL
                                        clef = rs.Fields(0)
                                        rs.Close
                                
                                        If cliente = "3882" Then
                                                'para Phillips tenemos que guardar las facturas en el campo de observacion
                                                num_observacion = "Factura(s): " & Replace(oRS.Fields(6), "'", "''") & vbCrLf & "CONTENIDO FERRETERIA"
                                                
                                                'tambien el pedido es una agregacion de varios campos:
                                                num_pedido_cliente = "R: " & Replace(oRS.Fields(1), "'", "''") & "    " & Replace(oRS.Fields(3), "'", "''")
                                                
                                                'agregamos el destinatario a la factura
                                                num_factura = num_factura & "-" & oRS.Fields(id_cclclave)
                                        End If
                
                        
                                        Call log_SQL("carga_archivo", "preparacion insercion talon", cliente)
                        
                                        'insercion
'                    SQL = "INSERT INTO WEB_LTL ( " & vbCrLf
'                    SQL = SQL & " WELCLAVE                                  , WEL_CLICLEF  "
'                    SQL = SQL & "   , WELCONS_GENERAL                       , WEL_ALLCLAVE_ORI "
'                    SQL = SQL & "   , WEL_DISCLEF                           , WEL_ALLCLAVE_DEST "
'                    SQL = SQL & "   , WEL_WCCLCLAVE                         , WEL_FECHA_RECOLECCION "
'                    SQL = SQL & "   , WELFACTURA                            , WEL_ORDEN_COMPRA "
'                    SQL = SQL & "   , WELIMPORTE                            , WEL_CDAD_BULTOS "
'                    SQL = SQL & "   , WELRECOL_DOMICILIO                    , WELENTREGA_DOMICILIO "
'                    SQL = SQL & "   , WELPESO                               , WELVOLUMEN "
'                    SQL = SQL & "   , DATE_CREATED                          , CREATED_BY "
'                    SQL = SQL & "   , WEL_COLLECT_PREPAID                   , WEL_FIRMA   "
'                    SQL = SQL & "   , WELOBSERVACION)   " & vbCrLf
'                    SQL = SQL & " SELECT /*+ INDEX(WEB_LTL IDX_WEL_CLICLEF) */ "
'                    SQL = SQL & clef & "                                    , " & Cliente
'                    SQL = SQL & "   , NVL(MAX(WELCONS_GENERAL)+1,1)         , " & num_allclave_ori
'                    SQL = SQL & "   , " & num_disclef & "                   , " & num_allclave_dest
'                    SQL = SQL & "   , " & num_cclclave & "                  , NULL "
'                    SQL = SQL & "   , '" & num_factura & "'                 , '" & num_pedido_cliente & "'"
'                    SQL = SQL & "   , '" & num_importe & "'                 , " & num_cdad_bultos
'                    SQL = SQL & "   , 'S'                                   , 'S' "
'                    SQL = SQL & "   , " & num_peso & "                      , " & num_volumen
'                    SQL = SQL & "   , SYSDATE                               , SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30) "
'                    SQL = SQL & "   , '" & num_collect_prepaid & "'         , SUBSTR(MD5(" & clef & " || " & Cliente & " || NVL(MAX(WELCONS_GENERAL)+1,1)), 1, 12) "
'                    SQL = SQL & "   , '" & num_observacion & "' "
'                    SQL = SQL & " FROM WEB_LTL " & vbCrLf
'                    SQL = SQL & " WHERE WEL_CLICLEF = " & Cliente & vbCrLf
                                        
                                        'Primero busco el siguiente NUI disponible:
                                        SQL = "SELECT MIN(WELCLAVE) FROM WEB_LTL WHERE WELSTATUS = 3 AND WEL_CLICLEF = '" & cliente & "'" & vbCrLf
                                        rs.Open SQL
                                                clef = rs.Fields(0)
                                        rs.Close
                                        Debug.Print "NUI > " & clef
                                
                                        'Ahora se actualiza la informaci�n recibida:
                                        SQL = " UPDATE   WEB_LTL " & vbCrLf
                                        SQL = SQL & "   SET  WEL_ALLCLAVE_ORI       =   '" & num_allclave_ori & "' " & vbCrLf
                                        SQL = SQL & "       ,WEL_DISCLEF            =   '" & num_disclef & "' " & vbCrLf
                                        SQL = SQL & "       ,WEL_ALLCLAVE_DEST      =   '" & num_allclave_dest & "' " & vbCrLf
                                        If num_cclclave <> -1 Then
                                                '<<<--CHG-DESA-27022024-01
                                                                                                'SQL = SQL & "       ,WEL_WCCLCLAVE          =   '" & num_cclclave & "' " & vbCrLf
                                                SQL = SQL & "       ,WEL_CCLCLAVE          =   '" & num_cclclave & "' " & vbCrLf
                                                                                                'CHG-DESA-27022024-01-->>>
                                        End If
                                        '<<CHG-DESA-27022024-01
                                            If num_dieclave <> -1 Then
                                                SQL = SQL & "       ,WEL_DIECLAVE          =   '" & num_dieclave & "' " & vbCrLf
                                            End If
                                        'CHG-DESA-27022024-01>>>
                                        SQL = SQL & "       ,WEL_FECHA_RECOLECCION  =    " & "NULL" & " " & vbCrLf
                                        SQL = SQL & "       ,WELFACTURA             =   SUBSTR('" & num_factura & "',1,99) " & vbCrLf
                                        SQL = SQL & "       ,WEL_ORDEN_COMPRA       =   SUBSTR('" & num_pedido_cliente & "'1,49) " & vbCrLf
                                        If num_importe <> "" Then
                                                SQL = SQL & "       ,WELIMPORTE             =   '" & num_importe & "' " & vbCrLf
                                        End If
                                        SQL = SQL & "       ,WEL_CDAD_BULTOS        =   '" & num_cdad_bultos & "' " & vbCrLf
                                        SQL = SQL & "       ,WELRECOL_DOMICILIO     =   '" & "N" & "' " & vbCrLf
                                        SQL = SQL & "       ,WELENTREGA_DOMICILIO   =   '" & "S" & "' " & vbCrLf
                                        SQL = SQL & "       ,WELPESO                =   '" & num_peso & "' " & vbCrLf
                                        SQL = SQL & "       ,WELVOLUMEN             =   '" & num_volumen & "' " & vbCrLf
                                        SQL = SQL & "       ,DATE_CREATED           =    " & "SYSDATE" & " " & vbCrLf
                                        SQL = SQL & "       ,CREATED_BY             =    " & "SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 29) " & vbCrLf
                                        SQL = SQL & "       ,WEL_COLLECT_PREPAID    =   SUBSTR('" & num_collect_prepaid & "',1,9) " & vbCrLf
                                        SQL = SQL & "       ,WELOBSERVACION         =   SUBSTR('" & num_observacion & "',1,1999) " & vbCrLf
                                        SQL = SQL & "       ,WELSTATUS              =   1 " & vbCrLf
                                        SQL = SQL & " WHERE  WELCLAVE   =   '" & clef & "' " & vbCrLf
' CC    -CHG-20220525>>
                                        rs.Open SQL
                                        
                                        Call log_SQL("carga_archivo", "preparacion insercion talon", cliente)
                                        
                                        If facturas_insertadas <> "" Then facturas_insertadas = facturas_insertadas & ", "
                                        facturas_insertadas = facturas_insertadas & num_factura
                                        
                                        
                                        Call log_SQL("carga_archivo", "preparacion etiquetas ", cliente)
                                        
                                        'insercion de etiquetas
                                        SQL = "INSERT INTO ETRANS_ETIQUETAS_BULTOS ( " & vbCrLf
                                        SQL = SQL & "    TEBCLAVE, TEB_WELCLAVE, TEBCONS_ETIQ,  " & vbCrLf
                                        SQL = SQL & "    TEBTOT_ETIQ, CREATED_BY, DATE_CREATED)  " & vbCrLf
                                        SQL = SQL & " SELECT SEQ_ETRANS_ETIQUETAS_BULTOS.nextval, " & clef & ", rownum  " & vbCrLf
                                        SQL = SQL & " , to_number(" & num_cdad_bultos & "), SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE " & vbCrLf
                                        SQL = SQL & " FROM WEB_LTL " & vbCrLf
                                        SQL = SQL & " WHERE rownum <= to_number(" & num_cdad_bultos & ") "
                                        
                                        rs.Open SQL
                        
                                        Call log_SQL("carga_archivo", "etiquetas listas ", cliente)
                        
                                        If cliente = "3882" Then
                                                'consultamos el detalle de cajas
                                                oRS2.Open "Select * from [" & pestana_detalles & "] where F1 =""" & oRS.Fields(id_factura) & """ and F2=""" & oRS.Fields(id_cclclave) & """", oConn, adOpenStatic, adLockOptimistic
                                                
                                                Do While Not oRS2.EOF
                                                        status = get_referencia(cliente, oRS2.Fields(id_ref), num_eirclef)
                                                        If status <> "ok" Then facturas_insertadas = facturas_insertadas & " " & status
                                                        
                                                        'INSERT REFERENCIA
                                                        SQL = " INSERT INTO WLREFERENCIA (WLRCLAVE" & vbCrLf
                                                        SQL = SQL & " ,WLR_CANTIDAD_REF " & vbCrLf
                                                        SQL = SQL & " ,WLR_EIRCLEF " & vbCrLf
                                                        SQL = SQL & " ,WLR_WELCLAVE " & vbCrLf
                                                        SQL = SQL & " ,CREATED_BY " & vbCrLf
                                                        SQL = SQL & " ,DATE_CREATED)" & vbCrLf
                                                        SQL = SQL & " VALUES(SEQ_WLREFERENCIA.nextval" & vbCrLf
                                                        SQL = SQL & ",'" & oRS2.Fields(id_eirclef_cdad) & "'" & vbCrLf
                                                        SQL = SQL & ",'" & num_eirclef & "'" & vbCrLf
                                                        SQL = SQL & "," & clef & vbCrLf
                                                        SQL = SQL & ", SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE) "   'CREATED_BY,DATE_CREATED
                                                        rs.Open SQL
                                                        
                '                    'crear el detalle de bultos
                '                    SQL = "INSERT INTO WPALETA_LTL ( " & vbCrLf
                '                    SQL = SQL & " WPLCLAVE, WPL_WELCLAVE, WPL_IDENTICAS  " & vbCrLf
                '                    SQL = SQL & " , WPL_TPACLAVE, WPLLARGO  " & vbCrLf
                '                    SQL = SQL & " , WPLANCHO, WPLALTO, CREATED_BY, DATE_CREATED) " & vbCrLf
                '                    SQL = SQL & " SELECT SEQ_WPALETA_LTL.nextval, " & clef & ", GET_CAJAS(" & num_eirclef & "," & oRS2.Fields(id_eirclef_cdad) & ")" & vbCrLf
                '                    SQL = SQL & "     , 9, REPLARGO " & vbCrLf
                '                    SQL = SQL & "     , REPANCHO, REPALTO " & vbCrLf
                '                    SQL = SQL & "     , SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE "
                '                    SQL = SQL & " FROM EREFERENCIA_PRESENTACION " & vbCrLf
                '                    SQL = SQL & " WHERE REPCLEF = GET_PRESENTACION(" & num_eirclef & ", 1) " & vbCrLf
                '                    rs.Open SQL
                '
                '                    SQL = "SELECT GET_PESO(" & num_eirclef & ", " & oRS2.Fields(id_eirclef_cdad) & "), GET_VOLUMEN(" & num_eirclef & ", " & oRS2.Fields(id_eirclef_cdad) & ") FROM DUAL "
                '                    rs.Open SQL
                '                    If Not rs.EOF Then
                '                        num_peso = num_peso + CDbl(rs.Fields(0))
                '                        num_volumen = num_volumen + CDbl(rs.Fields(1))
                '                    End If
                '                    rs.Close
                                                        
                                                        oRS2.MoveNext
                                                Loop
                                                oRS2.Close
                                                
                                                'concepto de recoleccion a domicilio
                                                SQL = "select NVL(logis.facturacion_TRAD.GET_IMPORTE_DEL_CONCEPTO('WELCLAVE=" & clef & ";CLIENTE=" & cliente & ";DIV=MXN;CHOCLAVE=1684;EMP=10'), 0) from dual"
                                                rs.Open SQL
                                                If Not rs.EOF Then
                                                        If rs.Fields(0) <> "0" Then
                                                                SQL = "INSERT INTO WEB_LTL_CONCEPTOS ( "
                                                                SQL = SQL & " WLCCLAVE, WLC_WELCLAVE, WLC_CHOCLAVE  "
                                                                SQL = SQL & " , WLC_IMPORTE, CREATED_BY, DATE_CREATED) "
                                                                SQL = SQL & " VALUES ( SEQ_WEB_LTL_CONCEPTOS.nextval, " & clef & ", 1684"
                                                                SQL = SQL & "     , " & rs.Fields(0) & " , SUBSTR('CargaWeb_" & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "', 1, 30), SYSDATE ) "
                                                                rs2.Open
                                                        End If
                                                End If
                                                rs.Close
                
                                                
                                                'actualizamos el peso/volumen de la LTL
                                                SQL = " UPDATE WEB_LTL " & vbCrLf
                                                SQL = SQL & " SET WELPESO = " & num_peso & vbCrLf
                                                SQL = SQL & "   , WELVOLUMEN = " & num_volumen & vbCrLf
                                                SQL = SQL & " WHERE WELCLAVE = " & clef
                                                rs.Open SQL
                                        End If
                                
                                End If
                                'rs.ActiveConnection.RollbackTrans
                        
                                Call log_SQL("carga_archivo", "fin nuevo registro", cliente)
                        Else
                                Call log_SQL("carga_archivo", "inicio actualiza pedido", cliente)
                                
                                SQL = "         SELECT WCD.WCDCLAVE " & vbCrLf
                                SQL = SQL & "        , NVL(TO_CHAR(DXP.DXP_FECHA_ENTREGA,'DD/MM/YYYY HH24:MI'), TO_CHAR(DXP2.DXP_FECHA_ENTREGA,'DD/MM/YYYY HH24:MI')) FECHA_ENTREGA " & vbCrLf
                                SQL = SQL & "     FROM WCROSS_DOCK WCD, EDET_EXPEDICIONES DXP, EDET_EXPEDICIONES DXP2 " & vbCrLf
                                SQL = SQL & "    WHERE DXP.DXP_TRACLAVE(+) = WCD.WCD_TRACLAVE " & vbCrLf
                                SQL = SQL & "      AND DXP.DXP_TDCDCLAVE(+) = WCD.WCD_TDCDCLAVE " & vbCrLf
                                SQL = SQL & "      AND DXP2.DXPCLAVE = " & vbCrLf
                                SQL = SQL & "          (SELECT NVL(MAX(DEX.DXPCLAVE), DXP.DXPCLAVE) " & vbCrLf
                                SQL = SQL & "             FROM EDET_EXPEDICIONES DEX " & vbCrLf
                                SQL = SQL & "            WHERE DEX.DXP_TIPO_ENTREGA = 'DIRECTO' " & vbCrLf
                                SQL = SQL & "           CONNECT BY PRIOR DEX.DXPCLAVE = DEX.DXP_DXPCLAVE " & vbCrLf
                                SQL = SQL & "            START WITH DEX.DXPCLAVE = " & vbCrLf
                                SQL = SQL & "                       (SELECT DXPCLAVE " & vbCrLf
                                SQL = SQL & "                          FROM WCROSS_DOCK               WCD2, " & vbCrLf
                                SQL = SQL & "                               ETRANS_DETALLE_CROSS_DOCK TDCD, " & vbCrLf
                                SQL = SQL & "                               ETRANSFERENCIA_TRADING TRA, " & vbCrLf
                                SQL = SQL & "                               EDET_EXPEDICIONES DXP " & vbCrLf
                                SQL = SQL & "                         Where WCD2.WCDCLAVE = WCD.WCDCLAVE " & vbCrLf
                                SQL = SQL & "                           AND TDCD.TDCDCLAVE = WCD.WCD_TDCDCLAVE " & vbCrLf
                                SQL = SQL & "                           AND TRA.TRACLAVE = TDCD.TDCD_TRACLAVE " & vbCrLf
                                SQL = SQL & "                           AND DXP.DXP_TDCDCLAVE = TDCD.TDCDCLAVE " & vbCrLf
                                SQL = SQL & "                           AND TRA.TRASTATUS = '1')) " & vbCrLf
                                'SQL = SQL & "     AND NVL(DXP.DXP_FECHA_ENTREGA, DXP2.DXP_FECHA_ENTREGA) IS NULL " & vbCrLf
                                SQL = SQL & "     AND WCD.WCDFACTURA = '" & oRS.Fields(id_factura) & "' " & vbCrLf
                                SQL = SQL & "     AND WCD.WCD_CLICLEF = " & cliente & vbCrLf
                                SQL = SQL & "     AND WCD.WCDSTATUS IN (1,2) " & vbCrLf
                                SQL = SQL & "   UNION " & vbCrLf
                                SQL = SQL & "  SELECT WCD.WCDCLAVE " & vbCrLf
                                SQL = SQL & "        , TO_CHAR(DXP.DXP_FECHA_ENTREGA,'DD/MM/YYYY HH24:MI') FECHA_ENTREGA " & vbCrLf
                                SQL = SQL & "     FROM WCROSS_DOCK WCD, EDET_EXPEDICIONES DXP " & vbCrLf
                                SQL = SQL & "    WHERE DXP.DXP_TRACLAVE(+) = WCD.WCD_TRACLAVE " & vbCrLf
                                SQL = SQL & "      AND DXP.DXP_TDCDCLAVE(+) = WCD.WCD_TDCDCLAVE " & vbCrLf
                                SQL = SQL & "     AND WCD.WCDFACTURA = '" & NVL(oRS.Fields(id_factura)) & "' " & vbCrLf
                                SQL = SQL & "     AND WCD.WCD_CLICLEF = " & cliente & vbCrLf
                                SQL = SQL & "     AND WCD.WCDSTATUS IN (1,2) " & vbCrLf
                                
                                rs.Open SQL
                                If Not rs.EOF Then
                                        If NVL(rs.Fields(1)) = "" Then
                                                If id_fecha_cita_prog > -1 Then
                                                        '20180616 -- > Carga GSK
                                                        '   20181001 -- > Carga GSK se agregan 20501 y 20502
                                                        If cliente = "20341" Or cliente = "20305" Or cliente = "20501" Or cliente = "20502" Then
                                                        '   20181001 -- > Carga GSK se agregan 20501 y 20502
                                                                If Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") <> "" Then
                                                                        If Replace(NVL(oRS.Fields(id_fecha_cita_prog + 1)), "'", "''") <> "" Then
                                                                                num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & " " & Replace(Mid(NVL(oRS.Fields(id_fecha_cita_prog + 1)), 1, 5), "'", "''") & "', 'DD/MM/YYYY hh24:mi')"
                                                                        Else
                                                                                num_fecha_cita_prog = "TO_DATE('" & Replace(NVL(oRS.Fields(id_fecha_cita_prog)), "'", "''") & "', 'DD/MM/YYYY')"
                                                                        End If
                                                                Else
                                                                        num_fecha_cita_prog = ""
                                                                End If
                                                        Else
                                                                num_fecha_cita_prog = ""
                                                        End If
                                                        '20180616 < --
                                                End If
                                        
                                                'Actualiza factura solo si no tiene fecha de entrega
                                                SQL = " UPDATE WCROSS_DOCK WCD" & vbCrLf
                                                SQL = SQL & " SET WCD.WCD_PEDIDO_CLIENTE = '" & oRS.Fields(id_pedido_cliente) & "'" & vbCrLf
                                                If num_fecha_cita_prog <> "" Then
                                                        SQL = SQL & " , WCD_FEC_CITA_PROGRAMADA = " & num_fecha_cita_prog & vbCrLf
                                                End If
                                                SQL = SQL & " WHERE WCD.WCDCLAVE = " & rs.Fields(0)
                                                rs2.Open SQL
                                                
                                                If facturas_actualizadas <> "" Then facturas_actualizadas = facturas_actualizadas & ", "
                                                facturas_actualizadas = facturas_actualizadas & NVL(oRS.Fields(id_factura))
                                        Else
                                                If facturas_error <> "" Then facturas_error = facturas_error & ", "
                                                facturas_error = facturas_error & NVL(oRS.Fields(id_factura))
                                        End If
                                Else
                                        If facturas_error2 <> "" Then facturas_error2 = facturas_error2 & ", "
                                        facturas_error2 = facturas_error2 & oRS.Fields(id_factura)
                                End If
                                rs.Close
                        End If
                        '20180614 < --
                End If
        End If
    oRS.MoveNext
Loop
     
Call log_SQL("carga_archivo", "registros insertados", cliente)
     

jmail.From = mail_From
jmail.FromName = mail_FromName
jmail.ClearRecipients

'para debug, estoy en los contactos ;)
jmail.AddRecipientBCC mail_grupo_error(0)


For i = 0 To UBound(Split(Replace(correo_electronico, ",", ";"), ";"))
    jmail.AddRecipient Trim(Split(Replace(correo_electronico, ",", ";"), ";")(i))
Next


'<---- CC-CHG-DESA-13032024-01: Se integra el grupo cargamasiva_smo@logis.com.mx a peticion del usuario
jmail.AddRecipient "cargamasiva_smo@logis.com.mx"
'CC-CHG-DESA-13032024-01 ---->




If UCase(tipo_carga) = "SIN_FACTURA" Then
        
        If nuis_insertados <> "" Then
                jmail.subject = "Exito carga de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
                jmail.body = "Hola, se cargo exitosamente el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", " & vbCrLf & vbCrLf
                jmail.body = jmail.body & " -Se insertaron los siguientes NUI�s: " & vbCrLf & nuis_insertados & vbCrLf
        Else
                jmail.subject = "Error al cargar el archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
                jmail.body = "Hola, se encontraron errores al cargar el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", " & vbCrLf & vbCrLf
                jmail.body = jmail.body & status & vbCrLf
        End If
        
        jmail.body = jmail.body & vbCrLf & " correspondiente al cliente: " & cliente & "." & vbCrLf
                
Else
        If (facturas_insertadas <> "" Or facturas_actualizadas <> "") And (facturas_error = "" And facturas_error2 = "") Then
                jmail.subject = "Exito carga de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
                jmail.body = "Hola, se cargo exitosamente el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", " & vbCrLf & vbCrLf
        ElseIf (facturas_insertadas <> "" Or facturas_actualizadas <> "") And (facturas_error <> "" Or facturas_error2 <> "") Then
                jmail.subject = "Carga parcial de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
                jmail.body = "Hola, se cargo parcialmente el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", " & vbCrLf & vbCrLf
        ElseIf (facturas_insertadas = "" And facturas_actualizadas = "") And (facturas_error <> "" Or facturas_error2 <> "") Then
                jmail.subject = "Error al cargar el archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
                jmail.body = "Hola, se encontraron errores al cargar el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", " & vbCrLf & vbCrLf
        Else
                jmail.subject = "Carga del archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
                jmail.body = "Hola, se cargo el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", " & vbCrLf & vbCrLf
        End If
        
        If facturas_insertadas <> "" Then
                jmail.body = jmail.body & " -Se insertaron " & UBound(Split(facturas_insertadas, ",")) + 1 & " facturas:" & vbCrLf & facturas_insertadas & vbCrLf
        End If

        If facturas_actualizadas <> "" Then
                jmail.body = jmail.body & " -Se actualizaron " & UBound(Split(facturas_actualizadas, ",")) + 1 & " facturas:" & vbCrLf & facturas_actualizadas & vbCrLf
        End If

        If facturas_error <> "" Then
                jmail.body = jmail.body & " -No se pudieron actualizar " & UBound(Split(facturas_error, ",")) + 1 & " facturas porque ya cuentan con fecha de entrega:" & vbCrLf & facturas_error & vbCrLf
        End If

        If facturas_error2 <> "" Then
                jmail.body = jmail.body & " -No se pudieron actualizar " & UBound(Split(facturas_error2, ",")) + 1 & " facturas:" & vbCrLf & facturas_error2 & vbCrLf
        End If
        
        jmail.body = jmail.body & vbCrLf & " correspondiente al cliente: " & cliente & "." & vbCrLf
End If

jmail.body = jmail.body & vbCrLf & "Saludos."
'jmail.subject = "Exito carga de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\")))
'jmail.body = "Hola, se cargo exitosamente el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", se insertaron " & UBound(Split(facturas_insertadas, ",")) + 1 & " facturas:" & vbCrLf & facturas_insertadas


If FSO.FileExists(Archivo) Then
        jmail.AddAttachment Archivo
End If

jmail.Send mail_server
    
End Sub


Private Function get_direccion_entrega_cd(mi_cliclef As String, mi_dec_num_dir_cliente As String, ByRef mi_die_cclclave As Long, _
    ByRef mi_dieclave As Long, ByRef mi_dieclave_entrega As Long, ByRef mi_allclave_dest As Integer, Optional mi_cliente_direccion As String, _
    Optional mi_cil_num_cliente_client As String) As String
    
   
    
    'recupera un DEC_NUM_DIR_CLIENTE y regresa los CCLCLAVE, DIECLAVE y cedis de entrega
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
    SQL = "SELECT DEC_DIECLAVE, DIE_CCLCLAVE, NVL((SELECT DER_ALLCLAVE FROM EDESTINOS_POR_RUTA WHERE DER_VILCLEF = DIEVILLE AND DER_ALLCLAVE IS NOT NULL AND ROWNUM = 1), 1) ALLCLAVE_DEST "
    SQL = SQL & "   , NVL(DEC_DIECLAVE_ENTREGA, -1) DEC_DIECLAVE_ENTREGA " & vbCrLf
    SQL = SQL & " FROM EDIRECCION_ENTR_CLIENTE_LIGA " & vbCrLf
    SQL = SQL & "   , EDIRECCIONES_ENTREGA " & vbCrLf
    If mi_cil_num_cliente_client <> "" Then
        SQL = SQL & "   , ECLIENT_CLIENTE_LIGA " & vbCrLf
    End If
    SQL = SQL & " WHERE DEC_CLICLEF = '" & Replace(mi_cliclef, "'", "''") & "'" & vbCrLf
    SQL = SQL & "   AND DEC_NUM_DIR_CLIENTE = '" & Replace(mi_dec_num_dir_cliente, "'", "''") & "'" & vbCrLf
    SQL = SQL & "   AND DEC_DIECLAVE = DIECLAVE " & vbCrLf
    SQL = SQL & "   AND DIE_STATUS = 1 " & vbCrLf
    '20180606 -- > Restriccion
    SQL = SQL & "   AND EXISTS (" & vbCrLf
    SQL = SQL & "       SELECT NULL FROM EDESTINOS_POR_RUTA " & vbCrLf
    SQL = SQL & "        WHERE DER_VILCLEF = DIEVILLE " & vbCrLf
    SQL = SQL & "          AND NVL(DER_ALLCLAVE, 1) > 0 " & vbCrLf
    SQL = SQL & "          AND DER_TIPO_ENTREGA NOT IN ('INSEGURO', 'INVALIDO') " & vbCrLf
    SQL = SQL & "          AND SF_LOGIS_CLIENTE_RESTRIC(DEC_CLICLEF, DER_TIPO_ENTREGA) = 1 " & vbCrLf
    SQL = SQL & "       ) " & vbCrLf
    '20180606 < --
    If mi_cil_num_cliente_client <> "" Then
        SQL = SQL & "  AND CIL_CLICLEF = '" & Replace(mi_cliclef, "'", "''") & "'" & vbCrLf
        SQL = SQL & "  AND CIL_CCLCLAVE = DIE_CCLCLAVE " & vbCrLf
        SQL = SQL & "  AND CIL_NUM_CLIENTE_CLIENT = '" & Replace(mi_cil_num_cliente_client, "'", "''") & "'" & vbCrLf
    End If
    
    rs.Open SQL
    If rs.EOF Then
        '20180606 -- > Restriccion
        'get_direccion_entrega_cd = "- direccion inexistente:"
        '20180606 < --
        get_direccion_entrega_cd = "- direccion inexistente, o destino INSEGURO, INVALIDO o TIPO DE ENTREGA no autorizado:" & _
            " id: " & mi_dec_num_dir_cliente & _
            vbCrLf & "direccion: " & mi_cliente_direccion
    ElseIf rs.RecordCount > 1 Then
        get_direccion_entrega_cd = "- existe mas de un registro ligado a este numero direccion:" & _
            " id: " & mi_dec_num_dir_cliente & _
            vbCrLf & "direccion: " & mi_cliente_direccion
    Else
        get_direccion_entrega_cd = "ok"
        mi_die_cclclave = rs.Fields("DIE_CCLCLAVE")
        mi_dieclave = rs.Fields("DEC_DIECLAVE")
        mi_dieclave_entrega = rs.Fields("DEC_DIECLAVE_ENTREGA")
        mi_allclave_dest = rs.Fields("ALLCLAVE_DEST")
    End If
    rs.Close
    
    Set rs = Nothing

End Function

Private Function get_direccion_entrega_ltl(mi_wcclclave As String, ByRef mi_cclclave As Long, ByRef mi_allclave_dest As Integer, Optional mi_cliente_direccion As String) As String
    
    
    
    'recupera un WCCLCLAVE y regresa el cedis de entrega
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
        
    SQL = "SELECT WCCLCLAVE, NVL((SELECT DER_ALLCLAVE FROM EDESTINOS_POR_RUTA WHERE DER_VILCLEF = WCCL_VILLE AND DER_ALLCLAVE IS NOT NULL AND ROWNUM = 1), 1) ALLCLAVE_DEST "
    SQL = SQL & " FROM WEB_CLIENT_CLIENTE "
    SQL = SQL & " WHERE WCCLCLAVE = '" & Replace(mi_wcclclave, "'", "''") & "'"
    '20180606 -- > Restriccion
    SQL = SQL & "   AND EXISTS (" & vbCrLf
    SQL = SQL & "       SELECT NULL FROM EDESTINOS_POR_RUTA " & vbCrLf
    SQL = SQL & "        WHERE DER_VILCLEF = WCCL_VILLE " & vbCrLf
    SQL = SQL & "          AND NVL(DER_ALLCLAVE, 1) > 0 " & vbCrLf
    SQL = SQL & "          AND DER_TIPO_ENTREGA NOT IN ('INSEGURO', 'INVALIDO') " & vbCrLf
    SQL = SQL & "          AND SF_LOGIS_CLIENTE_RESTRIC(WCCL_CLICLEF, DER_TIPO_ENTREGA) = 1 " & vbCrLf
    SQL = SQL & "       ) " & vbCrLf
        
                   
        
    '20180606 < --
    rs.Open SQL
    If rs.EOF Then
        '20180606 -- > Restriccion
        'get_direccion_entrega_ltl = "- direccion inexistente:"
        '20180606 < --
        get_direccion_entrega_ltl = "- direccion inexistente, o destino INSEGURO, INVALIDO o TIPO DE ENTREGA no autorizado:" & _
            " id: " & mi_wcclclave & _
            vbCrLf & "direccion: " & mi_cliente_direccion
        
    Else
        get_direccion_entrega_ltl = "ok"
        mi_allclave_dest = rs.Fields("ALLCLAVE_DEST")
        mi_cclclave = rs.Fields("WCCLCLAVE")
    End If
    rs.Close
    
    Set rs = Nothing

End Function

Private Function get_remitente(mi_cliclef As String, mi_disclef As String, ByRef mi_allclave_ori As Integer) As String
    
    'recupera un cliente y disclef y regresa el cedis de origen
    
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
    SQL = "SELECT DISCLEF, NVL((SELECT DER_ALLCLAVE FROM EDESTINOS_POR_RUTA WHERE DER_VILCLEF = DISVILLE AND DER_ALLCLAVE IS NOT NULL AND ROWNUM = 1), 1) ALLCLAVE_ORI "
    SQL = SQL & " FROM EDISTRIBUTEUR "
    SQL = SQL & " WHERE DISCLEF = '" & Replace(mi_disclef, "'", "''") & "'"
    SQL = SQL & "   AND DISCLIENT = '" & Replace(mi_cliclef, "'", "''") & "'"
    rs.Open SQL
    If rs.EOF Then
        get_remitente = "- remitente inexistente:" & _
            " id: " & mi_disclef & vbCrLf & vbCrLf
    Else
        get_remitente = "ok"
        mi_disclef = rs.Fields("DISCLEF")
        mi_allclave_ori = rs.Fields("ALLCLAVE_ORI")
    End If
    rs.Close
    
    Set rs = Nothing

End Function

Private Function get_factura_duplicada(mi_factura As String, mi_cliente As String) As String
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
    Dim rs2 As New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenForwardOnly
    rs2.LockType = adLockBatchOptimistic
    rs2.ActiveConnection = Db_link_orfeo
    
    SQL = "select 'web' TIPO, wcdclave, wcdfactura FACTURA, DECODE(wcdstatus, 1, 'Activo', 'Cancelado') STATUS " & vbCrLf
    SQL = SQL & "  , wco_manif_num MANIF, tracons_general NUM_ENTRADA" & vbCrLf
    SQL = SQL & "  from WCROSS_DOCK " & vbCrLf
    SQL = SQL & "  , WCDET_CONVERTIDOR " & vbCrLf
    SQL = SQL & "  ,WCONVERTIDOR " & vbCrLf
    SQL = SQL & "  , ETRANS_DETALLE_CROSS_DOCK " & vbCrLf
    SQL = SQL & "  , ETRANSFERENCIA_TRADING " & vbCrLf
    SQL = SQL & "  where wcdfactura = '" & Replace(mi_factura, "'", "''") & "' " & vbCrLf
    SQL = SQL & "  and wcd_cliclef = " & mi_cliente & vbCrLf
    SQL = SQL & "  and wcdc_wcdclave(+) = wcdclave " & vbCrLf
    SQL = SQL & "  and wcoclave(+) = wcdc_wcoclave " & vbCrLf
    SQL = SQL & "  and Tdcdclave(+) = wcd_Tdcdclave " & vbCrLf
    SQL = SQL & "  and Traclave(+) = wcd_Traclave " & vbCrLf
    SQL = SQL & "  union all " & vbCrLf
    SQL = SQL & "  select 'orfeo', tdcdclave, tdcdfactura, DECODE(trastatus, '1', DECODE(tdcdstatus, '1', 'Activo', 'Cancelado'), 'Cancelado'), null, tracons_general " & vbCrLf
    SQL = SQL & "  from ETRANS_DETALLE_CROSS_DOCK " & vbCrLf
    SQL = SQL & "    , ETRANSFERENCIA_TRADING " & vbCrLf
    SQL = SQL & "  where tdcdfactura = '" & Replace(mi_factura, "'", "''") & "' " & vbCrLf
    SQL = SQL & "  and traclave = tdcd_traclave " & vbCrLf
    SQL = SQL & "  and tdcdstatus = '1' " & vbCrLf
    SQL = SQL & "  and trastatus = '1' " & vbCrLf
    SQL = SQL & "  and tra_cliclef = " & mi_cliente & vbCrLf
    SQL = SQL & "  and tdcd_dxpclave_ori is null " & vbCrLf
    SQL = SQL & "  and not exists ( " & vbCrLf
    SQL = SQL & "        select null " & vbCrLf
    SQL = SQL & "    from WCROSS_DOCK " & vbCrLf
    SQL = SQL & "    where wcd_tdcdclave = tdcdclave)"
    rs.Open SQL
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            get_factura_duplicada = get_factura_duplicada & "- la factura " & rs.Fields("FACTURA") & " ya existente en " & rs.Fields("TIPO") & " con el status " & rs.Fields("STATUS") & "."
            If NVL(rs.Fields("MANIF")) <> "" Then get_factura_duplicada = get_factura_duplicada & " Manifiesto " & rs.Fields("MANIF")
            If NVL(rs.Fields("NUM_ENTRADA")) <> "" Then get_factura_duplicada = get_factura_duplicada & " Entrada " & rs.Fields("NUM_ENTRADA")
            
            If rs.Fields("TIPO") = "web" And rs.Fields("STATUS") = "Cancelado" And NVL(rs.Fields("NUM_ENTRADA")) = "" Then
                'reactivamos esta factura
                SQL = "UPDATE WCROSS_DOCK SET WCDSTATUS = 1 WHERE WCDCLAVE = " & rs.Fields("WCDCLAVE")
                rs2.Open SQL
                get_factura_duplicada = get_factura_duplicada & " Se reactivo esta factura en web."
            End If
            
            rs.MoveNext
        Loop
    Else
        get_factura_duplicada = "ok"
    End If
    
    rs.Close
    
    Set rs = Nothing
    Set rs2 = Nothing


End Function

Private Function get_referencia(mi_cliclef As String, mi_ref As String, ByRef mi_eirclef) As String
    
    'recupera un cliente y eirclef para verificar la existencia de la referencia con sus dimensiones
    
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
    SQL = "SELECT EIRCLEF, NVL(REPALTO * REPANCHO * REPLARGO, 0) VOL, NVL(REPPESO, 0) PESO  " & vbCrLf
    SQL = SQL & " FROM EINVENTARIO_REFERENCIA " & vbCrLf
    SQL = SQL & "   , EREFERENCIA_PRESENTACION " & vbCrLf
    SQL = SQL & " WHERE EIR_CLICLEF = '" & Replace(mi_cliclef, "'", "''") & "'" & vbCrLf
    SQL = SQL & "   AND EIRREFERENCIA = SUBSTR('" & Replace(mi_ref, "'", "''") & "', 1, 20)" & vbCrLf
    SQL = SQL & "   AND REPCLEF(+) = GET_PRESENTACION(EIRCLEF, 1)"
    rs.Open SQL
    If rs.EOF Then
        get_referencia = "- referencia inexistente:" & _
            " id: " & mi_ref
    Else
        If (rs.Fields("VOL") = "0" Or rs.Fields("PESO") = "0") And mi_cliclef <> "3885" Then
            get_referencia = "- la referencia " & mi_ref & " tiene datos logisticos incorrectos."
        Else
            get_referencia = "ok"
            mi_eirclef = rs.Fields("EIRCLEF")
        End If
    End If
    rs.Close
    
    Set rs = Nothing

End Function

Private Function IsInArray_Multi(ByVal clef As String, array_tab() As String, Col_num As Integer, Optional start_with As Long) As Long
    'clef : value to find in the array array_tab at the col : col_num
    'give number of line where is the data
    'start_with permite definir una linea para no empezar en 0
    Dim i As Long
'    If array_tab(0, 0) = "" Then
'        IsInArray_Multi = -1
'        Exit Function
'    End If
    For i = start_with To UBoundCheck(array_tab, 2)
        If CStr(array_tab(Col_num, i)) = CStr(clef) Then
            IsInArray_Multi = i
            Exit Function
        End If
    Next
    IsInArray_Multi = -1
End Function


Private Function get_cclclave(destinatario, cliente) As String
    SQL = "        SELECT DIE.DIE_CCLCLAVE  " & vbCrLf
SQL = SQL & " FROM EDIRECCION_ENTR_CLIENTE_LIGA   DIL   " & vbCrLf
SQL = SQL & "INNER JOIN  EDIRECCIONES_ENTREGA         DIE ON DIE.DIECLAVE = DIL.DEC_DIECLAVE    " & vbCrLf
SQL = SQL & "INNER JOIN  ECIUDADES                    CIU    ON CIU.VILCLEF = DIEVILLE  " & vbCrLf
SQL = SQL & "INNER JOIN  EESTADOS                     EST    ON EST.ESTESTADO = CIU.VIL_ESTESTADO       " & vbCrLf
SQL = SQL & "INNER JOIN  ECLIENT_CLIENTE              CCL    ON CCL.CCLCLAVE = DIE.DIE_CCLCLAVE " & vbCrLf
SQL = SQL & "INNER JOIN  EDESTINOS_POR_RUTA           DER    ON DER_VILCLEF = VILCLEF   " & vbCrLf
SQL = SQL & "WHERE  1=1 " & vbCrLf
SQL = SQL & "    AND    CCL.CCL_STATUS      =   1   " & vbCrLf
SQL = SQL & "    AND    DIE.DIE_STATUS      =   1   " & vbCrLf
SQL = SQL & "    AND    EST.EST_PAYCLEF     =   'N3' " & vbCrLf
SQL = SQL & "   AND DER.DER_ALLCLAVE    >   0   " & vbCrLf
SQL = SQL & "   AND SF_LOGIS_CLIENTE_RESTRIC(DIL.DEC_CLICLEF, DER.DER_TIPO_ENTREGA) =   1   " & vbCrLf
SQL = SQL & "    AND DEC_NUM_DIR_CLIENTE = '" & destinatario & "' " & vbCrLf
SQL = SQL & "    AND DEC_CLICLEF = " & cliente & "    "

rs.Open SQL
    If Not rs.EOF Then
        get_cclclave = rs.Fields("DIE_CCLCLAVE")
        
'        Do While Not rs.EOF
'            get_cclclave = rs.Fields("DIE_CCLCLAVE")
'
'            SQL = " SELECT WCCLCLAVE, NVL((SELECT DER_ALLCLAVE FROM EDESTINOS_POR_RUTA WHERE DER_VILCLEF = WCCL_VILLE AND DER_ALLCLAVE IS NOT NULL AND ROWNUM = 1), 1) ALLCLAVE_DEST  FROM WEB_CLIENT_CLIENTE  WHERE WCCLCLAVE = '" & get_cclclave & "'   /*AND EXISTS (" & vbCrLf
'            SQL = SQL & " SELECT NULL FROM EDESTINOS_POR_RUTA " & vbCrLf
'            SQL = SQL & " WHERE DER_VILCLEF = WCCL_VILLE " & vbCrLf
'            SQL = SQL & " AND NVL(DER_ALLCLAVE, 1) > 0  " & vbCrLf
'            SQL = SQL & " AND DER_TIPO_ENTREGA NOT IN ('INSEGURO', 'INVALIDO') " & vbCrLf
'            SQL = SQL & " AND SF_LOGIS_CLIENTE_RESTRIC(WCCL_CLICLEF, DER_TIPO_ENTREGA) = 1 " & vbCrLf
'            SQL = SQL & " )*/ "
'
'            rs2.Open SQL
'            If Not rs2.EOF Then
'                rs2.Close
'                Exit Do
'            End If
'            rs2.Close
'            rs.MoveNext
'        Loop
    End If
rs.Close
    
'Set rs = Nothing

End Function
Function es_valida_factura_cliente(CliClef, wel_factura)
        Dim res, sqlValFact
        
        res = True
        sqlValFact = ""
        
        sqlValFact = sqlValFact & " SELECT      WELFACTURA " & vbCrLf
        sqlValFact = sqlValFact & " FROM        WEB_LTL " & vbCrLf
        sqlValFact = sqlValFact & " WHERE       WEL_CLICLEF     <>      '" & CliClef & "' " & vbCrLf
        sqlValFact = sqlValFact & "     AND     WELFACTURA      =       '" & wel_factura & "' " & vbCrLf
        sqlValFact = sqlValFact & "     AND     WELFACTURA      <>      '_PENDIENTE_' " & vbCrLf
        sqlValFact = sqlValFact & "     AND     WELSTATUS       NOT     IN      (0,3) " & vbCrLf
        
        rs.Open sqlValFact
        If rs.EOF Then
                res = False
        End If
        rs.Close
        
        If res = True Then
                sqlValFact = " SELECT   FD.NO_FACTURA "
                sqlValFact = sqlValFact & " FROM        EFACTURAS_DOC FD "
                sqlValFact = sqlValFact & "     INNER JOIN WEB_LTL WEL "
                sqlValFact = sqlValFact & "             ON FD.NUI = WEL.WELCLAVE "
                sqlValFact = sqlValFact & " WHERE       FD.NO_FACTURA = '" & wel_factura & "' "
                sqlValFact = sqlValFact & "     AND     WEL.WEL_CLICLEF <> '" & CliClef & "' "
                
                rs.Open sqlValFact
                If rs.EOF Then
                        res = False
                End If
                rs.Close
        End If
        
        es_valida_factura_cliente = res
End Function
Function existe_factura_cliente(CliClef, wel_factura)
        Dim res, sqlValFact
        
        res = False
        sqlValFact = ""
        
        sqlValFact = sqlValFact & " SELECT      WELFACTURA " & vbCrLf
        sqlValFact = sqlValFact & " FROM        WEB_LTL " & vbCrLf
        sqlValFact = sqlValFact & " WHERE       WEL_CLICLEF     =      '" & CliClef & "' " & vbCrLf
        sqlValFact = sqlValFact & "     AND     WELFACTURA      =       '" & wel_factura & "' " & vbCrLf
        sqlValFact = sqlValFact & "     AND     WELFACTURA      <>      '_PENDIENTE_' " & vbCrLf
        sqlValFact = sqlValFact & "     AND     WELSTATUS       NOT     IN      (0,3) " & vbCrLf
        
        rs.Open sqlValFact
        If Not rs.EOF Then
            res = True
        End If
        rs.Close
        
        If res = False Then
                sqlValFact = " SELECT   FD.NO_FACTURA "
                sqlValFact = sqlValFact & " FROM        EFACTURAS_DOC FD "
                sqlValFact = sqlValFact & "     INNER JOIN WEB_LTL WEL "
                sqlValFact = sqlValFact & "             ON FD.NUI = WEL.WELCLAVE "
                sqlValFact = sqlValFact & " WHERE       FD.NO_FACTURA = '" & wel_factura & "' "
                sqlValFact = sqlValFact & "     AND     WEL.WEL_CLICLEF = '" & CliClef & "' "
                
                rs.Open sqlValFact
                If Not rs.EOF Then
                    res = True
                End If
                rs.Close
        End If
        Debug.Print wel_factura & " - " & res
        existe_factura_cliente = res
End Function

'COMPLEMENTO
Function obtiene_valor_complemento_pedido(valor As String)
        Dim res As String
        
        res = ""
        
        If UCase(valor) = "SI" Or UCase(valor) = "S" Then
                res = "S"
        End If
        
        If UCase(valor) = "NO" Or UCase(valor) = "N" Then
                res = "N"
        End If
        
        obtiene_valor_complemento_pedido = res
        
End Function
Function ValidarLayOut(tipo_carga)
    Dim msj As String
    
    msj = ""
    
    oRS3.Open "Select * from [" & pestana_encabezados & "] ", oConn, adOpenStatic, adLockOptimistic
    
    
    
    If oRS3.EOF Then
        msj = "El archivo no cuenta con registros."
    Else
        If tipo_carga = "SIN_FACTURA" Then
            If Trim(UCase(oRS3.Fields(0).Name)) <> "REFERENCIA" Then
                msj = msj & " El archivo no tiene la columna 'REFERENCIA' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(1).Name)) <> "N.CLIENTE" And Trim(UCase(oRS3.Fields(1).Name)) <> "N#CLIENTE" Then
                msj = msj & " El archivo no tiene la columna 'N.CLIENTE' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(2).Name)) <> "CAJAS" Then
                msj = msj & " El archivo no tiene la columna 'CAJAS' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(3).Name)) <> "TARIMAS" Then
                msj = msj & " El archivo no tiene la columna 'TARIMAS' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(4).Name)) <> "FECHA" Then
                msj = msj & " El archivo no tiene la columna 'FECHA' o se encuentra en otra posici�n. " & vbCrLf
            End If
        ElseIf tipo_carga = "CON_FACTURA" Then
            If Trim(UCase(oRS3.Fields(0).Name)) <> "OC" Then
                msj = msj & " El archivo no tiene la columna 'OC' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(1).Name)) <> "PEDIDO" Then
                msj = msj & " El archivo no tiene la columna 'PEDIDO' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(2).Name)) <> "FACTURA" Then
                msj = msj & " El archivo no tiene la columna 'FACTURA' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(3).Name)) <> "COMPLEMENTO" Then
                msj = msj & " El archivo no tiene la columna 'COMPLEMENTO' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(4).Name)) <> "N.CLIENTE" And Trim(UCase(oRS3.Fields(4).Name)) <> "N#CLIENTE" Then
                msj = msj & " El archivo no tiene la columna 'N.CLIENTE' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(5).Name)) <> "CAJAS" Then
                msj = msj & " El archivo no tiene la columna 'CAJAS' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(6).Name)) <> "TARIMAS" Then
                msj = msj & " El archivo no tiene la columna 'TARIMAS' o se encuentra en otra posici�n. " & vbCrLf
            End If
            If Trim(UCase(oRS3.Fields(7).Name)) <> "FECHA" Then
                msj = msj & " El archivo no tiene la columna 'FECHA' o se encuentra en otra posici�n. " & vbCrLf
            End If
        End If
    End If
    
    ValidarLayOut = msj
End Function
'<<CHG-DESA-27022024-01
Private Function Obtiene_DIECLAVE_CCLCLAVE(ByVal NUM_DIR_CLIENTE As String, ByVal CliClef As String, ByRef dieclave As Long, ByRef cclclave As Long)
        Dim SQL_DIE_CCL As String, res As Boolean
        
        res = False
        SQL_DIE_CCL = ""
        SQL_DIE_CCL = SQL_DIE_CCL & " SELECT  DIE.DIE_CCLCLAVE CCLCLAVE, DIE.DIECLAVE DIECLAVE" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & " FROM    EDIRECCION_ENTR_CLIENTE_LIGA DEC" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "     INNER JOIN  EDIRECCIONES_ENTREGA DIE" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "         ON  DEC.DEC_DIECLAVE    =   DIE.DIECLAVE" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & " WHERE   DIE.DIE_STATUS          =   1" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "     AND EXISTS  (" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "                     SELECT  NULL" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "                     FROM    EDESTINOS_POR_RUTA DER" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "                     WHERE   DER.DER_VILCLEF =   DIE.DIEVILLE" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "                         AND NVL(DER.DER_ALLCLAVE, 1)    >   0" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "                         AND DER.DER_TIPO_ENTREGA    NOT IN  ('INSEGURO', 'INVALIDO')" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "                         AND SF_LOGIS_CLIENTE_RESTRIC(DEC.DEC_CLICLEF, DER.DER_TIPO_ENTREGA) =   1" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "     )" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "     AND DEC.DEC_NUM_DIR_CLIENTE =   '" & NUM_DIR_CLIENTE & "'" & vbCrLf
        SQL_DIE_CCL = SQL_DIE_CCL & "     AND DEC.DEC_CLICLEF         =   '" & CliClef & "'" & vbCrLf
        
        rs3.Open SQL_DIE_CCL
        If Not rs3.EOF Then
                cclclave = rs3.Fields("CCLCLAVE")
                dieclave = rs3.Fields("DIECLAVE")
                res = True
        End If
        rs3.Close
        
        Obtiene_DIECLAVE_CCLCLAVE = res
End Function

Private Function Valida_CCLCLAVE(ByVal cclclave As Long)
        
        Dim SQL_VAL_CCL As String
        Dim res As Boolean
        res = False
        SQL_VAL_CCL = ""
                
        '<<<--CHG-DESA-27022024-01
        'SQL_VAL_CCL = SQL_VAL_CCL & " SELECT 1" & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & " FROM WEB_CLIENT_CLIENTE WCCL " & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & " WHERE WCCL.WCCLCLAVE  = '" & cclclave & "'" & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & "   AND EXISTS ( " & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & "       SELECT NULL FROM EDESTINOS_POR_RUTA DER " & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & "        WHERE DER.DER_VILCLEF = WCCL_VILLE " & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & "          AND NVL(DER.DER_ALLCLAVE, 1) > 0 " & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & "          AND DER.DER_TIPO_ENTREGA NOT IN ('INSEGURO', 'INVALIDO') " & vbCrLf
        'SQL_VAL_CCL = SQL_VAL_CCL & "          AND SF_LOGIS_CLIENTE_RESTRIC(WCCL.WCCL_CLICLEF, DER.DER_TIPO_ENTREGA) = 1) " & vbCrLf
        
                
        SQL_VAL_CCL = " SELECT 1 " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & " FROM ECLIENT_CLIENTE CCL  " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & " WHERE 1=1 " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & " AND CCL.CCLCLAVE  = '" & cclclave & "' " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & " AND EXISTS (  " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & " SELECT NULL FROM EDESTINOS_POR_RUTA DER  " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & " WHERE DER.DER_VILCLEF = CCL_VILLE  " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & "   AND NVL(DER.DER_ALLCLAVE, 1) > 0  " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & "   AND DER.DER_TIPO_ENTREGA NOT IN ('INSEGURO', 'INVALIDO')  " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & "   /*AND SF_LOGIS_CLIENTE_RESTRIC(WCCL.WCCL_CLICLEF, DER.DER_TIPO_ENTREGA) = 1*/ " & vbCrLf
        SQL_VAL_CCL = SQL_VAL_CCL & "   )       " & vbCrLf
        'CHG-DESA-27022024-01-->>>
        
        rs3.Open SQL_VAL_CCL
        
        If Not rs3.EOF Then
                res = True
        Else
                res = False
        End If
        rs3.Close
        
        Valida_CCLCLAVE = res
End Function
'CHG-DESA-27022024-01>>
