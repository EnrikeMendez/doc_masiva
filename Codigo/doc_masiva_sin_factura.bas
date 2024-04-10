Attribute VB_Name = "doc_masiva_sin_factura"
Option Explicit
Option Base 0

Sub doc_masiva_sin_fact(Archivo As String, cliente As String, correo_electronico As String, mi_disclef As String, idCron As String)
	On Error GoTo catch
	
	' ' ' ' '
	'Declaracion de variables:
	Private My_excel As Excel.Application
	Private oConn As ADODB.Connection
	Private pestana_encabezados As String
	Private HDR As String
	Private usuario As String
	Private usuario_can As String
	Private Res As String
	Private allclave_ori As Integer
	Private allclave_dest As Integer
	Private msg As String
	Private oRS As New ADODB.Recordset
	
	Private tmp_dest As String
	Private ccl_clave As String
	Private die_clave As String
	Private cant_nuis As Double
	Private id_factura As Integer
	Private id_cdad_bultos As Integer
	
	Private col_referencia As Integer
	Private col_n_destinatario As Integer
	Private col_bultos_totales As Integer
	Private col_bultos_granel As Integer
	Private col_tarimas As Integer
	Private col_bultos_constitutivos As Integer
	Private col_fecha As Integer
	Private col_valor_mercancia As Integer
	Private col_condiciones_entrega As Integer
	Private col_observaciones As Integer
	
	Private s_Referencia As String
	Private s_Destinatario As String
	Private i_BultosTotales As Double
	Private i_BultosGranel As Double
	Private i_Tarimas As Double
	Private i_BultosConstitutivos As Double
	Private d_Fecha As String
	Private i_ValorMercancia As Double
	Private s_CondicionesEntrega As String
	Private s_Observaciones As String
	
	Private SQL As String
	Private iNUI As Double
	Private lst_NUIs_insertados As String
	Private lst_REFs_insertadas As String
	
	
	
	Call log_SQL("doc_masiva_sin_fact", "inicio", cliente)
	' ' ' ' '
	'Inicializar variables:
	'Call init_var
	mi_disclef = Trim(mi_disclef)
	cliente = Trim(cliente)
	col_referencia = 0
	col_n_destinatario = 1
	col_bultos_totales = 2
	col_bultos_granel = 3
	col_tarimas = 4
	col_bultos_constitutivos = 5
	col_fecha = 6
	col_valor_mercancia = 7
	col_condiciones_entrega = 8
	col_observaciones = 9
	
	s_Referencia= ""
	s_Destinatario = ""
	i_BultosTotales = 0
	i_BultosGranel = 0
	i_Tarimas = 0
	i_BultosConstitutivos = 0
	d_Fecha = ""
	i_ValorMercancia = 0
	s_CondicionesEntrega = ""
	s_Observaciones = ""
	allclave_ori = -1
	allclave_dest = -1
	iNUI = -1
	


	' ' ' ' '
	'Funciones iniciales:
	obtener_cedis_x_remitente(cliente, mi_disclef, allclave_ori)
	usuario = obtener_nombre_usuario(Split(Split(Archivo, "temp")(1), "\")(1))
	'usuario_can = obtener_nombre_usuario("CAN_" & Split(Split(Archivo, "temp")(1), "\")(1))


	' ' ' ' '
	'Abrir archivo Excel:
	Set My_excel = New Excel.Application
	HDR = "Yes"
	Archivo = "\\" & Split(Archivo, "|")(1) & Replace(Split(Archivo, "|")(0), Split(Split(Archivo, "|")(0), "\")(0), "")
	Set oConn = New ADODB.Connection
	oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
				"Data Source=" & Archivo & ";" & _
				"Extended Properties=""Excel 8.0;HDR=" & HDR & ";IMEX=1;"""
	Set colSheets = GetAllXLSheetNames_UNIQUE(Archivo, False)
	pestana_encabezados = colSheets.Item(1)
	Call log_SQL("doc_masiva_sin_fact", "excel abierto", cliente)


	' ' ' ' '
	'Validaciones:
	msg = ""
	SQL = ""
	tmp_dest = ""
	cant_nuis = 0
	lst_NUIs_insertados = ""
	lst_REFs_insertadas = ""
	oRS.Open "Select * from [" & pestana_encabezados & "] order by 3,2 ", oConn, adOpenStatic, adLockOptimistic
	Call log_SQL("doc_masiva_sin_fact", "ocbd excel abierto", cliente)

	Do While Not oRS.EOF
		If validar_destinatario(NVL(oRS.Fields(col_n_destinatario)),cliente) <> -1 Then
			msg = msg & "El Destinatario en la linea " & oRS.AbsolutePosition + 1 & " es incorrecto." & vbCrLf
		End If
		If validar_bultos_totales(NVL(oRS.Fields(col_bultos_totales)),NVL(oRS.Fields(col_bultos_granel)),NVL(oRS.Fields(col_tarimas)),NVL(oRS.Fields(col_bultos_constitutivos)),cliente) = False Then
			msg = msg & "La cantidad de bultos totales en la linea " & oRS.AbsolutePosition + 1 & " no coincide con la cantidad de tarimas y los bultos a granel." & vbCrLf
		End If
		If validar_cdad_bultos_granel(NVL(oRS.Fields(col_bultos_granel)),cliente) = False Then
			msg = msg & "La cantidad de bultos granel en la linea " & oRS.AbsolutePosition + 1 & " no es correcta." & vbCrLf
		End If
		If validar_cdad_tarimas(NVL(oRS.Fields(col_tarimas))) = False Then
			msg = msg & "La cantidad de tarimas en la linea " & oRS.AbsolutePosition + 1 & " no es correcta." & vbCrLf
		End If
		If validar_bultos_por_tarima(NVL(oRS.Fields(col_bultos_constitutivos))) = False Then
			msg = msg & "La cantidad de bultos por tarimas en la linea " & oRS.AbsolutePosition + 1 & " no es correcta." & vbCrLf
		End If
		If validar_valor_mercancia(NVL(oRS.Fields(col_valor_mercancia)),cliente) = False Then
			msg = msg & "El valor de la mercancía en la linea " & oRS.AbsolutePosition + 1 & " no es correcto." & vbCrLf
		End If
		If validar_observaciones(NVL(oRS.Fields(col_observaciones)),cliente) = False Then
			msg = msg & "La cantidad maxima de caracteres que puede tener el campo observaciones en la linea " & oRS.AbsolutePosition + 1 & " es de 80." & vbCrLf
		End If
		
		If tmp_dest <> NVL(oRS.Fields(col_n_destinatario)) Then
			cant_nuis = cant_nuis + 1
			tmp_dest = NVL(oRS.Fields(col_n_destinatario))
		End If
		
		oRS.MoveNext
	Loop
	oRS.Close
	
	If validar_cantidad_nuis_disponibles(cliente,cant_nuis) = False Then
		msg = msg & "La cantidad de NUI's disponibles es menor a la cantidad de NUI's necesarios para procesar este archivo." & vbCrLf
	End If
	
	Call log_SQL("doc_masiva_sin_fact", "terminan validaciones", cliente)

	
	If msg <> "" Then
		Call log_SQL("doc_masiva_sin_fact", "error de carga", cliente)
		notifica_error(cliente, correo_electronico, Archivo, msg)
	Else
		''''
		'Proceso de documentacion:
		Call log_SQL("doc_masiva_sin_fact", "inicia documentacion", cliente)
		tmp_dest = ""
		oRS.Open "Select * from [" & pestana_encabezados & "] order by 3,2 ", oConn, adOpenStatic, adLockOptimistic
		Do While Not oRS.EOF
			'Agrupar NUI's por Destinatario:
			If tmp_dest = "" then
				'Primer ciclo
				s_Referencia = NVL(oRS.Fields(col_referencia))
				s_Destinatario = NVL(oRS.Fields(col_n_destinatario))
				i_BultosTotales = NVL_num(oRS.Fields(col_bultos_totales))
				i_BultosGranel = NVL_num(oRS.Fields(col_bultos_granel))
				i_Tarimas = NVL_num(oRS.Fields(col_tarimas))
				i_BultosConstitutivos = NVL_num(oRS.Fields(col_bultos_constitutivos))
				d_Fecha = NVL(oRS.Fields(col_fecha))
				i_ValorMercancia = NVL_num(oRS.Fields(col_valor_mercancia))
				s_Observaciones = NVL(oRS.Fields(col_observaciones))
				
				tmp_dest = s_Destinatario
			ElseIf tmp_dest = NVL(oRS.Fields(col_n_destinatario)) Then
				'Ciclos siguientes: Acumular valores y concatenar cadenas:
				i_BultosTotales = i_BultosTotales + NVL_num(oRS.Fields(col_bultos_totales))
				i_BultosGranel = i_BultosGranel + NVL_num(oRS.Fields(col_bultos_granel))
				i_Tarimas = i_Tarimas + NVL_num(oRS.Fields(col_tarimas))
				i_BultosConstitutivos = i_BultosConstitutivos + NVL_num(oRS.Fields(col_bultos_constitutivos))
				i_ValorMercancia = i_ValorMercancia + NVL_num(oRS.Fields(col_valor_mercancia))
				If s_Referencia = "" Then
					s_Referencia = NVL(oRS.Fields(col_referencia))
				Else
					s_Referencia = s_Referencia & ", " & NVL(oRS.Fields(col_referencia))
				End If
			Else
				'Ciclos determinantes: Documentar valores acumulados y reiniciar variables:
				obtener_destinatario(cliente, tmp_dest, ccl_clave, die_clave, allclave_dest)
				s_CondicionesEntrega = obtener_prepagado_por_cobrar(cliente)
				s_Observaciones = obtener_dice_contener(cliente)
				iNUI = obtener_nui_disponible(cliente)

				SQL = " UPDATE	WEB_LTL " & vbCrLf
				SQL = SQL & " 	SET " & vbCrLf
				SQL = SQL & " 		,WELSTATUS				=	1 " & vbCrLf
				SQL = SQL & " 		 DATE_CREATED			=	SYSDATE " & vbCrLf
				SQL = SQL & " 		 MODIFIED_BY			=	'" & usuario & "' " & vbCrLf
				SQL = SQL & " 		,WEL_COLLECT_PREPAID	=	'" & s_CondicionesEntrega & "' " & vbCrLf
				SQL = SQL & " 		,WELOBSERVACION			=	SUBSTR('" & s_Observaciones & "',1,1999) " & vbCrLf
				
				If s_Referencia = "" Then
					SQL = SQL & " 		,WELFACTURA	=	'_PENDIENTE_' " & vbCrLf
				Else
					SQL = SQL & " 		,WELFACTURA	=	'" & s_Referencia & "' " & vbCrLf
				End If
				If mi_disclef <> "" Then
					SQL = SQL & " 		,WEL_DISCLEF	=	'" & mi_disclef & "' " & vbCrLf
				End If
				If ccl_clave <> "" Then
					SQL = SQL & " 		,WEL_CCLCLAVE	=	'" & ccl_clave & "' " & vbCrLf
				End If
				If die_clave <> "" Then
					SQL = SQL & " 		,WEL_DIECLAVE	=	'" & die_clave & "' " & vbCrLf
				End If
				If i_ValorMercancia > 0 Then
					SQL = SQL & " 		,WELIMPORTE	=	'" & i_ValorMercancia & "' " & vbCrLf
				End If
				If allclave_ori <> -1 Then
					SQL = SQL & " 		,WEL_ALLCLAVE_ORI	=	'" & allclave_ori & "' " & vbCrLf
				End If
				If allclave_dest <> -1 Then
					SQL = SQL & " 		,WEL_ALLCLAVE_DEST	=	'" & allclave_dest & "' " & vbCrLf
				End If
				If i_BultosTotales >= 0 Then
					SQL = SQL & " 		,WEL_CDAD_BULTOS	=	'" & i_BultosTotales & "' " & vbCrLf
				End If
				If i_Tarimas >= 0 Then
					SQL = SQL & " 		,WEL_CDAD_TARIMAS	=	'" & i_Tarimas & "' " & vbCrLf
				End If
				If i_BultosConstitutivos >= 0 Then
					SQL = SQL & " 		,WEL_CAJAS_TARIMAS	=	'" & i_BultosConstitutivos & "' " & vbCrLf
				End If
				If i_BultosGranel >= 0 Then
					SQL = SQL & " 		,WELCDAD_CAJAS	=	'" & i_BultosGranel & "' " & vbCrLf
				End If
				
				SQL = SQL & " WHERE	WELCLAVE = '" & iNUI & "' " & vbCrLf
				Db_link_orfeo.Execute SQL
				
				
				SQL = ""
				SQL = SQL & " UPDATE	WEB_TRACKING_STAGE " & vbCrLf
				SQL = SQL & " 	SET	 USR_DOC				=	'" & usuario & "' " & vbCrLf
				SQL = SQL & " 		,FECHA_DOCUMENTACION	=	SYSDATE " & vbCrLf
				SQL = SQL & " WHERE	 NUI					=	'" & iNUI & "' " & vbCrLf
				Db_link_orfeo.Execute SQL
				
				
				If i_Tarimas > 0 Then
					SQL = ""
					SQL = SQL & " INSERT INTO	TB_LOGIS_WPALETA_LTL " & vbCrLf
					SQL = SQL & " 	( " & vbCrLf
					SQL = SQL & " 		 WPLCLAVE ,WPL_WELCLAVE " & vbCrLf
					SQL = SQL & " 		,WPL_IDENTICAS ,WPL_TPACLAVE " & vbCrLf
					SQL = SQL & " 		,WPLLARGO ,WPLANCHO ,WPLALTO " & vbCrLf
					SQL = SQL & " 		,WPL_CDAD_EMPAQUES_X_BULTO ,WPL_BULTO_TPACLAVE " & vbCrLf
					SQL = SQL & " 		,CREATED_BY ,DATE_CREATED " & vbCrLf
					SQL = SQL & " 	) " & vbCrLf
					SQL = SQL & "  	VALUES " & vbCrLf
					SQL = SQL & "   	( " & vbCrLf
					SQL = SQL & "   	 	 SEQ_WPALETA_LTL.nextval ,'" & iNUI & "' " & vbCrLf
					SQL = SQL & "   	 	,'" & i_Tarimas & "' ,1 " & vbCrLf
					SQL = SQL & "   	 	,0 ,0 ,0 " & vbCrLf
					SQL = SQL & "   	 	 '" & i_BultosConstitutivos & "' ,9 " & vbCrLf
					SQL = SQL & "   	 	 '" & usuario & "' ,SYSDATE " & vbCrLf
					SQL = SQL & "   	) " & vbCrLf
					Db_link_orfeo.Execute SQL
				End If
				
				If i_BultosGranel > 0 Then
					SQL = ""
					SQL = SQL & " INSERT INTO	TB_LOGIS_WPALETA_LTL " & vbCrLf
					SQL = SQL & " 	( " & vbCrLf
					SQL = SQL & " 		 WPLCLAVE ,WPL_WELCLAVE " & vbCrLf
					SQL = SQL & " 		,WPL_IDENTICAS ,WPL_TPACLAVE " & vbCrLf
					SQL = SQL & " 		,WPLLARGO ,WPLANCHO ,WPLALTO " & vbCrLf
					SQL = SQL & " 		,CREATED_BY ,DATE_CREATED " & vbCrLf
					SQL = SQL & " 	) " & vbCrLf
					SQL = SQL & "   VALUES " & vbCrLf
					SQL = SQL & "   	( " & vbCrLf
					SQL = SQL & "   	 	 SEQ_WPALETA_LTL.nextval ,'" & iNUI & "' " & vbCrLf
					SQL = SQL & "   	 	,'" & i_BultosGranel & "' ,9 " & vbCrLf
					SQL = SQL & "   	 	,0 ,0 ,0 " & vbCrLf
					SQL = SQL & "   	 	 '" & usuario & "' ,SYSDATE " & vbCrLf
					SQL = SQL & "   	) " & vbCrLf
					Db_link_orfeo.Execute SQL
				End If
				
				registrar_segundos_envios(iNUI,cliente,usuario)
				registrar_recol_domicilio(iNUI,cliente,usuario)
				Call CHECK_VALID_LTL(iNUI)
				
				lst_NUIs_insertados = lst_NUIs_insertados & ", " & iNUI & "(" & s_Referencia & ")"
				
				
				s_Referencia = NVL(oRS.Fields(col_referencia))
				s_Destinatario = NVL(oRS.Fields(col_n_destinatario))
				i_BultosTotales = NVL_num(oRS.Fields(col_bultos_totales))
				i_BultosGranel = NVL_num(oRS.Fields(col_bultos_granel))
				i_Tarimas = NVL_num(oRS.Fields(col_tarimas))
				i_BultosConstitutivos = NVL_num(oRS.Fields(col_bultos_constitutivos))
				d_Fecha = NVL(oRS.Fields(col_fecha))
				i_ValorMercancia = NVL_num(oRS.Fields(col_valor_mercancia))
				
				cant_nuis = cant_nuis + 1
				tmp_dest = NVL(oRS.Fields(col_n_destinatario))
			End If
			
			oRS.MoveNext
		Loop
		oRS.Close
		Call log_SQL("doc_masiva_sin_fact", "termina documentacion", cliente)
		
		notifica_exito(cliente, correo_electronico, Archivo, cant_nuis, lst_NUIs_insertados)
		borrar_id_cron(idCron)
	End If
catch:
End Sub