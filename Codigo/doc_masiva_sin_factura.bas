Attribute VB_Name = "doc_masiva_sin_factura"
Option Explicit
Option Base 0


Private colSheets As New Collection


Sub doc_masiva_sin_fact(Archivo As String, cliente As String, correo_electronico As String, mi_disclef As String, Optional idCron As String = "")
        On Error GoTo catch
        
        ' ' ' ' '
        'Declaracion de variables:
        Dim My_excel As Excel.Application
        Dim oConn As ADODB.Connection
        Dim pestana_encabezados As String
        Dim HDR As String
        Dim usuario As String
        Dim usuario_can As String
        Dim Res As String
        Dim allclave_ori As Integer
        Dim allclave_dest As Integer
        Dim msg As String
        Dim oRS As New ADODB.Recordset
        
        Dim tmp_dest As String
        Dim ccl_clave As String
        Dim die_clave As String
        Dim cant_nuis As Double
        Dim id_factura As Integer
        Dim id_cdad_bultos As Integer
        
        Dim col_referencia As Integer
        Dim col_n_destinatario As Integer
        Dim col_bultos_totales As Integer
        Dim col_bultos_granel As Integer
        Dim col_tarimas As Integer
        Dim col_bultos_constitutivos As Integer
        Dim col_fecha As Integer
        Dim col_valor_mercancia As Integer
        Dim col_condiciones_entrega As Integer
        Dim col_observaciones As Integer
        
        Dim s_Referencia As String
        Dim s_Destinatario As String
        Dim i_BultosTotales As Double
        Dim i_BultosGranel As Double
        Dim i_Tarimas As Double
        Dim i_BultosConstitutivos As Double
        Dim d_Fecha As String
        Dim i_ValorMercancia As Double
        Dim s_CondicionesEntrega As String
        Dim s_Observaciones As String
        
        Dim SQL As String
        Dim iNUI As Double
        Dim lst_NUIs_insertados As String
        Dim lst_REFs_insertadas As String
        
        
        
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
        
        s_Referencia = ""
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
        Call obtener_cedis_x_remitente(cliente, mi_disclef, allclave_ori)
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
        Set colSheets = GetAllXLSheetNames_MASIVE(Archivo, False)
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
                If validar_destinatario(NVL(oRS.Fields(col_n_destinatario)), cliente) = -1 Then
                        msg = msg & "El Destinatario " & NVL(oRS.Fields(col_n_destinatario)) & " en la linea " & oRS.AbsolutePosition + 1 & " es incorrecto." & vbCrLf
                End If
                If validar_bultos_totales(NVL(oRS.Fields(col_bultos_totales)), NVL(oRS.Fields(col_bultos_granel)), NVL(oRS.Fields(col_tarimas)), NVL(oRS.Fields(col_bultos_constitutivos))) = False Then
                        msg = msg & "La cantidad de bultos totales en la linea " & oRS.AbsolutePosition & " no coincide con la cantidad de tarimas y los bultos a granel." & vbCrLf
                End If
                If validar_cdad_bultos_granel(NVL(oRS.Fields(col_bultos_granel)), cliente) = False Then
                        msg = msg & "La cantidad de bultos granel en la linea " & oRS.AbsolutePosition & " no es correcta." & vbCrLf
                End If
                If validar_cdad_tarimas(NVL(oRS.Fields(col_tarimas))) = False Then
                        msg = msg & "La cantidad de tarimas en la linea " & oRS.AbsolutePosition & " no es correcta." & vbCrLf
                End If
                If validar_bultos_por_tarima(NVL(oRS.Fields(col_bultos_constitutivos))) = False Then
                        msg = msg & "La cantidad de bultos por tarimas en la linea " & oRS.AbsolutePosition & " no es correcta." & vbCrLf
                End If
                If validar_valor_mercancia(NVL(oRS.Fields(col_valor_mercancia)), cliente) = False Then
                        msg = msg & "El valor de la mercancía en la linea " & oRS.AbsolutePosition & " no es correcto." & vbCrLf
                End If
                If validar_observaciones(NVL(oRS.Fields(col_observaciones))) = False Then
                        msg = msg & "La cantidad maxima de caracteres que puede tener el campo observaciones en la linea " & oRS.AbsolutePosition & " es de 80." & vbCrLf
                End If
                
                If tmp_dest <> NVL(oRS.Fields(col_n_destinatario)) Then
                        cant_nuis = cant_nuis + 1
                        tmp_dest = NVL(oRS.Fields(col_n_destinatario))
                End If
                
                oRS.MoveNext
        Loop
        oRS.Close
        
        If validar_cantidad_nuis_disponibles(cliente, cant_nuis) = False Then
                msg = msg & "La cantidad de NUI's disponibles es menor a la cantidad de NUI's necesarios para procesar este archivo." & vbCrLf
        End If
        
        Call log_SQL("doc_masiva_sin_fact", "terminan validaciones", cliente)

        
        If msg <> "" Then
                Call log_SQL("doc_masiva_sin_fact", "error de carga", cliente)
                Call notifica_error(cliente, correo_electronico, Archivo, msg)
        Else
                ''''
                'Proceso de documentacion:
                Call log_SQL("doc_masiva_sin_fact", "inicia documentacion", cliente)
                tmp_dest = ""
                cant_nuis = 0
                oRS.Open "Select * from [" & pestana_encabezados & "] order by 2,3 ", oConn, adOpenStatic, adLockOptimistic
                Do While Not oRS.EOF
                        'Agrupar NUI's por Destinatario:
                        If tmp_dest = "" Then
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
                                
                                Call obtener_destinatario(cliente, tmp_dest, ccl_clave, die_clave, allclave_dest)
                                
                                '''''''''
                                iNUI = Documenta_NUI(cliente, usuario, s_CondicionesEntrega, s_Observaciones, s_Referencia, mi_disclef, ccl_clave, _
                                die_clave, i_ValorMercancia, allclave_ori, allclave_dest, i_BultosTotales, i_Tarimas, i_BultosConstitutivos, i_BultosGranel)
                                '''''''''
                                
                                lst_NUIs_insertados = lst_NUIs_insertados & ", " & iNUI & "(" & s_Referencia & ")" & vbCrLf
                                
                                
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
                        
                        ''''' Documenta Ultimo NUI:
                        If oRS.EOF = True Then
                            iNUI = Documenta_NUI(cliente, usuario, s_CondicionesEntrega, s_Observaciones, s_Referencia, mi_disclef, ccl_clave, _
                            die_clave, i_ValorMercancia, allclave_ori, allclave_dest, i_BultosTotales, i_Tarimas, i_BultosConstitutivos, i_BultosGranel)
                            
                            lst_NUIs_insertados = vbCrLf & lst_NUIs_insertados & ", " & iNUI & "(" & s_Referencia & ")" & vbCrLf
                        End If
                        
                Loop
                oRS.Close
                Call log_SQL("doc_masiva_sin_fact", "termina documentacion", cliente)
                
                Call notifica_exito(cliente, correo_electronico, Archivo, cant_nuis, lst_NUIs_insertados)
                borrar_id_cron (idCron)
        End If
catch:
End Sub

Function Documenta_NUI(cliente As String, usuario As String, s_CondicionesEntrega As String, s_Observaciones As String, s_Referencia As String, mi_disclef As String, ccl_clave As String, _
die_clave As String, i_ValorMercancia As Double, allclave_ori As Integer, allclave_dest As Integer, i_BultosTotales As Double, i_Tarimas As Double, i_BultosConstitutivos As Double, i_BultosGranel As Double) As Double


'Ciclos determinantes: Documentar valores acumulados y reiniciar variables:
        Dim iNUI As Double
        Dim SQL As String
        s_CondicionesEntrega = obtener_prepagado_por_cobrar(cliente)
        s_Observaciones = obtener_dice_contener(cliente)
        iNUI = obtener_nui_disponible(cliente)
        
        Debug.Print iNUI & vbCrLf
        
        SQL = " UPDATE  WEB_LTL " & vbCrLf
        SQL = SQL & "   SET " & vbCrLf
        SQL = SQL & "           WELSTATUS                              =       1 " & vbCrLf
        SQL = SQL & "           ,DATE_CREATED                   =       SYSDATE " & vbCrLf
        SQL = SQL & "           ,MODIFIED_BY                    =       '" & usuario & "' " & vbCrLf
        SQL = SQL & "           ,WEL_COLLECT_PREPAID    =       '" & s_CondicionesEntrega & "' " & vbCrLf
        SQL = SQL & "           ,WELOBSERVACION                 =       SUBSTR('" & s_Observaciones & "',1,1999) " & vbCrLf
        
        If s_Referencia = "" Then
                SQL = SQL & "           ,WELFACTURA     =       '_PENDIENTE_' " & vbCrLf
        Else
                SQL = SQL & "           ,WELFACTURA     =       '" & s_Referencia & "' " & vbCrLf
        End If
        If mi_disclef <> "" Then
                SQL = SQL & "           ,WEL_DISCLEF    =       '" & mi_disclef & "' " & vbCrLf
        End If
        If ccl_clave <> "" Then
                SQL = SQL & "           ,WEL_CCLCLAVE   =       '" & ccl_clave & "' " & vbCrLf
        End If
        If die_clave <> "" Then
                SQL = SQL & "           ,WEL_DIECLAVE   =       '" & die_clave & "' " & vbCrLf
        End If
        If i_ValorMercancia > 0 Then
                SQL = SQL & "           ,WELIMPORTE     =       '" & i_ValorMercancia & "' " & vbCrLf
        End If
        If allclave_ori <> -1 Then
                SQL = SQL & "           ,WEL_ALLCLAVE_ORI       =       '" & allclave_ori & "' " & vbCrLf
        End If
        If allclave_dest <> -1 Then
                SQL = SQL & "           ,WEL_ALLCLAVE_DEST      =       '" & allclave_dest & "' " & vbCrLf
        End If
        If i_BultosTotales >= 0 Then
                SQL = SQL & "           ,WEL_CDAD_BULTOS        =       '" & i_BultosTotales & "' " & vbCrLf
        End If
        If i_Tarimas >= 0 Then
                SQL = SQL & "           ,WEL_CDAD_TARIMAS       =       '" & i_Tarimas & "' " & vbCrLf
        End If
        If i_BultosConstitutivos >= 0 Then
                SQL = SQL & "           ,WEL_CAJAS_TARIMAS      =       '" & i_BultosConstitutivos & "' " & vbCrLf
        End If
        If i_BultosGranel >= 0 Then
                SQL = SQL & "           ,WELCDAD_CAJAS  =       '" & i_BultosGranel & "' " & vbCrLf
        End If
        
        SQL = SQL & " WHERE     WELCLAVE = '" & iNUI & "' " & vbCrLf
        Db_link_orfeo.Execute SQL
        
        
        SQL = ""
        SQL = SQL & " UPDATE    WEB_TRACKING_STAGE " & vbCrLf
        SQL = SQL & "   SET      USR_DOC                                =       '" & usuario & "' " & vbCrLf
        SQL = SQL & "           ,FECHA_DOCUMENTACION    =       SYSDATE " & vbCrLf
        SQL = SQL & " WHERE      NUI                                    =       '" & iNUI & "' " & vbCrLf
        Db_link_orfeo.Execute SQL
        
        
        If i_Tarimas > 0 Then
                SQL = ""
                SQL = SQL & " INSERT INTO       TB_LOGIS_WPALETA_LTL " & vbCrLf
                SQL = SQL & "   ( " & vbCrLf
                SQL = SQL & "            WPLCLAVE ,WPL_WELCLAVE " & vbCrLf
                SQL = SQL & "           ,WPL_IDENTICAS ,WPL_TPACLAVE " & vbCrLf
                SQL = SQL & "           ,WPLLARGO ,WPLANCHO ,WPLALTO " & vbCrLf
                SQL = SQL & "           ,WPL_CDAD_EMPAQUES_X_BULTO ,WPL_BULTO_TPACLAVE " & vbCrLf
                SQL = SQL & "           ,CREATED_BY ,DATE_CREATED " & vbCrLf
                SQL = SQL & "   ) " & vbCrLf
                SQL = SQL & "   VALUES " & vbCrLf
                SQL = SQL & "           ( " & vbCrLf
                SQL = SQL & "                    SEQ_WPALETA_LTL.nextval ,'" & iNUI & "' " & vbCrLf
                SQL = SQL & "                   ,'" & i_Tarimas & "' ,1 " & vbCrLf
                SQL = SQL & "                   ,0 ,0 ,0 " & vbCrLf
                SQL = SQL & "                   , '" & i_BultosConstitutivos & "' ,9 " & vbCrLf
                SQL = SQL & "                   , '" & usuario & "' ,SYSDATE " & vbCrLf
                SQL = SQL & "           ) " & vbCrLf
                Db_link_orfeo.Execute SQL
        End If
        
        If i_BultosGranel > 0 Then
                SQL = ""
                SQL = SQL & " INSERT INTO       TB_LOGIS_WPALETA_LTL " & vbCrLf
                SQL = SQL & "   ( " & vbCrLf
                SQL = SQL & "            WPLCLAVE ,WPL_WELCLAVE " & vbCrLf
                SQL = SQL & "           ,WPL_IDENTICAS ,WPL_TPACLAVE " & vbCrLf
                SQL = SQL & "           ,WPLLARGO ,WPLANCHO ,WPLALTO " & vbCrLf
                SQL = SQL & "           ,CREATED_BY ,DATE_CREATED " & vbCrLf
                SQL = SQL & "   ) " & vbCrLf
                SQL = SQL & "   VALUES " & vbCrLf
                SQL = SQL & "           ( " & vbCrLf
                SQL = SQL & "                    SEQ_WPALETA_LTL.nextval ,'" & iNUI & "' " & vbCrLf
                SQL = SQL & "                   ,'" & i_BultosGranel & "' ,9 " & vbCrLf
                SQL = SQL & "                   ,0 ,0 ,0 " & vbCrLf
                SQL = SQL & "                   , '" & usuario & "' ,SYSDATE " & vbCrLf
                SQL = SQL & "           ) " & vbCrLf
                Db_link_orfeo.Execute SQL
        End If
        
        Call registrar_segundos_envios(iNUI, cliente, usuario)
        Call registrar_recol_domicilio(iNUI, cliente, usuario)
        Call CHECK_VALID_LTL(iNUI)
        
        
        Documenta_NUI = iNUI

End Function

