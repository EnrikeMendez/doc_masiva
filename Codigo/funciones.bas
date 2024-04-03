Attribute VB_Name = "funciones"
Option Explicit

Private SQL As String
Private idStatus As Integer
Private ConnString As String
Private TDCDFACTURA As String
Private arrTmp() As String
Private arrTmp1() As String
Private arrTracking() As String
Private arrNvoStatus() As String

Private rst1 As New ADODB.Recordset
Private rst2 As New ADODB.Recordset
Private rst3 As New ADODB.Recordset
Private rst4 As New ADODB.Recordset
Private rsT5 As New ADODB.Recordset
Private rsT7 As New ADODB.Recordset
Private rsAVC As New ADODB.Recordset

Private tieneVAS As Boolean
Private statusCancelado As Boolean
Private statusStandby As Boolean
Private statusEntregado As Boolean

Private catEstatus() As String
Private arrInfoTalon() As String

Private c As Integer, i As Integer, j As Integer, k As Integer
Private stClase As String, stTexto As String, sTxtCatalogo As String
Private stStatus As String, cveStatus As String, tdCDclave As String
Private EstatusTalon As String, sNUI As String, stObservaciones As String

'Variables utilizadas para interpretar el estatus:
Private incidencia As String, fecha_entrega As String, last_entrada As String, ObtieneInfoBDs As Boolean
Dim tstResult As String

'Concatenar los resultados del query en una cadena
Private rs_str As New ADODB.Recordset
Private str_result As String


'   '   '   '   '   '   '
'   FUNCIÓN PRINCIPAL   '
'   '   '   '   '   '   '

'Función principal para obtener el estatus de la guía
''' Obtiene el estatus a partir del seguimiento que se le da a un talón/factura LTL y CD'''
Public Function ObtieneStatusTalon_txt(wTalonRastreo As String) As String
    On Error GoTo error_function

    'Inicialización de variables:
    Call init_var
	
    tstResult = ""
    stTexto = ""
    statusCancelado = EsGuiaCancelada(wTalonRastreo)
    statusStandby = GuiaEnStndBy(wTalonRastreo)
    
	If statusCancelado <> True And statusStandby <> True Then
        statusEntregado = GuiaEntregada(wTalonRastreo)
    End If

    If statusCancelado = True Then
        stClase = "rojo"
        stTexto = "Cancelado"
        stObservaciones = "Guia cancelada"
        cveStatus = "0"
        stStatus = "0"
        tstResult = "funcion cancelado"
    ElseIf statusStandby = True Then
        stClase = "naranja"
        stTexto = "StdBy"
        stObservaciones = "Guia en Stand By"
        cveStatus = "2"
        stStatus = "2"
        tstResult = "funcion standBy"
    ElseIf statusEntregado = True Then
        stClase = "verde"
        stTexto = "Entregado"
        stObservaciones = "Entrega al Cliente"
        cveStatus = "6"
        stStatus = "6"
        tstResult = "funcion entregado"
    Else
        stClase = "amarillo " & wTalonRastreo
        catEstatus = obtenerCatalogoEstatus()
        arrInfoTalon = obtenerInfoTalon(wTalonRastreo)

        If esArregloConElementos(arrInfoTalon) Then
            'Se obtiene la clave que sirve para obtener el Tracking de la guia LTL/CD:
            tdCDclave = arrInfoTalon(0, 15)
            tstResult = tstResult & " |tdCDclave:" & CStr(tdCDclave) & vbCrLf

            If CStr(tdCDclave) <> "" Then
                arrTracking = obtenerTrackingTalon(tdCDclave)

                If esArregloConElementos(arrTracking) Then
                    '====================================================================='
                    'Se obtienen los parámetros: incidencia, fecha_entrega y last_entrada.'
                    '====================================================================='
                    fecha_entrega = arrTracking(6, 0)
                    last_entrada = arrTracking(8, 0)
                    incidencia = arrTracking(9, 0)
                    ObtieneInfoBDs = True
                    
                    tstResult = tstResult & "|fecha_entrega:" & CStr(fecha_entrega)
                    tstResult = tstResult & "|last_entrada:" & CStr(last_entrada)
                    tstResult = tstResult & "|incidencia:" & CStr(incidencia) & vbCrLf

                    If NVL(arrTracking(3, 0)) = "DIRECTO" _
                        And (arrTracking(10, 0) = "N" Or (arrTracking(10, 0) = "S" And arrTracking(9, 0) = "4")) _
                        And arrTracking(9, 0) <> "5" Then
                        'no recuperar las reexpediciones o los VAS
                        incidencia = arrTracking(9, 0)
                        fecha_entrega = NVL(arrTracking(6, 0))
                    Else
                        If arrTracking(9, 0) = "5" Then
                            If i = UBound(arrTracking, 2) - 1 Then
                                incidencia = arrTracking(9, 0)
                                fecha_entrega = ""
                            End If
                        End If
                    End If

                    last_entrada = arrTracking(8, 0)

                    If arrTracking(11, 0) = "VAS" And arrTracking(9, 0) <> "0" Then
                        tieneVAS = True
                        tstResult = tstResult & "|VAS:" & CStr(arrTracking(11, 0))
                    Else
                        tieneVAS = False
                    End If
                Else
                    ObtieneInfoBDs = False
                End If
            Else
                tstResult = tstResult & "|sin tdCDclave"
                Debug.Print (Now & " No se puede obtener el tracking completo del talón debido a que no cuenta con tdCDclave. " & tdCDclave)
            End If


            '==========================='
            'Interpretación del estatus.'
            '==========================='
            If CStr(incidencia) = "0" And CStr(fecha_entrega) <> "" Then
                stClase = "verde"
                stTexto = "Entregado"
                stStatus = "6"
                cveStatus = "6"
            Else
                Select Case incidencia
                    Case "0"
                        If fecha_entrega <> "" Then
                            'entrega normal, no pasa nada
                            stClase = "verde"
                            stTexto = "Entregado"
                            stStatus = "6"
                            cveStatus = "6"
                        Else
                            stClase = "naranja"
                            stTexto = "En transito"
                        End If
                    Case "4"
                        stClase = "rojo"
                        stTexto = "No entregado"
                    Case "3"
                        stClase = "rojo"
                        stTexto = "Entrega incompleta"
                    Case Else
                        If last_entrada <> "24" Then
                            'No hubo entrada de rechazo todavia entonces el status esta en transito borramos la fecha de entrega:
                            fecha_entrega = ""
                        Else
                            stClase = "rojo"
                            stTexto = "Rechazado"
                        End If
                End Select
            End If

            If tieneVAS = True Then
                stClase = "rojo-claro"
                stTexto = "No   entregado (Intento de entrega fallido)"
                stStatus = "8"
                cveStatus = "8"
            End If
        Else
            tstResult = tstResult & "sqlSinElementos|"
            arrNvoStatus = obtenerEstatusSinDocumentar(wTalonRastreo)

            If esArregloConElementos(arrNvoStatus) Then
                tstResult = tstResult & "arrNvoStatusConElementos|"
                ObtieneInfoBDs = True
                cveStatus = arrNvoStatus(0, 0)

                If arrNvoStatus(0, 0) = "0" Then
                    If arrNvoStatus(1, 0) = "RESERVADO - CANCELADO " Then
                        stClase = "rojo-claro"
                        stTexto = "Reservado - Cancelado"
                    Else
                        stClase = "rojo"
                        stTexto = "Cancelado"
                    End If
                Else
                    If arrNvoStatus(0, 0) = "3" Then
                        stClase = "gris"
                        stTexto = "Reservado"
                    End If
                End If
            Else
                ObtieneInfoBDs = False
                tstResult = tstResult & "arrNvoStatusSinElementos|"
            End If
        End If
        tstResult = tstResult & "|incidencia:" & CStr(incidencia) & "|fecha_entrega:" & CStr(fecha_entrega)

        'Si la guía tiene incidencia, se mantiene el estatus de la incidencia, si tiene fecha de entrega también se mantiene, de lo contrario se procesa el nuevo estatus:
        If incidencia <> "3" And incidencia <> "4" And CStr(fecha_entrega) = "" Then
            ''' =========================================== '''
            '''  Nuevo proceso para interpretar el estatus  '''
            ''' =========================================== '''
            ' 1.- Obtener información del talón;
            ' 2.- Replicar el proceso de la pantalla Tracking;
            ' 3.- Ajustar los estatus de acuerdo a las reglas que están en el excel (8 eventos);
            ' 4.- Aplicar las reglas de los colores que se van a mostrar en la pantalla;
            ' NOTA: todo se debe basar en el texto que está en los registros de seguimiento que se encuentra en la BD's.

            If esArregloConElementos(arrTracking) Then
                stObservaciones = arrTracking(2, 0)
            End If
            tstResult = tstResult & "|stObservaciones:" & CStr(stObservaciones)

            If esArregloConElementos(arrTracking) Then
                For i = 0 To UBound(arrTracking, 2)
                    For j = 0 To UBound(arrTracking)
                        If j = 2 Or j = 5 Then
                            For k = 0 To UBound(catEstatus, 2)
                                If InStr(UCase(arrTracking(j, i)), UCase(catEstatus(k, 2))) > 0 Then
                                    idStatus = k
                                    cveStatus = UCase(catEstatus(k, 0))
                                    stStatus = ""
                                End If
                            Next
                        End If
                    Next
                Next

                If idStatus <> -1 Then
                    stTexto = catEstatus(idStatus, 1)
                    stClase = catEstatus(idStatus, 3)
                    sTxtCatalogo = catEstatus(idStatus, 2)
                End If
            End If
        End If

        If stStatus = "0" Then
            stClase = "rojo"
            stTexto = "Cancelado"
        Else
            If stStatus = "3" Then
                stClase = "gris"
                stTexto = "Reservado"
            End If
        End If


        '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '
        '   Obtener información del catálogo de acuerdo al estatus obtenido.    '
        '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '   '
        For k = 0 To UBound(catEstatus, 2)
            If InStr(UCase(stTexto), UCase(catEstatus(k, 2))) > 0 Then
                stTexto = catEstatus(k, 1)
                stClase = catEstatus(k, 3)
                sTxtCatalogo = catEstatus(k, 2)
                cveStatus = UCase(catEstatus(k, 0))
                tstResult = tstResult & "CveCatalogo2:" & cveStatus & "|"
            End If
        Next

        If stTexto = "" And sTxtCatalogo <> "" Then stTexto = sTxtCatalogo
    End If


    '   '   '   '   '   '   '   '   '
    '   PRESENTACIÓN DE RESULTADOS  '
    '   '   '   '   '   '   '   '   '
    ReDim arrResult(1, 5)
    EstatusTalon = "<td cveStatus='" & cveStatus & "' sStatus='" & stStatus & "' class='" & stClase & "' style='text-align:center;'>" & stTexto & "</td>"

    arrResult(1, 0) = stClase
    arrResult(1, 1) = stTexto
    arrResult(1, 2) = EstatusTalon
    arrResult(1, 3) = True
    arrResult(1, 4) = stObservaciones
    arrResult(1, 5) = cveStatus

    If wTalonRastreo = "" Then
        arrResult = Null
    End If

error_function:
    If Err.Number <> 0 Then
        Debug.Print (Now & " Error: " & wTalonRastreo & " " & Err.Description)
        tstResult = tstResult & "|ObtieneStatusTalon_txt:" & Err.Description
    End If

    Debug.Print (Now & " El talón " & wTalonRastreo & " tiene el estatus *" & stTexto & "*.")
    ObtieneStatusTalon_txt = stTexto
End Function

'   '   '   '   '   '   '
'   FUNCIONES INTERNAS  '
'   '   '   '   '   '   '
Private Sub init_var()
On Error GoTo Catch

    i = 0
    j = 0
    k = 0
    
    idStatus = -1
    incidencia = -1
    
    tieneVAS = False
    statusStandby = False
    statusCancelado = False
    statusEntregado = False
    
    cveStatus = "-"
    fecha_entrega = ""
    ObtieneInfoBDs = False
    
    stClase = "amarillo"
    stTexto = "Documentado"
    sTxtCatalogo = "Documentado"
    
    If esArregloConElementos(arrTracking) Then
        ReDim arrTracking(0)
    End If
    
        
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
    End If
    
    Set rst1 = Nothing
    Set rst1 = New ADODB.Recordset
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.LockType = adLockReadOnly
    rst1.ActiveConnection = Db_link_orfeo
    
    Set rst2 = Nothing
    Set rst2 = New ADODB.Recordset
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.LockType = adLockReadOnly
    rst2.ActiveConnection = Db_link_orfeo
    
    Set rst3 = Nothing
    Set rst3 = New ADODB.Recordset
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.LockType = adLockReadOnly
    rst3.ActiveConnection = Db_link_orfeo
    
    Set rst4 = Nothing
    Set rst4 = New ADODB.Recordset
    rst4.CursorLocation = adUseClient
    rst4.CursorType = adOpenForwardOnly
    rst4.LockType = adLockReadOnly
    rst4.ActiveConnection = Db_link_orfeo
    
    Set rsT5 = Nothing
    Set rsT5 = New ADODB.Recordset
    rsT5.CursorLocation = adUseClient
    rsT5.CursorType = adOpenForwardOnly
    rsT5.LockType = adLockReadOnly
    rsT5.ActiveConnection = Db_link_orfeo
    
    Set rsT7 = Nothing
    Set rsT7 = New ADODB.Recordset
    rsT7.CursorLocation = adUseClient
    rsT7.CursorType = adOpenForwardOnly
    rsT7.LockType = adLockReadOnly
    rsT7.ActiveConnection = Db_link_orfeo

    Set rsAVC = Nothing
    Set rsAVC = New ADODB.Recordset
    rsAVC.CursorLocation = adUseClient
    rsAVC.CursorType = adOpenForwardOnly
    rsAVC.LockType = adLockReadOnly
    rsAVC.ActiveConnection = Db_link_orfeo
    
    'Concatenar los resultados del query en una cadena
    Set rs_str = Nothing
    Set rs_str = New ADODB.Recordset
    rs_str.CursorLocation = adUseClient
    rs_str.CursorType = adOpenForwardOnly
    rs_str.LockType = adLockReadOnly
    rs_str.ActiveConnection = Db_link_orfeo
    
     ConnString = "PROVIDER=MSDAORA;" & _
             "DATA SOURCE=" & database_name & ";" & _
             "USER ID=" & user & ";PASSWORD=" & Password & ";"
    
End Sub

''' Retorna el catálogo de estatus que se estará manejando a nivel global para éste tema, los casos particulares se contemplan en la función principal.
Private Function obtenerCatalogoEstatus() As String()
On Error GoTo Catch
    SQL = " SELECT 1 AS No_Evento,  'Documentado' AS Estatus,    'Creacion de la' AS Observaciones,  'amarillo' AS Clase FROM DUAL   " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 2,'En Recoleccion','Recoleccion','naranja' FROM DUAL    " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 3,'En Transito','Entrada CEDIS Logis','naranja' FROM DUAL   " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 4,'En Transito','Expedicion','naranja' FROM DUAL    " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 5,'En Transito a destino final','Expedicion directa al cliente','naranja' FROM DUAL " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 6,'Entregado','Entrega al Cliente','verde' FROM DUAL   " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 7,'Intento de entrega fallido','Entrega incompleta','rojo' FROM DUAL   " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 8,'Intento de entrega fallido   (no   entregado)','Entrega al cliente con incidencia','rojo' FROM DUAL    " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 9,'No entregado','No entregado','rojo' FROM DUAL    " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 10,'Rechazado','Rechazado','rojo' FROM DUAL   " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 11,'Cancelado','Cancelado','rojo' FROM DUAL    " & vbCrLf
    SQL = SQL & " UNION " & vbCrLf
    SQL = SQL & " SELECT 12,'Reservado','Reservado','gris' FROM DUAL    " & vbCrLf
    
    rst1.Open SQL
    If Not rst1.EOF Then
        i = 0
        c = rst1.RecordCount
        ReDim arrTmp(c, 3)
        
        While Not rst1.EOF
            arrTmp(i, 0) = rst1.Fields("NO_EVENTO")
            arrTmp(i, 1) = rst1.Fields("ESTATUS")
            arrTmp(i, 2) = rst1.Fields("OBSERVACIONES")
            arrTmp(i, 3) = rst1.Fields("CLASE")
            i = i + 1
            rst1.MoveNext
        Wend
    End If
            
Catch:
    rst1.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerCatalogoEstatus:" & Err.Description
    End If

    obtenerCatalogoEstatus = arrTmp
End Function

'''Obtiene de BD's la información general del talón:
Private Function obtenerInfoTalon(wTalonRastreo As String) As String()
On Error GoTo Catch
    SQL = "SELECT TO_CHAR(WEL.WELCONS_GENERAL, 'FM0000000') || '-' ||GET_CLI_ENMASCARADO(WEL.WEL_CLICLEF) || DECODE(WEL_ORI.WELCLAVE, NULL, NULL, ' (talon ori: ' || TO_CHAR(WEL_ORI.WELCONS_GENERAL, 'FM0000000') || '-' ||GET_CLI_ENMASCARADO(WEL_ORI.WEL_CLICLEF) ||')') AS TALON " & vbCrLf
    SQL = SQL & "       , NVL(WEL.WEL_TALON_RASTREO, WEL.WEL_FIRMA) AS FIRMA " & vbCrLf
    SQL = SQL & "       , TO_CHAR( WEL.DATE_CREATED, 'DD/MM/YYYY HH24:MI') AS FECHA_CREACION " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAE_FECHA_RECOLECCION, 'DD/MM/YYYY HH24:MI') AS FECHA_RECOLECCION " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAEFECHALLEGADA, 'DD/MM/YYYY HH24:MI') AS FECHA_LLEGADA " & vbCrLf
    SQL = SQL & "       , WEL.WELRECOL_DOMICILIO AS RECOL_DOMICILIO " & vbCrLf
    SQL = SQL & "       , WEL.WELFACTURA AS FACTURA " & vbCrLf
    SQL = SQL & "       , WEL.WEL_CDAD_BULTOS AS CDAD_BULTOS " & vbCrLf
    SQL = SQL & "       , INITCAP(DIS.DISNOM) AS REMITENTE " & vbCrLf
    SQL = SQL & "       , InitCap(DISADRESSE1 || ' ' || ' ' || DISNUMEXT || '  ' || DISNUMINT || '  <br> ' ||DISADRESSE2 || DECODE(DISCODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DISCODEPOSTAL))  AS REMITENTE_DIREC " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_ORI.VILNOM || ' ('|| EST_ORI.ESTNOMBRE || ')') AS REMITENTE_ESTADO " & vbCrLf
    SQL = SQL & "       , INITCAP(NVL(DIE2.DIE_A_ATENCION_DE, DIE2.DIENOMBRE)) AS REMITENTE_1 " & vbCrLf
    SQL = SQL & "       , InitCap( DIE2.DIEADRESSE1|| ' ' || ' ' || DIE2.DIENUMEXT || '  ' || DIE2.DIENUMINT || '  <br> ' ||DIE2.DIEADRESSE2 || DECODE(DIE2.DIECODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DIE2.DIECODEPOSTAL)) AS REMITENTE_DIREC_1  " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_DEST.VILNOM || ' ('|| EST_DEST.ESTNOMBRE || ')') AS REMITENTE_ESTADO_1 " & vbCrLf
    SQL = SQL & "       , WEL.WELSTATUS AS ESTATUS " & vbCrLf
    SQL = SQL & "       , WEL.WEL_TDCDCLAVE AS TDCDCLAVE " & vbCrLf
    SQL = SQL & "       , 'LTL' AS TIPO " & vbCrLf
    SQL = SQL & "       , WEL.WEL_CLICLEF AS CLICLEF " & vbCrLf
    SQL = SQL & "       , WEL.WELOBSERVACION AS OBSERVACIONES " & vbCrLf
    SQL = SQL & "       , WEL.WELPESO AS PESO " & vbCrLf
    SQL = SQL & "       , WEL.WELVOLUMEN AS VOLUMEN " & vbCrLf
    SQL = SQL & "       , WEL.WELCLAVE AS NUI " & vbCrLf
    SQL = SQL & " FROM WEB_LTL WEL " & vbCrLf
    SQL = SQL & "       , EDIRECCIONES_ENTREGA DIE2 " & vbCrLf
    SQL = SQL & "       , EDISTRIBUTEUR DIS " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_ORI " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_ORI " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_DEST " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_DEST " & vbCrLf
    SQL = SQL & "       , ETRANS_DETALLE_CROSS_DOCK TDCD " & vbCrLf
    SQL = SQL & "       , ETRANSFERENCIA_TRADING TRA " & vbCrLf
    SQL = SQL & "       , ETRANS_ENTRADA TAE " & vbCrLf
    SQL = SQL & "       , WEB_LTL WEL_ORI " & vbCrLf
    SQL = SQL & " WHERE (WEL.WEL_FIRMA IN ('" & wTalonRastreo & "') OR WEL.WEL_TALON_RASTREO IN ('" & wTalonRastreo & "') ) " & vbCrLf
    SQL = SQL & "       AND DISCLEF = WEL.WEL_DISCLEF " & vbCrLf
    SQL = SQL & "       AND DIE2.DIECLAVE = WEL.WEL_DIECLAVE " & vbCrLf
    SQL = SQL & "       AND CIU_ORI.VILCLEF = DISVILLE " & vbCrLf
    SQL = SQL & "       AND EST_ORI.ESTESTADO = CIU_ORI.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND CIU_DEST.VILCLEF = DIE2.DIEVILLE " & vbCrLf
    SQL = SQL & "       AND EST_DEST.ESTESTADO = CIU_DEST.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND TDCDCLAVE(+) = WEL.WEL_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       AND TDCDSTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TRACLAVE(+) = WEL.WEL_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND TRASTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE(+) = WEL.WEL_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND WEL_ORI.WELCLAVE(+) = WEL.WEL_WELCLAVE " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE = TRACLAVE " & vbCrLf
    SQL = SQL & "UNION ALL " & vbCrLf
    SQL = SQL & " SELECT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA) " & vbCrLf
    SQL = SQL & "       , WCD.WCD_FIRMA " & vbCrLf
    SQL = SQL & "       , TO_CHAR( WCD.DATE_CREATED, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAE_FECHA_RECOLECCION, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAEFECHALLEGADA, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , 'n/a' " & vbCrLf
    SQL = SQL & "       , WCD.WCD_PEDIDO_CLIENTE " & vbCrLf
    SQL = SQL & "       , WCD.WCD_CDAD_BULTOS " & vbCrLf
    SQL = SQL & "       , INITCAP(DIS.DISNOM) REMITENTE " & vbCrLf
    SQL = SQL & "       , InitCap(DISADRESSE1 || ' ' || ' ' || DISNUMEXT || '  ' || DISNUMINT || '  <br> ' ||DISADRESSE2 || DECODE(DISCODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DISCODEPOSTAL)) " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_ORI.VILNOM || ' ('|| EST_ORI.ESTNOMBRE || ')') " & vbCrLf
    SQL = SQL & "       , INITCAP(CCL.CCL_NOMBRE || ' ' || NVL(DIE.DIE_A_ATENCION_DE, DIE.DIENOMBRE)) " & vbCrLf
    SQL = SQL & "       , InitCap( DIEADRESSE1|| ' ' || ' ' || DIENUMEXT || '  ' || DIENUMINT || '  <br> ' ||DIEADRESSE2 || DECODE(DIECODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DIECODEPOSTAL)) " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_DEST.VILNOM || ' ('|| EST_DEST.ESTNOMBRE || ')') " & vbCrLf
    SQL = SQL & "       , WCD.WCDSTATUS " & vbCrLf
    SQL = SQL & "       , WCD.WCD_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       , 'Cross Dock' " & vbCrLf
    SQL = SQL & "       , WCD_CLICLEF " & vbCrLf
    SQL = SQL & "       , WCD.WCDOBSERVACION " & vbCrLf
    SQL = SQL & "       , WCD.WCDPESO " & vbCrLf
    SQL = SQL & "       , WCD.WCDVOLUMEN " & vbCrLf
    SQL = SQL & "       , WCD.WCDCLAVE AS NUI " & vbCrLf
    SQL = SQL & " FROM WCROSS_DOCK WCD " & vbCrLf
    SQL = SQL & "       , EDIRECCIONES_ENTREGA DIE " & vbCrLf
    SQL = SQL & "       , ECLIENT_CLIENTE CCL " & vbCrLf
    SQL = SQL & "       , EDISTRIBUTEUR DIS " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_ORI " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_ORI " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_DEST " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_DEST " & vbCrLf
    SQL = SQL & "       , ETRANS_DETALLE_CROSS_DOCK TDCD " & vbCrLf
    SQL = SQL & "       , ETRANSFERENCIA_TRADING TRA " & vbCrLf
    SQL = SQL & "       , ETRANS_ENTRADA TAE " & vbCrLf
    SQL = SQL & " WHERE WCD.WCD_FIRMA IN ('" & wTalonRastreo & "') " & vbCrLf
    SQL = SQL & "       AND DISCLEF = WCD.WCD_DISCLEF " & vbCrLf
    SQL = SQL & "       AND DIECLAVE = NVL(NVL(TDCD_DIECLAVE_ENT, TDCD_DIECLAVE), WCD_DIECLAVE_ENTREGA) " & vbCrLf
    SQL = SQL & "       AND CCLCLAVE = NVL(TDCD_CCLCLAVE, WCD.WCD_CCLCLAVE) " & vbCrLf
    SQL = SQL & "       AND CIU_ORI.VILCLEF = DISVILLE " & vbCrLf
    SQL = SQL & "       AND EST_ORI.ESTESTADO = CIU_ORI.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND CIU_DEST.VILCLEF = DIEVILLE " & vbCrLf
    SQL = SQL & "       AND EST_DEST.ESTESTADO = CIU_DEST.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND TDCDCLAVE(+) = WCD.WCD_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       AND TDCDSTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TRACLAVE(+) = WCD.WCD_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND TRASTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE(+) = WCD.WCD_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE = TRACLAVE " & vbCrLf
    SQL = SQL & "UNION " & vbCrLf
    SQL = SQL & " SELECT TO_CHAR(WEL.WELCONS_GENERAL, 'FM0000000') || '-' ||GET_CLI_ENMASCARADO(WEL.WEL_CLICLEF) || DECODE(WEL_ORI.WELCLAVE, NULL, NULL, ' (talon ori: ' || TO_CHAR(WEL_ORI.WELCONS_GENERAL, 'FM0000000') || '-' ||GET_CLI_ENMASCARADO(WEL_ORI.WEL_CLICLEF) ||')') " & vbCrLf
    SQL = SQL & "       , NVL(WEL.WEL_TALON_RASTREO, WEL.WEL_FIRMA) AS WEL_FIRMA " & vbCrLf
    SQL = SQL & "       , TO_CHAR( WEL.DATE_CREATED, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAE_FECHA_RECOLECCION, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAEFECHALLEGADA, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , WEL.WELRECOL_DOMICILIO " & vbCrLf
    SQL = SQL & "       , WEL.WELFACTURA " & vbCrLf
    SQL = SQL & "       , WEL.WEL_CDAD_BULTOS " & vbCrLf
    SQL = SQL & "       , INITCAP(DIS.DISNOM) REMITENTE " & vbCrLf
    SQL = SQL & "       , InitCap(DISADRESSE1 || ' ' || ' ' || DISNUMEXT || '  ' || DISNUMINT || '  <br> ' ||DISADRESSE2 || DECODE(DISCODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DISCODEPOSTAL))  remitente_direc " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_ORI.VILNOM || ' ('|| EST_ORI.ESTNOMBRE || ')') " & vbCrLf
    SQL = SQL & "       , INITCAP(WCCL.WCCL_NOMBRE) " & vbCrLf
    SQL = SQL & "       , InitCap( WCCL_ADRESSE1|| ' ' || ' ' || WCCL_NUMEXT || '  ' || WCCL_NUMINT || '  <br> ' ||WCCL_ADRESSE2 || DECODE(WCCL_CODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || WCCL_CODEPOSTAL)) remitente_direc " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_DEST.VILNOM || ' ('|| EST_DEST.ESTNOMBRE || ')') " & vbCrLf
    SQL = SQL & "       , WEL.WELSTATUS " & vbCrLf
    SQL = SQL & "       , WEL.WEL_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       , 'LTL' " & vbCrLf
    SQL = SQL & "       , WEL.WEL_CLICLEF " & vbCrLf
    SQL = SQL & "       , WEL.WELOBSERVACION " & vbCrLf
    SQL = SQL & "       , WEL.WELPESO " & vbCrLf
    SQL = SQL & "       , WEL.WELVOLUMEN " & vbCrLf
    SQL = SQL & "       , WEL.WELCLAVE AS NUI " & vbCrLf
    SQL = SQL & " FROM WEB_LTL WEL " & vbCrLf
    SQL = SQL & "       , WEB_CLIENT_CLIENTE WCCL " & vbCrLf
    SQL = SQL & "       , EDISTRIBUTEUR DIS " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_ORI " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_ORI " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_DEST " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_DEST " & vbCrLf
    SQL = SQL & "       , ETRANS_DETALLE_CROSS_DOCK TDCD " & vbCrLf
    SQL = SQL & "       , ETRANSFERENCIA_TRADING TRA " & vbCrLf
    SQL = SQL & "       , ETRANS_ENTRADA TAE " & vbCrLf
    SQL = SQL & "       , WEB_LTL WEL_ORI " & vbCrLf
    SQL = SQL & " WHERE (WEL.WEL_FIRMA IN ('" & wTalonRastreo & "') OR WEL.WEL_TALON_RASTREO IN ('" & wTalonRastreo & "') ) " & vbCrLf
    SQL = SQL & "       AND DISCLEF = WEL.WEL_DISCLEF " & vbCrLf
    SQL = SQL & "       AND WCCLCLAVE = WEL.WEL_WCCLCLAVE " & vbCrLf
    SQL = SQL & "       AND CIU_ORI.VILCLEF = DISVILLE " & vbCrLf
    SQL = SQL & "       AND EST_ORI.ESTESTADO = CIU_ORI.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND CIU_DEST.VILCLEF = WCCL_VILLE " & vbCrLf
    SQL = SQL & "       AND EST_DEST.ESTESTADO = CIU_DEST.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND TDCDCLAVE(+) = WEL.WEL_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       AND TDCDSTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TRACLAVE(+) = WEL.WEL_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND TRASTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE(+) = WEL.WEL_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND WEL_ORI.WELCLAVE(+) = WEL.WEL_WELCLAVE " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE = TRACLAVE " & vbCrLf
    SQL = SQL & "UNION ALL " & vbCrLf
    SQL = SQL & " SELECT NVL(TDCD.TDCDFACTURA, WCD.WCDFACTURA) " & vbCrLf
    SQL = SQL & "       , WCD.WCD_FIRMA " & vbCrLf
    SQL = SQL & "       , TO_CHAR( WCD.DATE_CREATED, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAE_FECHA_RECOLECCION, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , TO_CHAR(TAE.TAEFECHALLEGADA, 'DD/MM/YYYY HH24:MI') " & vbCrLf
    SQL = SQL & "       , 'n/a' " & vbCrLf
    SQL = SQL & "       , WCD.WCD_PEDIDO_CLIENTE " & vbCrLf
    SQL = SQL & "       , WCD.WCD_CDAD_BULTOS " & vbCrLf
    SQL = SQL & "       , INITCAP(DIS.DISNOM) REMITENTE " & vbCrLf
    SQL = SQL & "       , InitCap(DISADRESSE1 || ' ' || ' ' || DISNUMEXT || '  ' || DISNUMINT || '  <br> ' ||DISADRESSE2 || DECODE(DISCODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DISCODEPOSTAL)) " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_ORI.VILNOM || ' ('|| EST_ORI.ESTNOMBRE || ')') " & vbCrLf
    SQL = SQL & "       , INITCAP(CCL.CCL_NOMBRE || ' ' || NVL(DIE.DIE_A_ATENCION_DE, DIE.DIENOMBRE)) " & vbCrLf
    SQL = SQL & "       , InitCap( DIEADRESSE1|| ' ' || ' ' || DIENUMEXT || '  ' || DIENUMINT || '  <br> ' ||DIEADRESSE2 || DECODE(DIECODEPOSTAL,NULL,NULL, ' <BR>C.P. ' || DIECODEPOSTAL)) " & vbCrLf
    SQL = SQL & "       , INITCAP(CIU_DEST.VILNOM || ' ('|| EST_DEST.ESTNOMBRE || ')') " & vbCrLf
    SQL = SQL & "       , WCD.WCDSTATUS " & vbCrLf
    SQL = SQL & "       , WCD.WCD_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       , 'Cross Dock' " & vbCrLf
    SQL = SQL & "       , WCD_CLICLEF " & vbCrLf
    SQL = SQL & "       , WCD.WCDOBSERVACION " & vbCrLf
    SQL = SQL & "       , WCD.WCDPESO " & vbCrLf
    SQL = SQL & "       , WCD.WCDVOLUMEN " & vbCrLf
    SQL = SQL & "       , WCD.WCDCLAVE AS NUI " & vbCrLf
    SQL = SQL & " FROM WCROSS_DOCK WCD " & vbCrLf
    SQL = SQL & "       , EDIRECCIONES_ENTREGA DIE " & vbCrLf
    SQL = SQL & "       , ECLIENT_CLIENTE CCL " & vbCrLf
    SQL = SQL & "       , EDISTRIBUTEUR DIS " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_ORI " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_ORI " & vbCrLf
    SQL = SQL & "       , ECIUDADES CIU_DEST " & vbCrLf
    SQL = SQL & "       , EESTADOS EST_DEST " & vbCrLf
    SQL = SQL & "       , ETRANS_DETALLE_CROSS_DOCK TDCD " & vbCrLf
    SQL = SQL & "       , ETRANSFERENCIA_TRADING TRA " & vbCrLf
    SQL = SQL & "       , ETRANS_ENTRADA TAE " & vbCrLf
    SQL = SQL & " WHERE WCD.WCD_FIRMA IN ('" & wTalonRastreo & "') " & vbCrLf
    SQL = SQL & "       AND DISCLEF = WCD.WCD_DISCLEF " & vbCrLf
    SQL = SQL & "       AND DIECLAVE = NVL(NVL(TDCD_DIECLAVE_ENT, TDCD_DIECLAVE), WCD_DIECLAVE_ENTREGA) " & vbCrLf
    SQL = SQL & "       AND CCLCLAVE = NVL(TDCD_CCLCLAVE, WCD.WCD_CCLCLAVE) " & vbCrLf
    SQL = SQL & "       AND CIU_ORI.VILCLEF = DISVILLE " & vbCrLf
    SQL = SQL & "       AND EST_ORI.ESTESTADO = CIU_ORI.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND CIU_DEST.VILCLEF = DIEVILLE " & vbCrLf
    SQL = SQL & "       AND EST_DEST.ESTESTADO = CIU_DEST.VIL_ESTESTADO " & vbCrLf
    SQL = SQL & "       AND TDCDCLAVE(+) = WCD.WCD_TDCDCLAVE " & vbCrLf
    SQL = SQL & "       AND TDCDSTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TRACLAVE(+) = WCD.WCD_TRACLAVE " & vbCrLf
    SQL = SQL & "       AND TRASTATUS (+) = '1' " & vbCrLf
    SQL = SQL & "       AND TAE_TRACLAVE = TRACLAVE " & vbCrLf

    rst4.Open SQL
    ReDim arrTmp(0, 0)
    If Not rst4.EOF Then
        i = 0
        c = rst4.RecordCount
        ReDim arrTmp(c, 22)
        
        While Not rst4.EOF
            If IsNull(rst4.Fields("TALON")) Then
                arrTmp(i, 0) = ""
            Else
                arrTmp(i, 0) = rst4.Fields("TALON")
            End If
            If IsNull(rst4.Fields("FIRMA")) Then
                arrTmp(i, 1) = ""
            Else
                arrTmp(i, 1) = rst4.Fields("FIRMA")
            End If
            If IsNull(rst4.Fields("FECHA_CREACION")) Then
                arrTmp(i, 2) = ""
            Else
                arrTmp(i, 2) = rst4.Fields("FECHA_CREACION")
            End If
            If IsNull(rst4.Fields("FECHA_RECOLECCION")) Then
                arrTmp(i, 3) = ""
            Else
                arrTmp(i, 3) = rst4.Fields("FECHA_RECOLECCION")
            End If
            If IsNull(rst4.Fields("FECHA_LLEGADA")) Then
                arrTmp(i, 4) = ""
            Else
                arrTmp(i, 4) = rst4.Fields("FECHA_LLEGADA")
            End If
            If IsNull(rst4.Fields("RECOL_DOMICILIO")) Then
                arrTmp(i, 5) = ""
            Else
                arrTmp(i, 5) = rst4.Fields("RECOL_DOMICILIO")
            End If
            If IsNull(rst4.Fields("FACTURA")) Then
                arrTmp(i, 6) = ""
            Else
                arrTmp(i, 6) = rst4.Fields("FACTURA")
            End If
            If IsNull(rst4.Fields("CDAD_BULTOS")) Then
                arrTmp(i, 7) = ""
            Else
                arrTmp(i, 7) = rst4.Fields("CDAD_BULTOS")
            End If
            If IsNull(rst4.Fields("REMITENTE")) Then
                arrTmp(i, 8) = ""
            Else
                arrTmp(i, 8) = rst4.Fields("REMITENTE")
            End If
            If IsNull(rst4.Fields("REMITENTE_DIREC")) Then
                arrTmp(i, 9) = ""
            Else
                arrTmp(i, 9) = rst4.Fields("REMITENTE_DIREC")
            End If
            If IsNull(rst4.Fields("REMITENTE_ESTADO")) Then
                arrTmp(i, 10) = ""
            Else
                arrTmp(i, 10) = rst4.Fields("REMITENTE_ESTADO")
            End If
            If IsNull(rst4.Fields("REMITENTE_1")) Then
                arrTmp(i, 11) = ""
            Else
                arrTmp(i, 11) = rst4.Fields("REMITENTE_1")
            End If
            If IsNull(rst4.Fields("REMITENTE_DIREC_1")) Then
                arrTmp(i, 12) = ""
            Else
                arrTmp(i, 12) = rst4.Fields("REMITENTE_DIREC_1")
            End If
            If IsNull(rst4.Fields("REMITENTE_ESTADO_1")) Then
                arrTmp(i, 13) = ""
            Else
                arrTmp(i, 13) = rst4.Fields("REMITENTE_ESTADO_1")
            End If
            If IsNull(rst4.Fields("ESTATUS")) Then
                arrTmp(i, 14) = ""
            Else
                arrTmp(i, 14) = rst4.Fields("ESTATUS")
            End If
            If IsNull(rst4.Fields("TDCDCLAVE")) Then
                arrTmp(i, 15) = ""
            Else
                arrTmp(i, 15) = rst4.Fields("TDCDCLAVE")
            End If
            If IsNull(rst4.Fields("TIPO")) Then
                arrTmp(i, 16) = ""
            Else
                arrTmp(i, 16) = rst4.Fields("TIPO")
            End If
            If IsNull(rst4.Fields("CLICLEF")) Then
                arrTmp(i, 17) = ""
            Else
                arrTmp(i, 17) = rst4.Fields("CLICLEF")
            End If
            If IsNull(rst4.Fields("OBSERVACIONES")) Then
                arrTmp(i, 18) = ""
            Else
                arrTmp(i, 18) = rst4.Fields("OBSERVACIONES")
            End If
            If IsNull(rst4.Fields("PESO")) Then
                arrTmp(i, 19) = ""
            Else
                arrTmp(i, 19) = rst4.Fields("PESO")
            End If
            If IsNull(rst4.Fields("VOLUMEN")) Then
                arrTmp(i, 20) = ""
            Else
                arrTmp(i, 20) = rst4.Fields("VOLUMEN")
            End If
            If IsNull(rst4.Fields("NUI")) Then
                arrTmp(i, 21) = ""
            Else
                arrTmp(i, 21) = rst4.Fields("NUI")
            End If
            i = i + 1
            rst4.MoveNext
        Wend
    End If
    
Catch:
    rst4.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerInfoTalon:" & Err.Description
    End If
    
    obtenerInfoTalon = arrTmp
End Function

'''Obtiene todo el tracking del talón:
Private Function obtenerTrackingTalon(tdCDclave As String)
	Dim SQL6 As String, indx As Integer, colu As Integer
	Dim rsT6 As New ADODB.Recordset
	Dim arrTmp6() As String

    Set rsT6 = Nothing
    Set rsT6 = New ADODB.Recordset
    rsT6.CursorLocation = adUseClient
    rsT6.CursorType = adOpenForwardOnly
    rsT6.LockType = adLockReadOnly
    rsT6.ActiveConnection = Db_link_orfeo
    
    indx = 0
    colu = 0
    SQL6 = ""

On Error GoTo Catch
    SQL6 = "SELECT       TO_CHAR(TAEFECHALLEGADA, 'DD/MM/YYYY HH24:MI') " & vbCrLf 'AS FECHA_LLEGADA " & vbCrLf
    SQL6 = SQL6 & "       , InitCap(CIU_EAL.VILNOM) || ' (' || InitCap(EST_EAL.ESTNOMBRE)  || ')' " & vbCrLf 'AS CIUDAD_ESTADO " & vbCrLf
    SQL6 = SQL6 & "       , 'Entrada CEDIS Logis (' || EAL_ORI.ALLCODIGO || ' - ' || InitCap(CIU_EAL.VILNOM) || ')' " & vbCrLf 'AS ENTRADA " & vbCrLf
    SQL6 = SQL6 & "       , DXP_TIPO_ENTREGA " & vbCrLf 'AS TIPO_ENTREGA " & vbCrLf
    SQL6 = SQL6 & "       , TO_CHAR(EXP_FECHA_SALIDA, 'DD/MM/YYYY HH24:MI')  " & vbCrLf 'AS FECHA_SALIDA " & vbCrLf
    SQL6 = SQL6 & "       , DECODE(DXP_TIPO_ENTREGA, 'DIRECTO', 'Expedicion directa al cliente', 'Expedicion de traslado al CEDIS Logis (' || EAL_DEST.ALLCODIGO || ' - ' || InitCap(CIU_DEST.VILNOM) || ')')  " & vbCrLf 'AS DESCRIPCION " & vbCrLf
    SQL6 = SQL6 & "       , TO_CHAR(DXP_FECHA_ENTREGA, 'DD/MM/YYYY HH24:MI') " & vbCrLf 'AS FECHA_ENTREGA " & vbCrLf
    SQL6 = SQL6 & "       , InitCap(DXP_TIPO_EVIDENCIA)  " & vbCrLf 'AS TIPO_EVIDENCIA " & vbCrLf
    SQL6 = SQL6 & "       , TRA_MEZTCLAVE_DEST   " & vbCrLf 'AS MEZTCLAVE_DEST " & vbCrLf
    SQL6 = SQL6 & "       , NVL(DXP_TINCLAVE, 0)  " & vbCrLf 'AS TINCLAVE " & vbCrLf
    SQL6 = SQL6 & "       , NVL(DXP_VAS, 'N') " & vbCrLf 'AS VAS " & vbCrLf
    SQL6 = SQL6 & "       , LOGIS.TIPO_OPERACION_FACT (TDCD.TDCDCLAVE, TDCD.TDCD_TRACLAVE) " & vbCrLf 'TIPO_OPERACION " & vbCrLf
    SQL6 = SQL6 & "FROM   ETRANS_DETALLE_CROSS_DOCK TDCD " & vbCrLf
    SQL6 = SQL6 & "       , ETRANSFERENCIA_TRADING TRA " & vbCrLf
    SQL6 = SQL6 & "       , EALMACENES_LOGIS EAL_ORI " & vbCrLf
    SQL6 = SQL6 & "       , ECIUDADES CIU_EAL " & vbCrLf
    SQL6 = SQL6 & "       , EESTADOS EST_EAL " & vbCrLf
    SQL6 = SQL6 & "       , ETRANS_ENTRADA TAE " & vbCrLf
    SQL6 = SQL6 & "       , EDET_EXPEDICIONES DXP " & vbCrLf
    SQL6 = SQL6 & "       , EEXPEDICIONES EXPED " & vbCrLf
    SQL6 = SQL6 & "       , EALMACENES_LOGIS EAL_DEST " & vbCrLf
    SQL6 = SQL6 & "       , ECIUDADES CIU_DEST " & vbCrLf
    SQL6 = SQL6 & "WHERE  TDCD.TDCDCLAVE IN ( " & vbCrLf
    SQL6 = SQL6 & "               SELECT  TO_CHAR(TDCDCLAVE) TDCDCLAVE " & vbCrLf
    SQL6 = SQL6 & "               FROM    ETRANS_DETALLE_CROSS_DOCK " & vbCrLf
    SQL6 = SQL6 & "               WHERE   TDCD_DXPCLAVE_ORI IN ( " & vbCrLf
    SQL6 = SQL6 & "                       SELECT  DXPCLAVE " & vbCrLf
    SQL6 = SQL6 & "                       FROM    EDET_EXPEDICIONES " & vbCrLf
    SQL6 = SQL6 & "                       WHERE   DXP_TIPO_ENTREGA IN ('TRASLADO', 'DIRECTO') " & vbCrLf
    SQL6 = SQL6 & "                       CONNECT BY PRIOR DXPCLAVE = DXP_DXPCLAVE " & vbCrLf
    SQL6 = SQL6 & "                       START   WITH DXP_TDCDCLAVE = '" & tdCDclave & "' " & vbCrLf
    SQL6 = SQL6 & "               ) " & vbCrLf
    SQL6 = SQL6 & "               UNION " & vbCrLf
    SQL6 = SQL6 & "               SELECT  '" & tdCDclave & "' TDCDCLAVE " & vbCrLf
    SQL6 = SQL6 & "               FROM    DUAL " & vbCrLf
    SQL6 = SQL6 & "       ) " & vbCrLf
    SQL6 = SQL6 & "       AND     TRACLAVE = TDCD.TDCD_TRACLAVE " & vbCrLf
    SQL6 = SQL6 & "       AND     TRASTATUS = '1' " & vbCrLf
    SQL6 = SQL6 & "       AND     TDCDSTATUS = '1' " & vbCrLf
    SQL6 = SQL6 & "       AND     EAL_ORI.ALLCLAVE = TRA_ALLCLAVE " & vbCrLf
    SQL6 = SQL6 & "       AND     CIU_EAL.VILCLEF = EAL_ORI.ALL_VILCLEF " & vbCrLf
    SQL6 = SQL6 & "       AND     EST_EAL.ESTESTADO = CIU_EAL.VIL_ESTESTADO " & vbCrLf
    SQL6 = SQL6 & "       AND     TAE_TRACLAVE = TRACLAVE " & vbCrLf
    SQL6 = SQL6 & "       AND     DXP_TDCDCLAVE(+) = TDCD.TDCDCLAVE " & vbCrLf
    SQL6 = SQL6 & "       AND     EXPCLAVE(+) = DXP_EXPCLAVE " & vbCrLf
    SQL6 = SQL6 & "       AND     EAL_DEST.ALLCLAVE(+) = DXP_ALLCLAVE_DEST " & vbCrLf
    SQL6 = SQL6 & "       AND     CIU_DEST.VILCLEF(+) = EAL_DEST.ALL_VILCLEF " & vbCrLf
    SQL6 = SQL6 & "ORDER  BY NVL(DXPCLAVE,0) DESC " & vbCrLf
    
    SQL6 = " SELECT 1,NULL,3,4,5,6,7,8,9,10,11,12 FROM DUAL "
    
    SQL6 = "SELECT       TO_CHAR(TAE.TAEFECHALLEGADA, 'DD/MM/YYYY HH24:MI')  FECHA_LLEGADA   " & vbCrLf
    SQL6 = SQL6 & "       , InitCap(CIU_EAL.VILNOM) || ' (' || InitCap(EST_EAL.ESTNOMBRE)  || ')'  CIUDAD_ESTADO   " & vbCrLf
    SQL6 = SQL6 & "       , 'Entrada CEDIS Logis (' || EAL_ORI.ALLCODIGO || ' - ' || InitCap(CIU_EAL.VILNOM) || ')'  ENTRADA   " & vbCrLf
    SQL6 = SQL6 & "       , DXP.DXP_TIPO_ENTREGA  TIPO_ENTREGA   " & vbCrLf
    SQL6 = SQL6 & "       , TO_CHAR(EXPED.EXP_FECHA_SALIDA, 'DD/MM/YYYY HH24:MI')   FECHA_SALIDA   " & vbCrLf
    SQL6 = SQL6 & "       , DECODE(DXP.DXP_TIPO_ENTREGA, 'DIRECTO', 'Expedicion directa al cliente', 'Expedicion de traslado al CEDIS Logis (' || EAL_DEST.ALLCODIGO || ' - ' || InitCap(CIU_DEST.VILNOM) || ')')   DESCRIPCION   " & vbCrLf
    SQL6 = SQL6 & "       , TO_CHAR(DXP.DXP_FECHA_ENTREGA, 'DD/MM/YYYY HH24:MI')  FECHA_ENTREGA   " & vbCrLf
    SQL6 = SQL6 & "       , InitCap(DXP.DXP_TIPO_EVIDENCIA)   TIPO_EVIDENCIA   " & vbCrLf
    SQL6 = SQL6 & "       , TRA.TRA_MEZTCLAVE_DEST    MEZTCLAVE_DEST   " & vbCrLf
    SQL6 = SQL6 & "       , NVL(DXP.DXP_TINCLAVE, 0)   TINCLAVE   " & vbCrLf
    SQL6 = SQL6 & "       , NVL(DXP.DXP_VAS, 'N')  VAS   " & vbCrLf
    SQL6 = SQL6 & "       , LOGIS.TIPO_OPERACION_FACT (TDCD.TDCDCLAVE, TDCD.TDCD_TRACLAVE) TIPO_OPERACION  " & vbCrLf
    SQL6 = SQL6 & "FROM   ETRANS_DETALLE_CROSS_DOCK TDCD  " & vbCrLf
    SQL6 = SQL6 & " INNER JOIN ETRANSFERENCIA_TRADING TRA  " & vbCrLf
    SQL6 = SQL6 & "     ON  TRA.TRACLAVE = TDCD.TDCD_TRACLAVE  " & vbCrLf
    SQL6 = SQL6 & " INNER JOIN EALMACENES_LOGIS EAL_ORI  " & vbCrLf
    SQL6 = SQL6 & "     ON  EAL_ORI.ALLCLAVE = TRA.TRA_ALLCLAVE  " & vbCrLf
    SQL6 = SQL6 & " INNER JOIN ECIUDADES CIU_EAL  " & vbCrLf
    SQL6 = SQL6 & "     ON  CIU_EAL.VILCLEF = EAL_ORI.ALL_VILCLEF  " & vbCrLf
    SQL6 = SQL6 & " INNER JOIN EESTADOS EST_EAL  " & vbCrLf
    SQL6 = SQL6 & "     ON  EST_EAL.ESTESTADO = CIU_EAL.VIL_ESTESTADO  " & vbCrLf
    SQL6 = SQL6 & " INNER JOIN ETRANS_ENTRADA TAE  " & vbCrLf
    SQL6 = SQL6 & "     ON  TAE.TAE_TRACLAVE = TRA.TRACLAVE  " & vbCrLf
    SQL6 = SQL6 & " LEFT JOIN EDET_EXPEDICIONES DXP  " & vbCrLf
    SQL6 = SQL6 & "     ON  TDCD.TDCDCLAVE = DXP.DXP_TDCDCLAVE  " & vbCrLf
    SQL6 = SQL6 & " LEFT JOIN EEXPEDICIONES EXPED  " & vbCrLf
    SQL6 = SQL6 & "     ON  DXP.DXP_EXPCLAVE = EXPED.EXPCLAVE  " & vbCrLf
    SQL6 = SQL6 & " LEFT JOIN EALMACENES_LOGIS EAL_DEST  " & vbCrLf
    SQL6 = SQL6 & "     ON  DXP.DXP_ALLCLAVE_DEST = EAL_DEST.ALLCLAVE  " & vbCrLf
    SQL6 = SQL6 & " LEFT JOIN ECIUDADES CIU_DEST  " & vbCrLf
    SQL6 = SQL6 & "     ON  EAL_DEST.ALL_VILCLEF = CIU_DEST.VILCLEF  " & vbCrLf
    SQL6 = SQL6 & "WHERE  TDCD.TDCDCLAVE IN (   " & vbCrLf
    SQL6 = SQL6 & "               SELECT  TO_CHAR(TDCD1.TDCDCLAVE) AS TDCDCLAVE   " & vbCrLf
    SQL6 = SQL6 & "               FROM    ETRANS_DETALLE_CROSS_DOCK TDCD1   " & vbCrLf
    SQL6 = SQL6 & "               WHERE   TDCD1.TDCD_DXPCLAVE_ORI IN (   " & vbCrLf
    SQL6 = SQL6 & "                       SELECT  EDX.DXPCLAVE   " & vbCrLf
    SQL6 = SQL6 & "                       FROM    EDET_EXPEDICIONES EDX  " & vbCrLf
    SQL6 = SQL6 & "                       WHERE   EDX.DXP_TIPO_ENTREGA IN ('TRASLADO', 'DIRECTO')   " & vbCrLf
    SQL6 = SQL6 & "                       CONNECT BY PRIOR EDX.DXPCLAVE = EDX.DXP_DXPCLAVE   " & vbCrLf
    SQL6 = SQL6 & "                       START   WITH EDX.DXP_TDCDCLAVE = '" & tdCDclave & "'   " & vbCrLf
    SQL6 = SQL6 & "               )   " & vbCrLf
    SQL6 = SQL6 & "               UNION   " & vbCrLf
    SQL6 = SQL6 & "               SELECT  '" & tdCDclave & "' AS TDCDCLAVE   " & vbCrLf
    SQL6 = SQL6 & "               FROM    DUAL   " & vbCrLf
    SQL6 = SQL6 & "       )   " & vbCrLf
    SQL6 = SQL6 & "       AND     TRA.TRASTATUS = '1'   " & vbCrLf
    SQL6 = SQL6 & "       AND     TDCD.TDCDSTATUS = '1'   " & vbCrLf
    SQL6 = SQL6 & "ORDER  BY NVL(DXP.DXPCLAVE,0) DESC " & vbCrLf
    
    rsT6.Open SQL6
    If Not rsT6.EOF Then
        ReDim arrTmp6(12, 1)
        
        indx = 0
        colu = 0
        tstResult = tstResult & "| inicia WHILE " & vbCrLf
        While colu < 12
            If IsNull(rsT6.Fields(colu)) Then
                arrTmp6(colu, indx) = ""
            Else
                arrTmp6(colu, indx) = CStr(rsT6.Fields(colu))
            End If
            colu = colu + 1
        Wend
        tstResult = tstResult & "| termina WHILE " & vbCrLf

        rsT6.Close
    Else
        tstResult = vbCrLf & tstResult & "|colu:" & colu & vbCrLf
    End If
        
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & vbCrLf & "  | arrTmp6(" & colu & "," & indx & ") " & vbCrLf & " SQL6: " & SQL6 & vbCrLf
        tstResult = tstResult & vbCrLf & " | obtenerTrackingTalon(catch): " & Err.Description & vbCrLf
    End If
    
    obtenerTrackingTalon = arrTmp6
End Function

'''Obtiene el estatus de la LTL que se encuentra registrado en BD's sin tomar en cuenta ninguna otra condición:
Private Function obtenerEstatusSinDocumentar(ByVal wTalonRastreo As String)
On Error GoTo Catch
    SQL = " SELECT WELSTATUS,WELFACTURA,WELCLAVE,WELCONS_GENERAL,WEL_TALON_RASTREO,WEL_FIRMA FROM WEB_LTL WHERE WEL_TALON_RASTREO = '" & wTalonRastreo & "'   " & vbCrLf
    
    rst2.Open SQL
    If Not rst2.EOF Then
        i = 0
        c = rst2.RecordCount
        ReDim arrTmp(c, 6)
        
        While Not rst2.EOF
            arrTmp(i, 0) = rst2.Fields("WELSTATUS")
            arrTmp(i, 1) = rst2.Fields("WELFACTURA")
            arrTmp(i, 2) = rst2.Fields("WELCLAVE")
            arrTmp(i, 3) = rst2.Fields("WELCONS_GENERAL")
            arrTmp(i, 4) = rst2.Fields("WEL_TALON_RASTREO")
            arrTmp(i, 5) = rst2.Fields("WEL_FIRMA")
            i = i + 1
            rst2.MoveNext
        Wend
    End If
    
Catch:
    rst2.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerEstatusSinDocumentar(catch): " & Err.Description
    End If
    
    obtenerEstatusSinDocumentar = arrTmp
End Function

Private Function ObtenerEstatusSinDocumentarCD(wTalonRastreo)
On Error GoTo Catch
   SQL = " SELECT " & vbCrLf
        SQL = SQL & "   WCDSTATUS " & vbCrLf
        SQL = SQL & " , WCDFACTURA " & vbCrLf
        SQL = SQL & " , WCDCLAVE " & vbCrLf
        SQL = SQL & " , WCD_CLICLEF" & vbCrLf
        SQL = SQL & " FROM WCROSS_DOCK WCD" & vbCrLf
        SQL = SQL & " WHERE WCD_FIRMA  = '" & wTalonRastreo & "'   " & vbCrLf
    
   rst3.Open SQL
    If Not rst3.EOF Then
        i = 0
        c = rst3.RecordCount
        ReDim arrTmp(c, 4)
        
        While Not rst3.EOF
            arrTmp(i, 0) = rst3.Fields("WCDSTATUS")
            arrTmp(i, 1) = rst3.Fields("WCDFACTURA")
            arrTmp(i, 2) = rst3.Fields("WCDCLAVE")
            arrTmp(i, 3) = rst3.Fields("WCD_CLICLEF")
           
            i = i + 1
            rst3.MoveNext
        Wend
    End If
    
Catch:
    rst3.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|ObtenerEstatusSinDocumentarCD(catch): " & Err.Description
    End If
    
    ObtenerEstatusSinDocumentarCD = arrTmp
End Function

'''Funcionalidad que indica si una operación / movimiento del talón es VAS, LTL ó StandBy:
Private Function obtenerTipoOperacion(wTalonRastreo As String)
	On Error GoTo Catch
    SQL = "SELECT LPAD(WELCONS_GENERAL, 7, '0') || '-' || WEL_CLICLEF FROM WEB_LTL WHERE WEL_TALON_RASTREO = '" & wTalonRastreo & "'"
    
    rsT5.Open SQL
    If rsT5.EOF = False Then
        While Not rsT5.EOF
            TDCDFACTURA = rsT5.Fields(0)
        Wend
    Else
        TDCDFACTURA = ""
    End If
        
    SQL = " SELECT     CROSS.TDCDFACTURA  AS  FACTURA_TALON   " & vbCrLf
    SQL = SQL & "       ,LOGIS.TIPO_OPERACION_FACT (CROSS.TDCDCLAVE, CROSS.TDCD_TRACLAVE)   AS  TIPO_OPERACION  " & vbCrLf
    SQL = SQL & "       ,CEDIS.ALLCODIGO ||'-'|| CEDIS.ALLNOMBRE    AS  CEDIS   " & vbCrLf
    SQL = SQL & " FROM   ETRANS_DETALLE_CROSS_DOCK CROSS    " & vbCrLf
    SQL = SQL & "       ,EALMACENES_LOGIS CEDIS " & vbCrLf
    SQL = SQL & " WHERE  CROSS.TDCDFACTURA      =   '" & TDCDFACTURA & "'   " & vbCrLf
    SQL = SQL & "   AND  CROSS.TDCD_ALLCLAVE    =   CEDIS.ALLCLAVE  " & vbCrLf
    SQL = SQL & " ORDER BY  CROSS.TDCDFECHA_BASE    " & vbCrLf
    
Catch:
    rsT5.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerTipoOperacion(catch): " & Err.Description
    End If
End Function

Private Function ordenarTracking(arrInfo)
    Dim arrTmp, x, y, z
	On Error GoTo Catch
    If IsArray(arrInfo) Then
        ReDim arrTmp((UBound(arrInfo, 2) + 2) * 10)
        y = 0
        For x = 0 To UBound(arrInfo, 2)
            For y = 0 To UBound(arrInfo)
                arrTmp(z) = arrInfo(y, x)
                z = z + 1
            Next
        Next
    Else
        arrTmp = arrInfo
    End If

Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|ordenarTracking(catch):" & Err.Description
    End If
    
    ordenarTracking = arrTmp
End Function

Private Function esArregloConElementos(arrInfo) As Boolean
    Dim numElementos As Double
    On Error GoTo Catch:
    
    If IsArray(arrInfo) = True Then
        numElementos = UBound(arrInfo)
        If numElementos = 0 Then
            esArregloConElementos = False
        Else
            esArregloConElementos = True
        End If
        Exit Function
    Else
        esArregloConElementos = False
    End If

Catch:
    If Err.Number <> 0 Then
        numElementos = -1
        esArregloConElementos = False
    End If
End Function

Function EsGuiaCancelada(sGuia_Firma) As Boolean
    Dim result As Boolean
    Dim arrInfo() As String
    result = False
    
    On Error GoTo Catch
    
    arrInfo = obtenerEstatusSinDocumentar(sGuia_Firma)
    
    If esArregloConElementos(arrInfo) Then
        If CStr(arrInfo(0, 0)) = "0" Then
            result = True
        End If
    Else
        arrInfo = ObtenerEstatusSinDocumentarCD(sGuia_Firma)
        
        If esArregloConElementos(arrInfo) Then
            If CStr(arrInfo(0, 0)) = "0" Then
                result = True
            End If
        End If
    End If
    
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|EsGuiaCancelada(catch):" & Err.Description
    End If
    
    EsGuiaCancelada = result
End Function

Function GuiaEnStndBy(sGuia_Firma As String) As Boolean
    Dim result As Boolean
    Dim arrInfo() As String
    result = False
    
    On Error GoTo Catch
    
    arrInfo = obtenerEstatusSinDocumentar(sGuia_Firma)
    
    If esArregloConElementos(arrInfo) Then
        If CStr(arrInfo(0, 0)) = "2" Then
            result = True
        End If
    Else
        arrInfo = ObtenerEstatusSinDocumentarCD(sGuia_Firma)
        
        If esArregloConElementos(arrInfo) Then
            If CStr(arrInfo(0, 0)) = "2" Then
                result = True
            End If
        End If
    End If
        
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|GuiaEnStndBy(catch):" & Err.Description
    End If

    GuiaEnStndBy = result
End Function

Function GuiaEntregada(sGuia_Firma As String) As Boolean
    Dim result_ent As Boolean
    Dim arrInfo_ent() As String
    Dim arrInfoTracking_ent() As String
    Dim iIncidencia_ent As Integer
    Dim sFecha_Entrega_ent As String, sTDcdClave_ent As String
    
    On Error GoTo Catch
    
    result_ent = False
    sTDcdClave_ent = ""
    iIncidencia_ent = -1
    sFecha_Entrega_ent = ""
    
    arrInfo_ent = obtenerInfoTalon(sGuia_Firma)
    
    If esArregloConElementos(arrInfo_ent) Then
        sTDcdClave_ent = arrInfo_ent(0, 15)
    End If
    
    If sTDcdClave_ent <> "" Then
        arrInfoTracking_ent = obtenerTrackingTalon(sTDcdClave_ent)
        
        If esArregloConElementos(arrInfoTracking_ent) Then
            iIncidencia_ent = arrInfoTracking_ent(9, 0)
            sFecha_Entrega_ent = arrInfoTracking_ent(6, 0)
            tstResult = tstResult & "|iIncidencia_ent:" & CStr(iIncidencia_ent)
            tstResult = tstResult & "|sFecha_Entrega_ent:" & CStr(sFecha_Entrega_ent)
        Else
            tstResult = tstResult & "|obtenerTrackingTalon: sin elementos"
        End If
        
        If iIncidencia_ent = "0" And sFecha_Entrega_ent <> "" Then
            tstResult = tstResult & "|result_ent:" & CStr(result_ent)
            result_ent = True
        End If
    End If
            
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|GuiaEntregada(catch):" & Err.Description
    End If
    
    GuiaEntregada = result_ent
End Function

' AVC
Function GetInfoAVC(pedimento As String, aduana As String, cliente As String, Optional anio As String = "") As String
    Dim servidor As String
    servidor = "\\192.168.0.103\"
    
    On Error GoTo Catch
    SQL = "SELECT  DISTINCT EDOCSAT.DSA_SGECLAVE DSA_SGECLAVE, REPLACE(EDOCSAT.DSACARPETA,'/','\') DSACARPETA, EDOCSAT.DSAARCHIVO DSAARCHIVO " & vbCrLf
    SQL = SQL & " FROM EDOCUMENTOS_SAT EDOCSAT " & vbCrLf
    SQL = SQL & "INNER JOIN ESAAI_M3_GENERAL SGE  ON SGE.SGECLAVE =  EDOCSAT.DSA_SGECLAVE " & vbCrLf
    SQL = SQL & "INNER JOIN TB_AVC_DET_PEDIMENTO AVC_PED ON AVC_PED.PEDNUMERO = SGE.SGEPEDNUMERO AND AVC_PED.PEDDOUANE = SGE.SGEDOUCLEF AND AVC_PED.PEDANIO = SGE.SGEANIO AND AVC_PED.CLIENTE = SGE.SGE_CLICLEF " & vbCrLf
    SQL = SQL & "INNER JOIN TB_AVC AVC ON  AVC.CODE =  AVC_PED.AVC_CODE " & vbCrLf
    SQL = SQL & "WHERE 1=1 " & vbCrLf
    SQL = SQL & "   AND UPPER(EDOCSAT.DSAARCHIVO) LIKE UPPER('%-AVC.pdf') " & vbCrLf
    SQL = SQL & "AND AVC.AVC_STATUS NOT IN ('X')" & vbCrLf
    SQL = SQL & "AND SGE.SGEPEDNUMERO =  '" & pedimento & "' " & vbCrLf
    SQL = SQL & "AND SGE.SGEDOUCLEF  = '" & aduana & "' " & vbCrLf
    
    If anio <> "" Then
        SQL = SQL & "AND SGE.SGEANIO  = '" & anio & "' " & vbCrLf
    End If
    
    SQL = SQL & "AND SGE.SGE_CLICLEF = '" & cliente & "' " & vbCrLf
    
    GetInfoAVC = ""
    
    Set rsAVC = Nothing
    Set rsAVC = New ADODB.Recordset
    rsAVC.CursorLocation = adUseClient
    rsAVC.CursorType = adOpenForwardOnly
    rsAVC.LockType = adLockReadOnly
    rsAVC.ActiveConnection = Db_link_orfeo
    
    rsAVC.Open SQL
    If Not rsAVC.EOF Then
        While Not rsAVC.EOF
            If GetInfoAVC = "" Then
                Debug.Print "Consultando archivo: " & servidor & rsAVC.Fields("DSACARPETA") & rsAVC.Fields("DSAARCHIVO")
                If FSO.FileExists(servidor & rsAVC.Fields("DSACARPETA") & rsAVC.Fields("DSAARCHIVO")) Then
                    GetInfoAVC = rsAVC.Fields("DSA_SGECLAVE") & "|" & servidor & rsAVC.Fields("DSACARPETA") & "|" & rsAVC.Fields("DSAARCHIVO")
                    '3094706|appwin\Expedientes_AVC\21180\|1734-AVC073008202200000152681-AVC.pdf
                            Else
                                    GetInfoAVC = ""
                End If
            End If
            rsAVC.MoveNext
        Wend
    Else
        GetInfoAVC = ""
    End If
    rsAVC.Close
    
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        GetInfoAVC = ""
    End If
End Function

'Función para concatenar los resultados del query en una sola cadena
'''NOTA: el query debe de retornar un solo campo
Function GetQueryString(ByVal sql_str As String)
    On Error GoTo Catch
    str_result = ""
    Dim varTest As String
    varTest = ""
    
    'Concatenar los resultados del query en una cadena
    Set rs_str = Nothing
    Set rs_str = New ADODB.Recordset
    rs_str.CursorLocation = adUseClient
    rs_str.CursorType = adOpenForwardOnly
    rs_str.LockType = adLockReadOnly
    rs_str.ActiveConnection = Db_link_orfeo
        
    varTest = Trim(sql_str)
    rs_str.Open varTest
    
    Do While Not rs_str.EOF
        If NVL(rs_str.Fields(0)) <> "" Then
            If str_result = "" Then
                str_result = rs_str.Fields(0)
            Else
                str_result = str_result & "," & rs_str.Fields(0)
            End If
        End If
        rs_str.MoveNext
    Loop
    rs_str.Close
Catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        str_result = ""
    End If
    GetQueryString = str_result
End Function

Function ValidaNombreArchivo(ByVal idRep As String)
    Dim res As Boolean
    res = True
    If idRep = "252" Or idRep = "253" Or idRep = "342" Or idRep = "343" Or idRep = "344" Then
        res = False
    End If
    ValidaNombreArchivo = res
End Function

Function cliente_con_seguro(CliClef)
        Dim res, sqlSeguro, arrSeguro, iCveEmpresa, iCCOClave
        Dim rs_str
        
        res = False
        sqlSeguro = ""
        iCCOClave = -1
        
        iCveEmpresa = obtener_clave_empresa(CliClef)
                
        sqlSeguro = sqlSeguro & " SELECT        CCOCLAVE, CCO_CLICLEF, CCO_BPCCLAVE, CCO_YFOCLEF, CCO_DOUCLEF, CCO_CHOCLAVE, CCO_PARCLAVE " & vbCrLf
        sqlSeguro = sqlSeguro & " FROM  ECLIENT_APLICA_CONCEPTOS " & vbCrLf
        sqlSeguro = sqlSeguro & " WHERE CCO_CLICLEF     =       '" & CliClef & "' /* CAMBIAR CLIENTE */ " & vbCrLf
        sqlSeguro = sqlSeguro & "       AND     CCO_CLICLEF     NOT IN (9954,9955,9956,9910,9929) " & vbCrLf
        sqlSeguro = sqlSeguro & "       AND     EXISTS  ( " & vbCrLf
        sqlSeguro = sqlSeguro & "                                       SELECT  NULL " & vbCrLf
        sqlSeguro = sqlSeguro & "                                       FROM    EBASES_POR_CONCEPT " & vbCrLf
        sqlSeguro = sqlSeguro & "                                       WHERE   BPCCLAVE        =       CCO_BPCCLAVE " & vbCrLf
        sqlSeguro = sqlSeguro & "                                               AND     BPC_CHOCLAVE    IN      (       SELECT  CHOCLAVE " & vbCrLf
        sqlSeguro = sqlSeguro & "                                                                                                       FROM    ECONCEPTOSHOJA " & vbCrLf
        sqlSeguro = sqlSeguro & "                                                                                                       WHERE   CHOTIPOIE               =       'I' " & vbCrLf
        sqlSeguro = sqlSeguro & "                                                                                                               AND     CHONUMERO               =       183 /* CONCEPTO SEGURO DE MERCANCÍA  / NO SE CAMBIA */ " & vbCrLf
        sqlSeguro = sqlSeguro & "                                                                                                               AND     CHO_EMPCLAVE    =       '" & iCveEmpresa & "' /* CAMBIAR EMPRESA */ " & vbCrLf
        sqlSeguro = sqlSeguro & "                                                                                               ) " & vbCrLf
        sqlSeguro = sqlSeguro & "                               ) " & vbCrLf
        
        Set rs_str = Nothing
        Set rs_str = New ADODB.Recordset
        rs_str.CursorLocation = adUseClient
        rs_str.CursorType = adOpenForwardOnly
        rs_str.LockType = adLockReadOnly
        rs_str.ActiveConnection = Db_link_orfeo

        rs_str.Open sqlSeguro
            Do While Not rs_str.EOF
                If NVL(rs_str.Fields(0)) <> "" Then
                    iCCOClave = CDbl(rs_str.Fields(0))
                End If
                rs_str.MoveNext
            Loop
        rs_str.Close
        
        If iCCOClave > 0 Then
                res = True
        End If
        
        cliente_con_seguro = res
End Function

Function obtener_clave_empresa(ByVal CliClef As String) As Double
    Dim sqlEmpresa As String, iCveEmpresa As Double
    Dim rs_str
    
    iCveEmpresa = -1
    sqlEmpresa = ""
	
    sqlEmpresa = sqlEmpresa & " SELECT  CET_EMPCLAVE CVE_EMPRESA " & vbCrLf
    sqlEmpresa = sqlEmpresa & " FROM    ECLIENT_EMPRESA_TRADING LIGA " & vbCrLf
    sqlEmpresa = sqlEmpresa & " WHERE   1=1 " & vbCrLf
    sqlEmpresa = sqlEmpresa & " AND LIGA.CET_CLICLEF = '" & CliClef & "' " & vbCrLf
    
	Set rs_str = Nothing
	Set rs_str = New ADODB.Recordset
	rs_str.CursorLocation = adUseClient
	rs_str.CursorType = adOpenForwardOnly
	rs_str.LockType = adLockReadOnly
	rs_str.ActiveConnection = Db_link_orfeo
    
    rs_str.Open sqlEmpresa
    If Not rs_str.EOF Then
        iCveEmpresa = CDbl(rs_str.Fields(0))
    End If
    rs_str.Close

    obtener_clave_empresa = iCveEmpresa
End Function

Function valida_valor_mercancia(val_mercancia As String)
    Dim res As Integer
    
	res = 0
	
	If val_mercancia = "" Then
		''la variable se encuentra vacia 
		res = -1
    ElseIf Not IsNumeric(val_mercancia) Then
        '' la variable  no es numerica 
		res = -2
    ElseIf CDbl(val_mercancia) <= 0 Then
        '' la variable contiene un número negativo o es igual a 0 
		res = -3
	Else 
		res = 1
    End If
    
	valida_valor_mercancia = res
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function es_captura_con_factura(ByVal num_client As String) As Boolean
	Dim res As Boolean
	Dim iConFactura, sqlConFactura, rs_capt

	res = False
	iConFactura = -1
	sqlConFactura = ""
	Set rs_capt = Nothing
	Set rs_capt = New ADODB.Recordset
	rs_capt.CursorLocation = adUseClient
	rs_capt.CursorType = adOpenForwardOnly
	rs_capt.LockType = adLockReadOnly
	rs_capt.ActiveConnection = Db_link_orfeo

	sqlConFactura = sqlConFactura & " SELECT	NVL(TIPO_DOCUMENTACION,-1) CON_FACTURA " & vbCrLf
	sqlConFactura = sqlConFactura & " FROM		TB_CONFIG_CLIENTE_DIST " & vbCrLf
	sqlConFactura = sqlConFactura & " WHERE		ID_CLIENTE	=	'" & num_client & "' " & vbCrLf

	rs_capt.Open sqlConFactura
    If Not rs_capt.EOF Then
        iConFactura = CDbl(rs_capt.Fields(0))
    End If
    rs_capt.Close

	If iConFactura = 1 Then
		res = True
	End If

	es_captura_con_factura = res
End Function
Function es_captura_sin_factura(ByVal num_client As String) As Boolean
	Dim res As Boolean
	Dim iSinFactura, sqlSinFactura, rs_capt

	res = False
	iSinFactura = -1
	sqlSinFactura = ""
	Set rs_capt = Nothing
	Set rs_capt = New ADODB.Recordset
	rs_capt.CursorLocation = adUseClient
	rs_capt.CursorType = adOpenForwardOnly
	rs_capt.LockType = adLockReadOnly
	rs_capt.ActiveConnection = Db_link_orfeo

	sqlSinFactura = sqlSinFactura & " SELECT	NVL(TIPO_DOCUMENTACION,-1) SIN_FACTURA " & vbCrLf
	sqlSinFactura = sqlSinFactura & " FROM		TB_CONFIG_CLIENTE_DIST " & vbCrLf
	sqlSinFactura = sqlSinFactura & " WHERE		ID_CLIENTE	=	'" & num_client & "' " & vbCrLf

	rs_capt.Open sqlSinFactura
    If Not rs_capt.EOF Then
        iSinFactura = CDbl(rs_capt.Fields(0))
    End If
    rs_capt.Close

	If iSinFactura = 0 Then
		res = True
	End If

	es_captura_sin_factura = res
End Function
Function es_captura_con_doc_fuente(ByVal num_client As String) As Boolean
	Dim res As Boolean
	Dim iConDocFuente, sqlConDocFuente, rs_capt
	
	res = False
	iConDocFuente = -1
	sqlConDocFuente = ""
	Set rs_capt = Nothing
	Set rs_capt = New ADODB.Recordset
	rs_capt.CursorLocation = adUseClient
	rs_capt.CursorType = adOpenForwardOnly
	rs_capt.LockType = adLockReadOnly
	rs_capt.ActiveConnection = Db_link_orfeo
	
	sqlConDocFuente = sqlConDocFuente & " SELECT	NVL(TIPO_DOCUMENTACION,-1) CON_DOCUMENTO_FUENTE " & vbCrLf
	sqlConDocFuente = sqlConDocFuente & " FROM		TB_CONFIG_CLIENTE_DIST " & vbCrLf
	sqlConDocFuente = sqlConDocFuente & " WHERE		ID_CLIENTE	=	'" & num_client & "' " & vbCrLf
	
	rs_capt.Open sqlConDocFuente
    If Not rs_capt.EOF Then
        iConDocFuente = CDbl(rs_capt.Fields(0))
    End If
    rs_capt.Close

	If iConDocFuente = 2 Then
		res = True
	End If
	
	es_captura_con_doc_fuente = res
End Function
Function validar_destinatario(ByVal clave_destino_cliente As String, ByVal num_client As String) As Double
	Dim res As Double
    Dim SQL As String
    Dim rs As New ADODB.Recordset

    res = -1
    SQL = ""
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo

	SQL = SQL & " SELECT	CCL.CCLCLAVE CCLCLAVE " & vbCrLf
	SQL = SQL & " FROM	EDIRECCION_ENTR_CLIENTE_LIGA DIL " & vbCrLf
	SQL = SQL & " 	INNER	JOIN	EDIRECCIONES_ENTREGA DIE	ON	DIE.DIECLAVE	=	DIL.DEC_DIECLAVE " & vbCrLf
	SQL = SQL & " 	INNER	JOIN	ECIUDADES CIU				ON	CIU.VILCLEF		=	DIE.DIEVILLE " & vbCrLf
	SQL = SQL & " 	INNER	JOIN	EESTADOS EST				ON	EST.ESTESTADO	=	CIU.VIL_ESTESTADO " & vbCrLf
	SQL = SQL & " 	INNER	JOIN	ECLIENT_CLIENTE CCL			ON	CCL.CCLCLAVE	=	DIE.DIE_CCLCLAVE " & vbCrLf
	SQL = SQL & " 	INNER	JOIN	EDESTINOS_POR_RUTA DER		ON	DER.DER_VILCLEF	=	CIU.VILCLEF " & vbCrLf
	SQL = SQL & " WHERE	1=1 " & vbCrLf
	SQL = SQL & " 	AND	CCL.CCL_STATUS		=	1 " & vbCrLf
	SQL = SQL & " 	AND	DIE.DIE_STATUS		=	1 " & vbCrLf
	SQL = SQL & " 	AND	EST.EST_PAYCLEF		=	'N3' " & vbCrLf
	SQL = SQL & " 	AND	DER.DER_ALLCLAVE	>	0 " & vbCrLf
	SQL = SQL & " 	AND	SF_LOGIS_CLIENTE_RESTRIC(DIL.DEC_CLICLEF, DER.DER_TIPO_ENTREGA)	=	1 " & vbCrLf
	SQL = SQL & " 	AND	DIL.DEC_NUM_DIR_CLIENTE	=	'" & clave_destino_cliente & "' " & vbCrLf
	SQL = SQL & " 	AND	DIL.DEC_CLICLEF			=	'" & num_client & "' " & vbCrLf

    On Error GoTo catch
		rs.Open SQL
		If Not rs.EOF Then
			res = CDbl(rs.Fields("CCLCLAVE"))
		End If
catch:
    rs.Close
    Set rs = Nothing

    validar_destinatario = res
End Function
Function tipo_tarifa_cliente(ByVal num_client As String) As String
	On Error GoTo catch
	
	'PENDIENTE DE QUE EL EQUIPO ORACLE FORMS NOS COMAPRTA EL QUERY PARA VALIDAR EL TIPO DE TARIFA
catch:
	tipo_tarifa_cliente = ""
End Function
Function validar_bultos_totales(ByVal bultos_totales As Double, ByVal bultos_granel As Double, ByVal cantidad_tarimas As Double, ByVal bultos_constitutivos_tarimas As Double) As Boolean
	Dim res As Boolean
	On Error GoTo catch
	
	res = False
	
	If bultos_totales = (bultos_granel + cantidad_tarimas) Then
		res = True
	End If

catch:
	validar_bultos_totales = res
End Function
Function validar_cdad_bultos_granel(ByVal bultos_granel As String, ByVal num_client As String) As Boolean
	Dim res As Boolean
	Dim cdad_bultos As Double
	Dim es_cliente_por_bulto As Boolean
	On Error GoTo catch
	
	res = False
	cdad_bultos = -1
	
	If tipo_tarifa_cliente(num_client) = "Bulto Constitutivo" Then
		If bultos_granel <> "" Then
			cdad_bultos = CDbl(bultos_granel)
			
			If cdad_bultos >= 0 Then
				res = True
			End If
		End If
	Else
		If bultos_granel = "" Then
			res = True
		End If
	End If
	
catch:
	validar_cdad_bultos_granel = res
End Function
Function validar_cdad_tarimas(ByVal cdad_tarimas As String) As Boolean
	Dim res As Boolean
	Dim cdad_bultos As Double
	On Error GoTo catch
	
	res = False
	cdad_bultos = -1
	
	If cdad_tarimas <> "" Then
		cdad_bultos = CDbl(cdad_tarimas)
		
		If cdad_bultos >= 0 Then
			res = True
		End If
	End If
	
catch:
	validar_cdad_tarimas = res
End Function
Function validar_bultos_por_tarima(ByVal cdad_bultos_tarima As String) As Boolean
	Dim res As Boolean
	Dim cdad_bultos As Double
	On Error GoTo catch
	
	res = False
	cdad_bultos = -1
	
	If cdad_bultos_tarima <> "" Then
		cdad_bultos = CDbl(cdad_bultos_tarima)
		
		If cdad_bultos >= 0 Then
			res = True
		End If
	Else
		res = True
	End If
	
catch:
	validar_bultos_por_tarima = res
End Function
Function validar_valor_mercancia(ByVal valor_mercancia As String, ByVal num_client As String) As Boolean
	Dim res As Boolean
	Dim val_merc As Double
	On Error GoTo catch
	
	res = False
	val_merc = -1
	
	If cliente_con_seguro(num_client) = True Then
		val_merc = CDbl(valor_mercancia)
		
		If val_merc >= 0 Then
			res = True
		End If
	Else
		res = True
	End If
	
catch:
	validar_valor_mercancia = res
End Function
Function validar_observaciones(ByVal txt_observaciones As String) As Boolean
	Dim res As Boolean
	Dim length As Double
	On Error GoTo catch
	
	res = False
	length = 0
	
	If txt_observaciones <> "" Then
		length = len(txt_observaciones)
	End If
	
	If length <= 80 Then
		res = True
	End If
	
catch:
	validar_observaciones = res
End Function
