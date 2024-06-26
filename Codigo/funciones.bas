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


'Variables de Correo
Private jmail As New jmail.Message


'   '   '   '   '   '   '
'   FUNCI�N PRINCIPAL   '
'   '   '   '   '   '   '

'Funci�n principal para obtener el estatus de la gu�a
''' Obtiene el estatus a partir del seguimiento que se le da a un tal�n/factura LTL y CD'''
Public Function ObtieneStatusTalon_txt(wTalonRastreo As String) As String
    On Error GoTo error_function

    'Inicializaci�n de variables:
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
                    'Se obtienen los par�metros: incidencia, fecha_entrega y last_entrada.'
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
                Debug.Print (Now & " No se puede obtener el tracking completo del tal�n debido a que no cuenta con tdCDclave. " & tdCDclave)
            End If


            '==========================='
            'Interpretaci�n del estatus.'
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

        'Si la gu�a tiene incidencia, se mantiene el estatus de la incidencia, si tiene fecha de entrega tambi�n se mantiene, de lo contrario se procesa el nuevo estatus:
        If incidencia <> "3" And incidencia <> "4" And CStr(fecha_entrega) = "" Then
            ''' =========================================== '''
            '''  Nuevo proceso para interpretar el estatus  '''
            ''' =========================================== '''
            ' 1.- Obtener informaci�n del tal�n;
            ' 2.- Replicar el proceso de la pantalla Tracking;
            ' 3.- Ajustar los estatus de acuerdo a las reglas que est�n en el excel (8 eventos);
            ' 4.- Aplicar las reglas de los colores que se van a mostrar en la pantalla;
            ' NOTA: todo se debe basar en el texto que est� en los registros de seguimiento que se encuentra en la BD's.

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
        '   Obtener informaci�n del cat�logo de acuerdo al estatus obtenido.    '
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
    '   PRESENTACI�N DE RESULTADOS  '
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

    Debug.Print (Now & " El tal�n " & wTalonRastreo & " tiene el estatus *" & stTexto & "*.")
    ObtieneStatusTalon_txt = stTexto
End Function

'   '   '   '   '   '   '
'   FUNCIONES INTERNAS  '
'   '   '   '   '   '   '
Private Sub init_var()
On Error GoTo catch

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
    
        
catch:
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

''' Retorna el cat�logo de estatus que se estar� manejando a nivel global para �ste tema, los casos particulares se contemplan en la funci�n principal.
Private Function obtenerCatalogoEstatus() As String()
On Error GoTo catch
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
            
catch:
    rst1.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerCatalogoEstatus:" & Err.Description
    End If

    obtenerCatalogoEstatus = arrTmp
End Function

'''Obtiene de BD's la informaci�n general del tal�n:
Private Function obtenerInfoTalon(wTalonRastreo As String) As String()
On Error GoTo catch
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
    
catch:
    rst4.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerInfoTalon:" & Err.Description
    End If
    
    obtenerInfoTalon = arrTmp
End Function

'''Obtiene todo el tracking del tal�n:
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

On Error GoTo catch
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
        
catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & vbCrLf & "  | arrTmp6(" & colu & "," & indx & ") " & vbCrLf & " SQL6: " & SQL6 & vbCrLf
        tstResult = tstResult & vbCrLf & " | obtenerTrackingTalon(catch): " & Err.Description & vbCrLf
    End If
    
    obtenerTrackingTalon = arrTmp6
End Function

'''Obtiene el estatus de la LTL que se encuentra registrado en BD's sin tomar en cuenta ninguna otra condici�n:
Private Function obtenerEstatusSinDocumentar(ByVal wTalonRastreo As String)
On Error GoTo catch
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
    
catch:
    rst2.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerEstatusSinDocumentar(catch): " & Err.Description
    End If
    
    obtenerEstatusSinDocumentar = arrTmp
End Function

Private Function ObtenerEstatusSinDocumentarCD(wTalonRastreo)
On Error GoTo catch
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
    
catch:
    rst3.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|ObtenerEstatusSinDocumentarCD(catch): " & Err.Description
    End If
    
    ObtenerEstatusSinDocumentarCD = arrTmp
End Function

'''Funcionalidad que indica si una operaci�n / movimiento del tal�n es VAS, LTL � StandBy:
Private Function obtenerTipoOperacion(wTalonRastreo As String)
        On Error GoTo catch
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
    
catch:
    rsT5.Close
    
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|obtenerTipoOperacion(catch): " & Err.Description
    End If
End Function

Private Function ordenarTracking(arrInfo)
    Dim arrTmp, x, y, z
        On Error GoTo catch
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

catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        tstResult = tstResult & "|ordenarTracking(catch):" & Err.Description
    End If
    
    ordenarTracking = arrTmp
End Function

Private Function esArregloConElementos(arrInfo) As Boolean
    Dim numElementos As Double
    On Error GoTo catch:
    
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

catch:
    If Err.Number <> 0 Then
        numElementos = -1
        esArregloConElementos = False
    End If
End Function

Function EsGuiaCancelada(sGuia_Firma) As Boolean
    Dim result As Boolean
    Dim arrInfo() As String
    result = False
    
    On Error GoTo catch
    
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
    
catch:
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
    
    On Error GoTo catch
    
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
        
catch:
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
    
    On Error GoTo catch
    
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
            
catch:
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
    
    On Error GoTo catch
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
    
catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        GetInfoAVC = ""
    End If
End Function

'Funci�n para concatenar los resultados del query en una sola cadena
'''NOTA: el query debe de retornar un solo campo
Function GetQueryString(ByVal sql_str As String)
    On Error GoTo catch
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
catch:
    If Err.Number <> 0 Then
        Debug.Print (Now & " - Ocurrio un error: " & Err.Description)
        str_result = ""
    End If
    GetQueryString = str_result
End Function

Function ValidaNombreArchivo(ByVal idRep As String)
    Dim Res As Boolean
    Res = True
    If idRep = "252" Or idRep = "253" Or idRep = "342" Or idRep = "343" Or idRep = "344" Then
        Res = False
    End If
    ValidaNombreArchivo = Res
End Function

Function cliente_con_seguro(CliClef)
        Dim Res, sqlSeguro, arrSeguro, iCveEmpresa, iCCOClave
        Dim rs_str
        
        Res = False
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
        sqlSeguro = sqlSeguro & "                                                                                                               AND     CHONUMERO               =       183 /* CONCEPTO SEGURO DE MERCANC�A  / NO SE CAMBIA */ " & vbCrLf
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
                Res = True
        End If
        
        cliente_con_seguro = Res
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
    Dim Res As Integer
    
        Res = 0
        
        If val_mercancia = "" Then
                ''la variable se encuentra vacia
                Res = -1
    ElseIf Not IsNumeric(val_mercancia) Then
        '' la variable  no es numerica
                Res = -2
    ElseIf CDbl(val_mercancia) <= 0 Then
        '' la variable contiene un n�mero negativo o es igual a 0
                Res = -3
        Else
                Res = 1
    End If
    
        valida_valor_mercancia = Res
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function obtener_nui_disponible(ByVal num_client As String) As Double
        Dim Res As Double
        Dim SQLnui, rs_Nui
        On Error GoTo catch
        
        Res = -1
        SQLnui = ""
        Set rs_Nui = Nothing
        Set rs_Nui = New ADODB.Recordset
        rs_Nui.CursorLocation = adUseClient
        rs_Nui.CursorType = adOpenForwardOnly
        rs_Nui.LockType = adLockReadOnly
        rs_Nui.ActiveConnection = Db_link_orfeo
        
        SQLnui = SQLnui & " SELECT      MIN(WEL.WELCLAVE) NUI " & vbCrLf
        SQLnui = SQLnui & " FROM        WEB_LTL WEL " & vbCrLf
        SQLnui = SQLnui & " WHERE       WEL.WEL_CLICLEF =       '" & num_client & "' " & vbCrLf
        SQLnui = SQLnui & " AND         WEL.WELSTATUS = 3" & vbCrLf
        
        rs_Nui.Open SQLnui
    If Not rs_Nui.EOF Then
        Res = CDbl(rs_Nui.Fields(0))
    End If
catch:
        rs_Nui.Close
        obtener_nui_disponible = Res
End Function
Function es_captura_con_factura(ByVal num_client As String) As Boolean
        Dim Res As Boolean
        Dim iConFactura, sqlConFactura, rs_con_f
        On Error GoTo catch

        Res = False
        iConFactura = -1
        sqlConFactura = ""
        Set rs_con_f = Nothing
        Set rs_con_f = New ADODB.Recordset
        rs_con_f.CursorLocation = adUseClient
        rs_con_f.CursorType = adOpenForwardOnly
        rs_con_f.LockType = adLockReadOnly
        rs_con_f.ActiveConnection = Db_link_orfeo

        sqlConFactura = sqlConFactura & " SELECT        NVL(TIPO_DOCUMENTACION,-1) CON_FACTURA " & vbCrLf
        sqlConFactura = sqlConFactura & " FROM          TB_CONFIG_CLIENTE_DIST " & vbCrLf
        sqlConFactura = sqlConFactura & " WHERE         ID_CLIENTE      =       '" & num_client & "' " & vbCrLf

        rs_con_f.Open sqlConFactura
    If Not rs_con_f.EOF Then
        iConFactura = CDbl(rs_con_f.Fields(0))
    End If
catch:
    rs_con_f.Close

        If iConFactura = 1 Then
                Res = True
        End If

        es_captura_con_factura = Res
End Function
Function es_captura_sin_factura(ByVal num_client As String) As Boolean
        Dim Res As Boolean
        Dim iSinFactura, sqlSinFactura, rs_sin_f
        On Error GoTo catch

        Res = False
        iSinFactura = -1
        sqlSinFactura = ""
        Set rs_sin_f = Nothing
        Set rs_sin_f = New ADODB.Recordset
        rs_sin_f.CursorLocation = adUseClient
        rs_sin_f.CursorType = adOpenForwardOnly
        rs_sin_f.LockType = adLockReadOnly
        rs_sin_f.ActiveConnection = Db_link_orfeo

        sqlSinFactura = sqlSinFactura & " SELECT        NVL(TIPO_DOCUMENTACION,-1) SIN_FACTURA " & vbCrLf
        sqlSinFactura = sqlSinFactura & " FROM          TB_CONFIG_CLIENTE_DIST " & vbCrLf
        sqlSinFactura = sqlSinFactura & " WHERE         ID_CLIENTE      =       '" & num_client & "' " & vbCrLf

        rs_sin_f.Open sqlSinFactura
    If Not rs_sin_f.EOF Then
        iSinFactura = CDbl(rs_sin_f.Fields(0))
    End If
catch:
    rs_sin_f.Close

        If iSinFactura = 0 Then
                Res = True
        End If

        es_captura_sin_factura = Res
End Function
Function es_captura_con_doc_fuente(ByVal num_client As String) As Boolean
        Dim Res As Boolean
        Dim iConDocFuente, sqlConDocFuente, rs_con_df
        On Error GoTo catch
        
        Res = False
        iConDocFuente = -1
        sqlConDocFuente = ""
        Set rs_con_df = Nothing
        Set rs_con_df = New ADODB.Recordset
        rs_con_df.CursorLocation = adUseClient
        rs_con_df.CursorType = adOpenForwardOnly
        rs_con_df.LockType = adLockReadOnly
        rs_con_df.ActiveConnection = Db_link_orfeo
        
        sqlConDocFuente = sqlConDocFuente & " SELECT    NVL(TIPO_DOCUMENTACION,-1) CON_DOCUMENTO_FUENTE " & vbCrLf
        sqlConDocFuente = sqlConDocFuente & " FROM              TB_CONFIG_CLIENTE_DIST " & vbCrLf
        sqlConDocFuente = sqlConDocFuente & " WHERE             ID_CLIENTE      =       '" & num_client & "' " & vbCrLf
        
        rs_con_df.Open sqlConDocFuente
    If Not rs_con_df.EOF Then
        iConDocFuente = CDbl(rs_con_df.Fields(0))
    End If
catch:
    rs_con_df.Close

        If iConDocFuente = 2 Then
                Res = True
        End If
        
        es_captura_con_doc_fuente = Res
End Function
Function validar_destinatario(ByVal clave_destino_cliente As String, ByVal num_client As String) As Double
        Dim Res As Double
    Dim SQL As String
    Dim rs_dest As New ADODB.Recordset
        On Error GoTo catch

    Res = -1
    SQL = ""
    Set rs_dest = New ADODB.Recordset
    rs_dest.CursorLocation = adUseClient
    rs_dest.CursorType = adOpenForwardOnly
    rs_dest.LockType = adLockBatchOptimistic
    rs_dest.ActiveConnection = Db_link_orfeo

        SQL = SQL & " SELECT    CCL.CCLCLAVE CCLCLAVE " & vbCrLf
        SQL = SQL & " FROM      EDIRECCION_ENTR_CLIENTE_LIGA DIL " & vbCrLf
        SQL = SQL & "   INNER   JOIN    EDIRECCIONES_ENTREGA DIE        ON      DIE.DIECLAVE    =       DIL.DEC_DIECLAVE " & vbCrLf
        SQL = SQL & "   INNER   JOIN    ECIUDADES CIU                           ON      CIU.VILCLEF             =       DIE.DIEVILLE " & vbCrLf
        SQL = SQL & "   INNER   JOIN    EESTADOS EST                            ON      EST.ESTESTADO   =       CIU.VIL_ESTESTADO " & vbCrLf
        SQL = SQL & "   INNER   JOIN    ECLIENT_CLIENTE CCL                     ON      CCL.CCLCLAVE    =       DIE.DIE_CCLCLAVE " & vbCrLf
        SQL = SQL & "   INNER   JOIN    EDESTINOS_POR_RUTA DER          ON      DER.DER_VILCLEF =       CIU.VILCLEF " & vbCrLf
        SQL = SQL & " WHERE     1=1 " & vbCrLf
        SQL = SQL & "   AND     CCL.CCL_STATUS          =       1 " & vbCrLf
        SQL = SQL & "   AND     DIE.DIE_STATUS          =       1 " & vbCrLf
        SQL = SQL & "   AND     EST.EST_PAYCLEF         =       'N3' " & vbCrLf
        SQL = SQL & "   AND     DER.DER_ALLCLAVE        >       0 " & vbCrLf
        SQL = SQL & "   AND     SF_LOGIS_CLIENTE_RESTRIC(DIL.DEC_CLICLEF, DER.DER_TIPO_ENTREGA) =       1 " & vbCrLf
        SQL = SQL & "   AND     DIL.DEC_NUM_DIR_CLIENTE =       '" & clave_destino_cliente & "' " & vbCrLf
        SQL = SQL & "   AND     DIL.DEC_CLICLEF                 =       '" & num_client & "' " & vbCrLf

        rs_dest.Open SQL
        If Not rs_dest.EOF Then
                Res = CDbl(rs_dest.Fields("CCLCLAVE"))
        End If
catch:
    rs_dest.Close
    Set rs_dest = Nothing

    validar_destinatario = Res
End Function
Function obtener_tipo_tarifa_cliente(ByVal num_client As String) As Double
        ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
        '       1       Peso/Volumen            '
        ' - - - - - - - - - - - - - '
        '       2       Caja/Tarima                     '
        ' - - - - - - - - - - - - - '
        '       3       Bulto Constitutivo      '
        ' - - - - - - - - - - - - - '
        '       4       Solo Tarima                     '
        ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
        
        Dim Res As Double
    Dim SQL As String
    Dim rs_tar As New ADODB.Recordset
        On Error GoTo catch

    Res = -1
    SQL = ""
    Set rs_tar = New ADODB.Recordset
    rs_tar.CursorLocation = adUseClient
    rs_tar.CursorType = adOpenForwardOnly
    rs_tar.LockType = adLockBatchOptimistic
    rs_tar.ActiveConnection = Db_link_orfeo
        
        SQL = SQL & " SELECT     DISTINCT " & vbCrLf
        SQL = SQL & "            CCD.ID_TIPO_TARIFA ID_TIPO_TARIFA " & vbCrLf
        SQL = SQL & "           ,TF.NOMBRE NOMBRE " & vbCrLf
        SQL = SQL & " FROM       LOGIS.TB_CONFIG_CLIENTE_DIST CCD " & vbCrLf
        SQL = SQL & "   INNER JOIN      LOGIS.TB_TIPO_TARIFA_DIST TF " & vbCrLf
        SQL = SQL & "           ON      CCD.ID_TIPO_TARIFA      =       TF.ID_TIPO_TARIFA " & vbCrLf
        SQL = SQL & " WHERE     CCD.ID_CLIENTE  =       '" & num_client & "' " & vbCrLf

        rs_tar.Open SQL
        If Not rs_tar.EOF Then
                Res = CDbl(rs_tar.Fields("ID_TIPO_TARIFA"))
        End If
catch:
        rs_tar.Close
    Set rs_tar = Nothing
        
        obtener_tipo_tarifa_cliente = Res
End Function
Function validar_bultos_totales(ByVal bultos_totales As Double, ByVal bultos_granel As Double, ByVal cantidad_tarimas As Double, ByVal bultos_constitutivos_tarimas As Double) As Boolean
        Dim Res As Boolean
        On Error GoTo catch
        
        Res = False
        
        If bultos_totales = (bultos_granel + cantidad_tarimas) Then
                Res = True
        End If

catch:
        validar_bultos_totales = Res
End Function
Function validar_cdad_bultos_granel(ByVal bultos_granel As String, ByVal num_client As String) As Boolean
        Dim Res As Boolean
        Dim cdad_bultos As Double
        Dim es_cliente_por_bulto As Boolean
        On Error GoTo catch
        
        Res = False
        cdad_bultos = -1
        
        If obtener_tipo_tarifa_cliente(num_client) = 3 Then
                If bultos_granel <> "" Then
                        cdad_bultos = CDbl(bultos_granel)
                        
                        If cdad_bultos >= 0 Then
                                Res = True
                        End If
                End If
        Else
                If bultos_granel = "" Then
                        Res = True
                End If
        End If
        
catch:
        validar_cdad_bultos_granel = Res
End Function
Function validar_cdad_tarimas(ByVal cdad_tarimas As String) As Boolean
        Dim Res As Boolean
        Dim cdad_bultos As Double
        On Error GoTo catch
        
        Res = False
        cdad_bultos = -1
        
        If cdad_tarimas <> "" Then
                cdad_bultos = CDbl(cdad_tarimas)
                
                If cdad_bultos >= 0 Then
                        Res = True
                End If
        End If
        
catch:
        validar_cdad_tarimas = Res
End Function
Function validar_bultos_por_tarima(ByVal cdad_bultos_tarima As String) As Boolean
        Dim Res As Boolean
        Dim cdad_bultos As Double
        On Error GoTo catch
        
        Res = False
        cdad_bultos = -1
        
        If cdad_bultos_tarima <> "" Then
                cdad_bultos = CDbl(cdad_bultos_tarima)
                
                If cdad_bultos >= 0 Then
                        Res = True
                End If
        Else
                Res = True
        End If
        
catch:
        validar_bultos_por_tarima = Res
End Function
Function validar_valor_mercancia(ByVal valor_mercancia As String, ByVal num_client As String) As Boolean
        Dim Res As Boolean
        Dim val_merc As Double
        On Error GoTo catch
        
        Res = False
        val_merc = -1
        
        If cliente_con_seguro(num_client) = True Then
                val_merc = CDbl(valor_mercancia)
                
                If val_merc >= 0 Then
                        Res = True
                End If
        Else
                Res = True
        End If
        
catch:
        validar_valor_mercancia = Res
End Function
Function validar_observaciones(ByVal txt_observaciones As String) As Boolean
        Dim Res As Boolean
        Dim length As Double
        On Error GoTo catch
        
        Res = False
        length = 0
        
        If txt_observaciones <> "" Then
                length = Len(txt_observaciones)
        End If
        
        If length <= 80 Then
                Res = True
        End If
        
catch:
        validar_observaciones = Res
End Function
Function obtener_condiciones_entrega(ByVal txt As String) As String
        Dim Res As String
        On Error GoTo catch
        
        Res = "N"
        
        If UCase(txt) = UCase("Entrega a domicilio") Or UCase(txt) = UCase("entrega_domicilio") Then
                Res = "S"
        End If
        
catch:
        obtener_condiciones_entrega = Res
End Function
Function obtener_dice_contener(ByVal num_client As String) As String
        Dim Res As String
    Dim SQL As String
    Dim rs_cont As New ADODB.Recordset
        On Error GoTo catch

    Res = ""
    SQL = ""
    Set rs_cont = New ADODB.Recordset
    rs_cont.CursorLocation = adUseClient
    rs_cont.CursorType = adOpenForwardOnly
    rs_cont.LockType = adLockBatchOptimistic
    rs_cont.ActiveConnection = Db_link_orfeo
        
        SQL = SQL & " SELECT    DESC_GEN_DISTRIBUCION DICE_CONTENER " & vbCrLf
        SQL = SQL & " FROM      TB_ECLIENT_CP " & vbCrLf
        SQL = SQL & " WHERE     CLICLEF =       '" & num_client & "' " & vbCrLf

        rs_cont.Open SQL
        If Not rs_cont.EOF Then
                Res = rs_cont.Fields("DICE_CONTENER")
        End If
        
catch:
        rs_cont.Close
    Set rs_cont = Nothing

        obtener_dice_contener = Res
End Function
Function obtener_recol_domicilio(ByVal num_client As String) As String
        Dim Res As String
    Dim SQL As String
    Dim rs_reco As New ADODB.Recordset
        On Error GoTo catch

    Res = "N"
        SQL = ""
    Set rs_reco = New ADODB.Recordset
    rs_reco.CursorLocation = adUseClient
    rs_reco.CursorType = adOpenForwardOnly
    rs_reco.LockType = adLockBatchOptimistic
    rs_reco.ActiveConnection = Db_link_orfeo
        
        SQL = SQL & " SELECT    NVL(RECOLECCION_A_DOMICILIO,0) RECO " & vbCrLf
        SQL = SQL & " FROM      LOGIS.TB_CONFIG_CLIENTE_DIST CCD " & vbCrLf
        SQL = SQL & "   INNER JOIN      LOGIS.TB_TIPO_TARIFA_DIST TF " & vbCrLf
        SQL = SQL & "           ON      CCD.ID_TIPO_TARIFA      =       TF.ID_TIPO_TARIFA " & vbCrLf
        SQL = SQL & " WHERE     CCD.ID_CLIENTE  =       '" & num_client & "' " & vbCrLf

        rs_reco.Open SQL
        If Not rs_reco.EOF Then
                If rs_reco.Fields("RECO") = "1" Then
                        Res = "S"
                End If
        End If
        
catch:
        rs_reco.Close
    Set rs_reco = Nothing

        obtener_recol_domicilio = Res
End Function
Function obtener_prepagado_por_cobrar(ByVal num_client As String) As String
        Dim Res As String
    Dim SQL As String
    Dim rs_prep As New ADODB.Recordset
        On Error GoTo catch

    Res = "PREPAGADO"
        SQL = ""
    Set rs_prep = New ADODB.Recordset
    rs_prep.CursorLocation = adUseClient
    rs_prep.CursorType = adOpenForwardOnly
    rs_prep.LockType = adLockBatchOptimistic
    rs_prep.ActiveConnection = Db_link_orfeo
        
        SQL = SQL & " SELECT    UPD.COBRAR_PREPAGO COBRAR_PREPAGO " & vbCrLf
        SQL = SQL & " FROM      USUARIO_PERMISO_DISTRIBUCION UPD " & vbCrLf
        SQL = SQL & " WHERE     UPD.NOMBRE_USUARIO      =       '" & num_client & "' " & vbCrLf

        rs_prep.Open SQL
        If Not rs_prep.EOF Then
                If rs_prep.Fields("COBRAR_PREPAGO") = "2" Then
                        Res = "POR COBRAR"
                End If
        End If
        
catch:
        rs_prep.Close
    Set rs_prep = Nothing

        obtener_prepagado_por_cobrar = Res
End Function
Function validar_cantidad_nuis_disponibles(ByVal num_client As String, ByVal nuis_necesarios As Double) As Boolean
        Dim Res As Boolean
        On Error GoTo catch

    Res = False
        If obtener_nuis_disponibles(num_client) >= nuis_necesarios Then
                Res = True
        End If
        
catch:
        validar_cantidad_nuis_disponibles = Res
End Function
Function obtener_nuis_disponibles(ByVal num_client As String) As Double
        Dim Res As Double
    Dim SQL As String
    Dim rs_disp As New ADODB.Recordset
        On Error GoTo catch

    Res = -1
        SQL = ""
    Set rs_disp = New ADODB.Recordset
    rs_disp.CursorLocation = adUseClient
    rs_disp.CursorType = adOpenForwardOnly
    rs_disp.LockType = adLockBatchOptimistic
    rs_disp.ActiveConnection = Db_link_orfeo
        
        SQL = SQL & " SELECT    COUNT(WEL.WELCLAVE) CANTIDAD " & vbCrLf
        SQL = SQL & " FROM              WEB_LTL WEL " & vbCrLf
        SQL = SQL & " WHERE             WEL.WELSTATUS   =       3 " & vbCrLf
        SQL = SQL & "   AND             WEL.WEL_CLICLEF =       '" & num_client & "' " & vbCrLf

        rs_disp.Open SQL
        If Not rs_disp.EOF Then
                If rs_disp.Fields("CANTIDAD") <> "" Then
                        Res = CDbl(rs_disp.Fields("CANTIDAD"))
                End If
        End If
        
catch:
        rs_disp.Close
    Set rs_disp = Nothing

        obtener_nuis_disponibles = Res
End Function
Function obtener_destinatario(ByVal num_client As String, ByVal n_client As String, ByRef cclclave As String, ByRef dieclave As String, ByRef all_clave_dest As Integer)
    Dim SQL As String
    Dim rs As New ADODB.Recordset
        On Error GoTo catch

        SQL = ""
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockBatchOptimistic
    rs.ActiveConnection = Db_link_orfeo
    
        SQL = "SELECT DEC_DIECLAVE, DIE_CCLCLAVE, NVL((SELECT DER_ALLCLAVE FROM EDESTINOS_POR_RUTA WHERE DER_VILCLEF = DIEVILLE AND DER_ALLCLAVE IS NOT NULL AND ROWNUM = 1), 1) ALLCLAVE_DEST "
    SQL = SQL & "   , NVL(DEC_DIECLAVE_ENTREGA, -1) DEC_DIECLAVE_ENTREGA " & vbCrLf
    SQL = SQL & " FROM EDIRECCION_ENTR_CLIENTE_LIGA DECL " & vbCrLf
        SQL = SQL & " INNER JOIN " & vbCrLf
        SQL = SQL & " EDIRECCIONES_ENTREGA DE " & vbCrLf
        SQL = SQL & " ON " & vbCrLf
        SQL = SQL & "  DECL.DEC_DIECLAVE = DE.DIECLAVE " & vbCrLf

    SQL = SQL & " WHERE DEC_CLICLEF = '" & Replace(num_client, "'", "''") & "'" & vbCrLf
    SQL = SQL & "   AND DEC_NUM_DIR_CLIENTE = '" & Replace(n_client, "'", "''") & "'" & vbCrLf
    SQL = SQL & "   AND DIE_STATUS = 1 " & vbCrLf

    SQL = SQL & "   AND EXISTS (" & vbCrLf
    SQL = SQL & "       SELECT NULL FROM EDESTINOS_POR_RUTA DPR" & vbCrLf
        SQL = SQL & " INNER JOIN " & vbCrLf
        SQL = SQL & " EDIRECCIONES_ENTREGA DE1 " & vbCrLf
        SQL = SQL & " ON " & vbCrLf
        SQL = SQL & " DPR.DER_VILCLEF = DE1.DIEVILLE " & vbCrLf
    SQL = SQL & "        WHERE NVL(DER_ALLCLAVE, 1) > 0 " & vbCrLf
    SQL = SQL & "          AND DER_TIPO_ENTREGA NOT IN ('INSEGURO', 'INVALIDO') " & vbCrLf
    SQL = SQL & "          AND SF_LOGIS_CLIENTE_RESTRIC(DEC_CLICLEF, DER_TIPO_ENTREGA) = 1 " & vbCrLf
    SQL = SQL & "       ) " & vbCrLf

        
        rs.Open SQL
    If rs.EOF Then
        obtener_destinatario = "- direccion inexistente, o destino INSEGURO, INVALIDO o TIPO DE ENTREGA no autorizado:" & _
            " id: " & n_client & vbCrLf
    ElseIf rs.RecordCount > 1 Then
        obtener_destinatario = "- existe mas de un registro ligado a este numero direccion:" & _
            " id: " & n_client & vbCrLf
    Else
        obtener_destinatario = "ok"
        cclclave = rs.Fields("DIE_CCLCLAVE")
        dieclave = rs.Fields("DEC_DIECLAVE")
        all_clave_dest = rs.Fields("ALLCLAVE_DEST")
    End If

catch:
        rs.Close
    Set rs = Nothing
End Function
Function obtener_nombre_usuario(ByVal usuario As String)
        Dim SQL As String
        Dim Res As String
        Dim rstNU As New ADODB.Recordset
        On Error GoTo catch

        Res = "DOC_MASIVA"
        Set rstNU = New ADODB.Recordset
        rstNU.CursorLocation = adUseClient
        rstNU.CursorType = adOpenForwardOnly
        rstNU.LockType = adLockBatchOptimistic
        rstNU.ActiveConnection = Db_link_orfeo

        If usuario = "" Then
                usuario = "USER"
        End If

        SQL = ""
        SQL = SQL & " SELECT " & vbCrLf
        SQL = SQL & "   REPLACE " & vbCrLf
        SQL = SQL & "           ( " & vbCrLf
        SQL = SQL & "                   REPLACE(REPLACE(UPPER(SUBSTR('DOC_MASIVA_' || '" & usuario & "',1,29)),':',''),'\','_') " & vbCrLf
        SQL = SQL & "                   ,'.','_' " & vbCrLf
        SQL = SQL & "           ) USR_DOC_MASIV " & vbCrLf
        SQL = SQL & " FROM      DUAL " & vbCrLf

        rstNU.Open SQL
        If Not rstNU.EOF Then
                If rstNU.Fields("USR_DOC_MASIV") <> "" Then
                        Res = rstNU.Fields("USR_DOC_MASIV")
                End If
        End If
        
catch:
        rstNU.Close
    Set rstNU = Nothing
        
        obtener_nombre_usuario = Res
End Function
Function GetAllXLSheetNames_MASIVE( _
              ByVal ExcelFile As String, _
              Optional ByVal HasRow1FieldNames As Boolean = True _
              ) As Collection
        'Make sure you have reference set to Microsoft
        'ActiveX Data Objects 2.X Library (if not default)
        'and Microsoft ADO Ext. 2.X for DDL and Security

        Dim oConn As ADODB.Connection   '
        Dim cat As ADOX.Catalog   '
        Dim tbl As ADOX.Table
        Dim rs As New ADODB.Recordset
        Dim myCollection As New Collection

        'Open the Excel File
        Set oConn = New ADODB.Connection
        Set cat = New ADOX.Catalog

        'Use
        oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                "Data Source=" & ExcelFile & ";" & _
                                "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;" & _
                                IIf(HasRow1FieldNames, "YES", "NO") & ";"""
        Set cat.ActiveConnection = oConn
        
        For Each tbl In cat.Tables
                'you could also stick some code here to work with
                'the sheet directly hence the import of Header YES/NO    Next
                myCollection.Add tbl.Name
        Next
        
        Set rs = Nothing
        Set tbl = Nothing
        Set cat = Nothing
        oConn.Close
        Set oConn = Nothing
        Set GetAllXLSheetNames_MASIVE = myCollection
        Set myCollection = Nothing
End Function
Public Function obtener_cedis_x_remitente(ByVal num_client As String, ByVal disclef As String, ByRef allclave_ori As Integer) As String
        Dim SQL As String
        Dim Res As String
        Dim rs As New ADODB.Recordset
        On Error GoTo catch

        SQL = ""
        Res = ""
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockBatchOptimistic
        rs.ActiveConnection = Db_link_orfeo

        SQL = SQL & " SELECT    DISCLEF DISCLEF, " & vbCrLf
        SQL = SQL & "                   (SELECT DER_ALLCLAVE " & vbCrLf
        SQL = SQL & "                    FROM   EDESTINOS_POR_RUTA " & vbCrLf
        SQL = SQL & "                    WHERE  DER_VILCLEF     =       DISVILLE " & vbCrLf
        SQL = SQL & "                           AND     DER_ALLCLAVE    IS NOT  NULL " & vbCrLf
        SQL = SQL & "                           AND     ROWNUM  = 1) ALLCLAVE_ORI " & vbCrLf
        SQL = SQL & " FROM      EDISTRIBUTEUR " & vbCrLf
        SQL = SQL & " WHERE     DISCLEF         =       '" & Replace(disclef, "'", "''") & "' " & vbCrLf
        SQL = SQL & "   AND     DISCLIENT       =       '" & Replace(num_client, "'", "''") & "' " & vbCrLf

        rs.Open SQL
        If rs.EOF Then
                Res = "- remitente inexistente:" & _
                                " id: " & disclef & vbCrLf & vbCrLf
        Else
                Res = "ok"
                disclef = rs.Fields("DISCLEF")
                allclave_ori = rs.Fields("ALLCLAVE_ORI")
    End If

catch:
    rs.Close
    Set rs = Nothing
        obtener_cedis_x_remitente = Res
End Function
Private Function obtener_nombre_x_cliente(ByVal num_client As String) As String
    Dim NombreCliente As String
    Dim SQL_Nombre As String
    Dim rs_nombre As New ADODB.Recordset
        On Error GoTo catch
    
    Set rs_nombre = New ADODB.Recordset
    rs_nombre.CursorLocation = adUseClient
    rs_nombre.CursorType = adOpenForwardOnly
    rs_nombre.LockType = adLockBatchOptimistic
    rs_nombre.ActiveConnection = Db_link_orfeo
    
    SQL_Nombre = ""
    SQL_Nombre = SQL_Nombre & " SELECT CLINOM || DECODE(CLIALIAS, NULL, NULL, ' - ' || CLIALIAS) NOMBRE_CLIENTE " & vbCrLf
    SQL_Nombre = SQL_Nombre & " FROM ECLIENT " & vbCrLf
    SQL_Nombre = SQL_Nombre & " WHERE CLICLEF = '" & num_client & "'" & vbCrLf
        
    
    rs_nombre.Open SQL_Nombre
    If Not rs_nombre.EOF Then
        NombreCliente = rs_nombre.Fields("NOMBRE_CLIENTE")
    End If
catch:
    rs_nombre.Close
    
    Set rs_nombre = Nothing
    
    obtener_nombre_x_cliente = NombreCliente
End Function
Function registrar_segundos_envios(ByVal nui As String, ByVal num_client As String, ByVal usuario As String) As Boolean
        Dim result As Boolean
        Dim SQL_seg_env As String
        Dim rs_seg_env As New ADODB.Recordset
        Dim choclave_segundos_envios As String
        On Error GoTo catch
        
        result = False
        SQL_seg_env = ""
        choclave_segundos_envios = "-1"
        
        Set rs_seg_env = New ADODB.Recordset
    rs_seg_env.CursorLocation = adUseClient
    rs_seg_env.CursorType = adOpenForwardOnly
    rs_seg_env.LockType = adLockBatchOptimistic
    rs_seg_env.ActiveConnection = Db_link_orfeo
        
        'Buscar la clave del Concepto Correspondiente por Empresa:
        SQL_seg_env = SQL_seg_env & " SELECT    CHOCLAVE " & vbCrLf
        SQL_seg_env = SQL_seg_env & " FROM              ECONCEPTOSHOJA " & vbCrLf
        SQL_seg_env = SQL_seg_env & " WHERE             1 = 1 " & vbCrLf
        SQL_seg_env = SQL_seg_env & "   AND             CHOTIPOIE               =       'I' " & vbCrLf
        SQL_seg_env = SQL_seg_env & "   AND             CHONUMERO               =       517 " & vbCrLf
        SQL_seg_env = SQL_seg_env & "   AND             CHO_EMPCLAVE    =       '" & obtener_clave_empresa(num_client) & "' " & vbCrLf

        rs_seg_env.Open SQL_seg_env
        If Not rs_seg_env.EOF Then
                choclave_segundos_envios = rs_seg_env.Fields("CHOCLAVE")
        End If
        rs_seg_env.Close
        
        If choclave_segundos_envios <> "-1" Then
                'Buscar que al cliente le corresponda este concepto:
                SQL_seg_env = ""
                SQL_seg_env = SQL_seg_env & " SELECT LIG.LIG_CLICLEF CLIENTE " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   ,CLI.CLINOM " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   ,CHO.CHONUMERO " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   ,CHO.CHONOMBRE " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   ,CHO2.CHONUMERO " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   ,CHO2.CHONOMBRE " & vbCrLf
                SQL_seg_env = SQL_seg_env & " FROM      ELIGA_TARIFAS LIG " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   JOIN    ECONCEPTOSHOJA CHO " & vbCrLf
                SQL_seg_env = SQL_seg_env & "           ON      CHO.CHOCLAVE    =       LIG.LIG_CHOCLAVE_APLICA " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   JOIN    ECONCEPTOSHOJA CHO2 " & vbCrLf
                SQL_seg_env = SQL_seg_env & "           ON      CHO2.CHOCLAVE   =       LIG.LIG_CHOCLAVE " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   JOIN    ECLIENT CLI " & vbCrLf
                SQL_seg_env = SQL_seg_env & "           ON      CLI.CLICLEF             =       LIG.LIG_CLICLEF " & vbCrLf
                SQL_seg_env = SQL_seg_env & " WHERE     LIG.LIG_CLICLEF         =       '" & num_client & "' " & vbCrLf
                SQL_seg_env = SQL_seg_env & "   AND     CHO2.CHONUMERO          IN      (517) " & vbCrLf

                rs_seg_env.Open SQL_seg_env
                If Not rs_seg_env.EOF Then
                        If num_client = rs_seg_env.Fields("CLIENTE") Then
                                result = True
                        End If
                End If
                rs_seg_env.Close
        
                If result = True Then
                        'Validar que el NUI no tenga asignado ya el concepto de SEGUNDOS ENVIOS:
                        SQL_seg_env = ""
                        SQL_seg_env = SQL_seg_env & " SELECT    WLCCLAVE ,WLC_WELCLAVE ,WLC_CHOCLAVE " & vbCrLf
                        SQL_seg_env = SQL_seg_env & " FROM      WEB_LTL_CONCEPTOS WLC " & vbCrLf
                        SQL_seg_env = SQL_seg_env & " WHERE     WLC.WLC_WELCLAVE        =       '" & nui & "' " & vbCrLf
                        SQL_seg_env = SQL_seg_env & "   AND     WLC.WLC_CHOCLAVE        =       '" & choclave_segundos_envios & "' " & vbCrLf
                        SQL_seg_env = SQL_seg_env & "   AND     WLC.WLCSTATUS           =       1 " & vbCrLf
                        
                        rs_seg_env.Open SQL_seg_env
                        If Not rs_seg_env.EOF Then
                                If choclave_segundos_envios = rs_seg_env.Fields("WLC_CHOCLAVE") Then
                                        result = True
                                End If
                        End If
                        rs_seg_env.Close
                        
                        If result <> True Then
                                'Si no est� asignado el concepto de segundos envios, se registra en cero:
                                SQL_seg_env = ""
                                SQL_seg_env = SQL_seg_env & " INSERT INTO       WEB_LTL_CONCEPTOS " & vbCrLf
                                SQL_seg_env = SQL_seg_env & " ( " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "    WLCCLAVE " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "   ,WLC_WELCLAVE ,WLC_CHOCLAVE ,WLC_IMPORTE " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "   ,CREATED_BY ,DATE_CREATED ,WLCSTATUS " & vbCrLf
                                SQL_seg_env = SQL_seg_env & " ) " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "   SELECT   SEQ_WEB_LTL_CONCEPTOS.nextval " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "                   ,'" & nui & "' ,'" & choclave_segundos_envios & "' ,'" & obtener_monto_x_concepto(nui, num_client, choclave_segundos_envios) & "' " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "                   ,'" & usuario & "' ,SYSDATE ,1 " & vbCrLf
                                SQL_seg_env = SQL_seg_env & "   FROM     DUAL " & vbCrLf

                                Db_link_orfeo.Execute SQL_seg_env
                                result = True
                        End If
                End If
        End If
        
catch:
        registrar_segundos_envios = result
End Function
Function registrar_recol_domicilio(ByVal nui As String, ByVal num_client As String, ByVal usuario As String) As String
        Dim Res As String
        Dim cant As Double
        Dim choclave As String
        Dim SQL_Reco As String
        Dim rs_reco_domi As New ADODB.Recordset
        On Error GoTo catch
        
        cant = 0
        Res = "-1"
        SQL_Reco = ""
        choclave = "-1"
        
        Set rs_reco_domi = New ADODB.Recordset
    rs_reco_domi.CursorLocation = adUseClient
    rs_reco_domi.CursorType = adOpenForwardOnly
    rs_reco_domi.LockType = adLockBatchOptimistic
    rs_reco_domi.ActiveConnection = Db_link_orfeo
        
        'Validar si el Cliente tiene habilitado el concepto:
        SQL_Reco = SQL_Reco & " SELECT  GET_CHOCLAVE_TRADING(184, WEL.WEL_CLICLEF) CHOCLAVE " & vbCrLf
        SQL_Reco = SQL_Reco & " FROM    WEB_LTL WEL " & vbCrLf
        SQL_Reco = SQL_Reco & " WHERE   WEL.WELCLAVE    =       '" & nui & "' " & vbCrLf
        
        rs_reco_domi.Open SQL_Reco
        If Not rs_reco_domi.EOF Then
                choclave = rs_reco_domi.Fields("CHOCLAVE")
        End If
        rs_reco_domi.Close
        
        If choclave <> "" And choclave <> "-1" Then
                'Validar si el cliente tiene configurada la tarifa de recoleccion:
                SQL_Reco = ""
                SQL_Reco = SQL_Reco & " SELECT  COUNT(0) CANTIDAD " & vbCrLf
                SQL_Reco = SQL_Reco & " FROM     ECONCEPTOSHOJA " & vbCrLf
                SQL_Reco = SQL_Reco & "                 ,ECLIENT_APLICA_CONCEPTOS " & vbCrLf
                SQL_Reco = SQL_Reco & " WHERE    CHOTIPOIE              =       'I' " & vbCrLf
                SQL_Reco = SQL_Reco & "         AND      CHONUMERO              =       184 " & vbCrLf
                SQL_Reco = SQL_Reco & "         AND      CHO_EMPCLAVE   =       GET_EMPRESA_TRADING(" & num_client & ") " & vbCrLf
                SQL_Reco = SQL_Reco & "         AND      CCO_CLICLEF    =       '" & num_client & "' " & vbCrLf
                SQL_Reco = SQL_Reco & "         AND      EXISTS                 ( " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                         SELECT  NULL " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                         FROM    EBASES_POR_CONCEPT " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                         WHERE   BPCCLAVE                =       CCO_BPCCLAVE " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                                 AND     BPC_CHOCLAVE    =       CHOCLAVE " & vbCrLf
                SQL_Reco = SQL_Reco & "                                         ) " & vbCrLf
                SQL_Reco = SQL_Reco & "         AND      NOT EXISTS             ( " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                         SELECT  NULL " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                         FROM    WEB_LTL_CONCEPTOS " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                         WHERE   WLC_CHOCLAVE    =       CHOCLAVE " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                                 AND     WLCSTATUS               =       1 " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                                 AND     WLC_WELCLAVE    =       '" & nui & "' " & vbCrLf
                SQL_Reco = SQL_Reco & "                                                 ) " & vbCrLf
                
                rs_reco_domi.Open SQL_Reco
                If Not rs_reco_domi.EOF Then
                        cant = rs_reco_domi.Fields("CANTIDAD")
                End If
                rs_reco_domi.Close
                
                If cant > 0 Then
                        'Registrar el concepto en 0 para que se recalcule al momento de imprimir el talon:
                        SQL_Reco = ""
                        SQL_Reco = SQL_Reco & " INSERT INTO     WEB_LTL_CONCEPTOS " & vbCrLf
                        SQL_Reco = SQL_Reco & "         ( " & vbCrLf
                        SQL_Reco = SQL_Reco & "                  WLCCLAVE, WLC_WELCLAVE, WLC_CHOCLAVE " & vbCrLf
                        SQL_Reco = SQL_Reco & "                 ,WLC_IMPORTE, CREATED_BY, DATE_CREATED " & vbCrLf
                        SQL_Reco = SQL_Reco & "         ) " & vbCrLf
                        SQL_Reco = SQL_Reco & "         VALUES " & vbCrLf
                        SQL_Reco = SQL_Reco & "                 ( " & vbCrLf
                        SQL_Reco = SQL_Reco & "                         SEQ_WEB_LTL_CONCEPTOS.nextval, '" & nui & "', '" & choclave & "'" & vbCrLf
                        SQL_Reco = SQL_Reco & "                         ,0 ,'" & usuario & "' ,SYSDATE " & vbCrLf
                        SQL_Reco = SQL_Reco & "                 ) " & vbCrLf
                        
                        Db_link_orfeo.Execute SQL_Reco
                End If
        End If
        
catch:
        registrar_recol_domicilio = choclave
End Function
Private Function obtener_monto_x_concepto(ByVal nui As String, ByVal num_client As String, choclave) As Double
        Dim Res As Double
        Dim montoNUI As Double
        Dim chonumero As String
        Dim porcentaje As Double
        Dim SQL_monto_concept As String
        Dim rs_monto_concept As New ADODB.Recordset
        On Error GoTo catch
        
        Res = 0
        montoNUI = 0
        porcentaje = 0
        chonumero = ""
        
        Set rs_monto_concept = New ADODB.Recordset
    rs_monto_concept.CursorLocation = adUseClient
    rs_monto_concept.CursorType = adOpenForwardOnly
    rs_monto_concept.LockType = adLockBatchOptimistic
    rs_monto_concept.ActiveConnection = Db_link_orfeo
        
        'Obtener el monto de Distribucion (CHONUMERO => 172)
        SQL_monto_concept = ""
        SQL_monto_concept = SQL_monto_concept & " SELECT         WLC.WLC_IMPORTE IMPORTE ,WLC.WLCCLAVE CLAVE_CONCEPTO ,WLC.WLC_WELCLAVE WELCLAVE ,WLC.WLC_CHOCLAVE CHOCLAVE " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "                       ,WLC.WLCSTATUS STATUS ,CHO.CHONUMERO CHONUMERO ,WLC.DATE_CREATED FECHA_CREACION " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & " FROM   WEB_LTL_CONCEPTOS WLC " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "       INNER   JOIN    ECONCEPTOSHOJA CHO " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "               ON      WLC.WLC_CHOCLAVE        =       CHO.CHOCLAVE " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & " WHERE  CHO.CHOTIPOIE                  =       'I' " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "       AND      CHO.CHONUMERO                  =       172 /* DISTRIBUCION */ " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "       AND      WLC.WLC_WELCLAVE               =       '" & nui & "' " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & " ORDER BY      WLC.DATE_CREATED        DESC " & vbCrLf
        
        rs_monto_concept.Open SQL_monto_concept
        If Not rs_monto_concept.EOF Then
                montoNUI = rs_monto_concept.Fields("IMPORTE")
        End If
        rs_monto_concept.Close
        
        If montoNUI <= 0 Then
                SQL_monto_concept = ""
                SQL_monto_concept = SQL_monto_concept & " SELECT " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ( " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "               CASE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "                       WHEN    LOGIS.FN_OBTEN_MONTO_DISTRIBUCION(WTS.NUI)      >       0 " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "                               THEN    LOGIS.FN_OBTEN_MONTO_DISTRIBUCION(WTS.NUI) " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "                       WHEN    NVL(NVL(WTS.IMP_DISTRIBUCION,NVL(WEL.WEL_PRECIO_TOTAL, WEL.WEL_PRECIO_ESTIMADO)),0)     >       0 " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "                               THEN    NVL(NVL(WTS.IMP_DISTRIBUCION,NVL(WEL.WEL_PRECIO_TOTAL, WEL.WEL_PRECIO_ESTIMADO)),0) " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "                       ELSE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "                               WEL.WELIMPORTE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "               END " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ) IMPORTE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & " FROM  WEB_LTL WEL " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       INNER   JOIN    WEB_TRACKING_STAGE WTS " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "               ON      WEL.WELCLAVE    =       WTS.NUI " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & " WHERE WEL.WELCLAVE    =       '" & nui & "' " & vbCrLf
                
                rs_monto_concept.Open SQL_monto_concept
                If Not rs_monto_concept.EOF Then
                        montoNUI = rs_monto_concept.Fields("IMPORTE")
                End If
                rs_monto_concept.Close
        End If
        
        'Obtener el CHONUMERO correspondiente a la CHOCLAVE del concepto
        SQL_monto_concept = ""
        SQL_monto_concept = SQL_monto_concept & " SELECT         CHO.CHONUMERO CHONUMERO ,CHO.CHO_EMPCLAVE ,CHO.CHOTIPOIE ,CHO.CHONOMBRE " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "                       ,WLC.WLCCLAVE ,WLC.WLC_WELCLAVE ,WLC.WLC_CHOCLAVE ,WLC.WLC_IMPORTE ,WLC.DATE_CREATED ,WLC.WLCSTATUS " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & " FROM  WEB_LTL_CONCEPTOS WLC " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "       INNER   JOIN    ECONCEPTOSHOJA CHO " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "               ON      WLC.WLC_CHOCLAVE        =       CHO.CHOCLAVE " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & " WHERE CHO.CHOTIPOIE           =       'I' " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "       AND     WLC.WLC_WELCLAVE        =       '" & nui & "' " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & "       AND WLC.WLC_CHOCLAVE    =       '" & choclave & "' " & vbCrLf
        SQL_monto_concept = SQL_monto_concept & " ORDER BY      WLC.WLCSTATUS DESC " & vbCrLf

        rs_monto_concept.Open SQL_monto_concept
        If Not rs_monto_concept.EOF Then
                chonumero = rs_monto_concept.Fields("CHONUMERO")
        End If
        rs_monto_concept.Close
        
        'Obtener el Porcentaje por concepto
        If chonumero <> "" Then
                SQL_monto_concept = ""
                SQL_monto_concept = SQL_monto_concept & " SELECT " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "        LIG.LIGPORCENTAJE_APLICA PORCENTAJE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ,LIG.LIG_CLICLEF CLIENTE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ,CLI.CLINOM NOMBRE_CLIENTE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ,CHO.CHONUMERO CHONUMERO " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ,CHO.CHONOMBRE CHONOMBRE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ,CHO2.CHONUMERO CHONUMERO2 " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       ,CHO2.CHONOMBRE CHONOMBRE2 " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & " FROM  ELIGA_TARIFAS LIG " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       JOIN    ECONCEPTOSHOJA CHO " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "               ON      CHO.CHOCLAVE    =       LIG.LIG_CHOCLAVE_APLICA " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       JOIN    ECONCEPTOSHOJA CHO2 " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "               ON      CHO2.CHOCLAVE   =       LIG.LIG_CHOCLAVE " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       JOIN    ECLIENT CLI " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "               ON      CLI.CLICLEF             =       LIG.LIG_CLICLEF " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & " WHERE CHO2.CHONUMERO          =       '" & chonumero & "' " & vbCrLf
                SQL_monto_concept = SQL_monto_concept & "       AND     LIG.LIG_CLICLEF         =       '" & num_client & "' " & vbCrLf
                
                rs_monto_concept.Open SQL_monto_concept
                If Not rs_monto_concept.EOF Then
                        porcentaje = rs_monto_concept.Fields("PORCENTAJE")
                End If
                rs_monto_concept.Close
        End If
        
        'Calcular el Monto del Concepto multiplicando el Monto de Distribucion por el Porcentaje
        If porcentaje <> 0 Then
                Res = montoNUI * (porcentaje / 100)
        End If
        
catch:
        obtener_monto_x_concepto = Res
End Function
Function borrar_id_cron(ByVal id_cron) As Boolean
        On Error GoTo catch
        Dim SQL As String
        Dim Res As Boolean
        
        SQL = "DELETE FROM REP_DETALLE_REPORTE WHERE ID_CRON = '" & id_cron & "'"
        Db_link_orfeo.Execute SQL
catch:
        If Err.Number <> 0 Then
                Res = False
        Else
                Res = True
        End If
        
        borrar_id_cron = Res
End Function
Function notifica_error(ByVal num_client As String, ByVal correo_electronico As String, ByVal Archivo As String, ByVal Errores As String)
        On Error GoTo catch
        jmail.From = mail_From
        jmail.FromName = mail_FromName
        jmail.ClearRecipients

        jmail.AddRecipientBCC mail_grupo_error(0)
        jmail.AddRecipient "cargamasiva_smo@logis.com.mx"

        For i = 0 To UBound(Split(Replace(correo_electronico, ",", ";"), ";"))
                jmail.AddRecipient Trim(Split(Replace(correo_electronico, ",", ";"), ";")(i))
        Next
        
        jmail.subject = "Error carga de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "(" & obtener_nombre_x_cliente(num_client) & ")"
        jmail.body = "Hola, se detectaron errores al cargar el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ", favor de revisar cada uno de ellos y volver a cargar el archivo." & vbCrLf & vbCrLf & Errores
        
        If FSO.FileExists(Archivo) Then
                jmail.AddAttachment Archivo
        End If
        
        jmail.Send mail_server
        
        If FSO.FileExists(Archivo) Then
                FSO.DeleteFile (Archivo)
        End If

catch:
End Function
Function notifica_exito(ByVal num_client As String, ByVal correo_electronico As String, ByVal Archivo As String, ByVal cantidad_nuis As Double, ByVal lista_nuis As String)
        On Error GoTo catch
        jmail.From = mail_From
        jmail.FromName = mail_FromName
        jmail.ClearRecipients

        jmail.AddRecipientBCC mail_grupo_error(0)
        jmail.AddRecipient "cargamasiva_smo@logis.com.mx"

        For i = 0 To UBound(Split(Replace(correo_electronico, ",", ";"), ";"))
                jmail.AddRecipient Trim(Split(Replace(correo_electronico, ",", ";"), ";")(i))
        Next
        
        jmail.subject = "Exito carga de archivo web " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & "(" & obtener_nombre_x_cliente(num_client) & ")"
        jmail.body = "Hola, se cargo exitosamente el archivo " & Split(Archivo, "\")(UBound(Split(Archivo, "\"))) & ". " & vbCrLf & vbCrLf
        
        If cantidad_nuis = 1 Then
                jmail.body = jmail.body & "     - Se registro un NUI. " & vbCrLf & vbCrLf
        ElseIf cantidad_nuis > 1 Then
                jmail.body = jmail.body & "     - Se registraron en total " & cantidad_nuis & " NUIs. " & vbCrLf & vbCrLf
        End If
        
        If lista_nuis <> "" Then
                jmail.body = jmail.body & lista_nuis & vbCrLf & vbCrLf
        End If
        
        jmail.body = jmail.body & vbCrLf & "Saludos."
        
        If FSO.FileExists(Archivo) Then
                jmail.AddAttachment Archivo
        End If
        
        jmail.Send mail_server
        
        If FSO.FileExists(Archivo) Then
                FSO.DeleteFile (Archivo)
        End If

catch:
End Function
Function registrar_tracking_stage_doc(ByVal nui As String, ByVal usuario As String) As String
	Dim Res As String
	Dim SQL_wts As String
	Dim rs_wts As New ADODB.Recordset
	On Error GoTo catch
	
	Res = "-1"
	SQL_wts = ""
        
	Set rs_wts = New ADODB.Recordset
	rs_wts.CursorLocation = adUseClient
	rs_wts.CursorType = adOpenForwardOnly
	rs_wts.LockType = adLockBatchOptimistic
	rs_wts.ActiveConnection = Db_link_orfeo
	
	'Registrar usuario y fecha de documentacion por NUI:
	SQL_wts = SQL_wts & " UPDATE	 WEB_TRACKING_STAGE " & vbCrLf
	SQL_wts = SQL_wts & " 	SET	 USR_DOC				=	'" & usuario & "' " & vbCrLf
	SQL_wts = SQL_wts & " 		,FECHA_DOCUMENTACION	=	SYSDATE " & vbCrLf
	SQL_wts = SQL_wts & " WHERE	 NUI					=	'" & iNUI & "' " & vbCrLf
	
	Db_link_orfeo.Execute SQL_wts
	Res = "ok"
        
catch:
	rs_wts = Nothing
	registrar_tracking_stage_doc = Res
End Function
Function registrar_tracking_stage_can(ByVal nui As String, ByVal usuario As String) As String
	Dim Res As String
	Dim SQL_wts As String
	Dim rs_wts As New ADODB.Recordset
	On Error GoTo catch
	
	Res = "-1"
	SQL_wts = ""
        
	Set rs_wts = New ADODB.Recordset
	rs_wts.CursorLocation = adUseClient
	rs_wts.CursorType = adOpenForwardOnly
	rs_wts.LockType = adLockBatchOptimistic
	rs_wts.ActiveConnection = Db_link_orfeo
	
	'Registrar usuario y fecha de documentacion por NUI:
	SQL_wts = SQL_wts & " UPDATE	 WEB_TRACKING_STAGE " & vbCrLf
	SQL_wts = SQL_wts & " 	SET	 USR_CAN			=	'CAN_" & usuario & "' " & vbCrLf
	SQL_wts = SQL_wts & " 		,FECHA_CANCELACION	=	SYSDATE " & vbCrLf
	SQL_wts = SQL_wts & " WHERE	 NUI				=	'" & iNUI & "' " & vbCrLf
	
	Db_link_orfeo.Execute SQL_wts
	Res = "ok"
        
catch:
	rs_wts = Nothing
	registrar_tracking_stage_can = Res
End Function
Function registrar_tarimas(ByVal iNUI As String, ByVal usuario As String, i_Tarimas As Double, i_BultosConstitutivos As Double) As String
	Dim Res As String
	Dim SQL_tar As String
	On Error GoTo catch
	
	Res = "-1"
	SQL_tar = ""
	
	'Registrar Tarimas y Bultos constitutivos por Tarima al NUI:
	SQL_tar = SQL_tar & " INSERT INTO	TB_LOGIS_WPALETA_LTL " & vbCrLf
	SQL_tar = SQL_tar & " 	( " & vbCrLf
	SQL_tar = SQL_tar & " 		 WPLCLAVE ,WPL_WELCLAVE " & vbCrLf
	SQL_tar = SQL_tar & " 		,WPL_IDENTICAS ,WPL_TPACLAVE " & vbCrLf
	SQL_tar = SQL_tar & " 		,WPLLARGO ,WPLANCHO ,WPLALTO " & vbCrLf
	SQL_tar = SQL_tar & " 		,WPL_CDAD_EMPAQUES_X_BULTO ,WPL_BULTO_TPACLAVE " & vbCrLf
	SQL_tar = SQL_tar & " 		,CREATED_BY ,DATE_CREATED " & vbCrLf
	SQL_tar = SQL_tar & " 	) " & vbCrLf
	SQL_tar = SQL_tar & " 	VALUES " & vbCrLf
	SQL_tar = SQL_tar & " 		( " & vbCrLf
	SQL_tar = SQL_tar & "			 SEQ_WPALETA_LTL.nextval ,'" & iNUI & "' " & vbCrLf
	SQL_tar = SQL_tar & "			,'" & i_Tarimas & "' ,1 " & vbCrLf
	SQL_tar = SQL_tar & "			,0 ,0 ,0 " & vbCrLf
	SQL_tar = SQL_tar & "			, '" & i_BultosConstitutivos & "' ,9 " & vbCrLf
	SQL_tar = SQL_tar & "			, '" & usuario & "' ,SYSDATE " & vbCrLf
	SQL_tar = SQL_tar & " 		) " & vbCrLf
	
	Db_link_orfeo.Execute SQL_tar
	Res = "ok"
        
catch:
	registrar_tarimas = Res
End Function
Function registrar_bultos_granel(ByVal iNUI As String, ByVal usuario As String, i_BultosGranel As Double) As String
	Dim Res As String
	Dim SQL_tar As String
	On Error GoTo catch
	
	Res = "-1"
	SQL_tar = ""
	
	'Registrar Bultos Sueltos al NUI:
	SQL = SQL & " INSERT INTO       TB_LOGIS_WPALETA_LTL " & vbCrLf
	SQL = SQL & " 	( " & vbCrLf
	SQL = SQL & " 		 WPLCLAVE ,WPL_WELCLAVE " & vbCrLf
	SQL = SQL & " 		,WPL_IDENTICAS ,WPL_TPACLAVE " & vbCrLf
	SQL = SQL & " 		,WPLLARGO ,WPLANCHO ,WPLALTO " & vbCrLf
	SQL = SQL & " 		,CREATED_BY ,DATE_CREATED " & vbCrLf
	SQL = SQL & " 	) " & vbCrLf
	SQL = SQL & " 	VALUES " & vbCrLf
	SQL = SQL & " 	( " & vbCrLf
	SQL = SQL & " 		 SEQ_WPALETA_LTL.nextval ,'" & iNUI & "' " & vbCrLf
	SQL = SQL & " 		,'" & i_BultosGranel & "' ,9 " & vbCrLf
	SQL = SQL & " 		,0 ,0 ,0 " & vbCrLf
	SQL = SQL & " 		, '" & usuario & "' ,SYSDATE " & vbCrLf
	SQL = SQL & " 	) " & vbCrLf
	
	Db_link_orfeo.Execute SQL_tar
	Res = "ok"
        
catch:
	registrar_bultos_granel = Res
End Function
Function documentar_nuevo_nui(cliente As String, usuario As String, s_CondicionesEntrega As String, s_Observaciones As String, s_Referencia As String, mi_disclef As String, ccl_clave As String, _
							  die_clave As String, i_ValorMercancia As Double, allclave_ori As Integer, allclave_dest As Integer, i_BultosTotales As Double, i_Tarimas As Double, i_BultosConstitutivos As Double, i_BultosGranel As Double) As Double
	'Ciclos determinantes: Documentar valores acumulados y reiniciar variables:
	Dim iNUI As Double
	Dim SQL As String
	
	On Error GoTo catch
	iNUI = -1
	SQL = ""
	
	s_CondicionesEntrega = obtener_prepagado_por_cobrar(cliente)
	s_Observaciones = s_Observaciones & " " & obtener_dice_contener(cliente)
	iNUI = obtener_nui_disponible(cliente)
	
	Debug.Print "Inicia documentacion del NUI: " & iNUI & vbCrLf
	
	SQL = SQL & " UPDATE	WEB_LTL " & vbCrLf
	SQL = SQL & " 	SET " & vbCrLf
	SQL = SQL & " 		 WELSTATUS				=	1 " & vbCrLf
	SQL = SQL & " 		,DATE_CREATED			=	SYSDATE " & vbCrLf
	SQL = SQL & " 		,MODIFIED_BY			=	'" & usuario & "' " & vbCrLf
	SQL = SQL & " 		,WEL_COLLECT_PREPAID	=	'" & s_CondicionesEntrega & "' " & vbCrLf
	SQL = SQL & " 		,WELOBSERVACION			=	SUBSTR('" & s_Observaciones & "',1,1999) " & vbCrLf
	
	If s_Referencia = "" Then
		SQL = SQL & "		,WELFACTURA		=	'_PENDIENTE_' " & vbCrLf
	Else
		SQL = SQL & "		,WELFACTURA		=	'" & s_Referencia & "' " & vbCrLf
	End If
	If mi_disclef <> "" Then
		SQL = SQL & "		,WEL_DISCLEF	=	'" & mi_disclef & "' " & vbCrLf
	End If
	If ccl_clave <> "" Then
		SQL = SQL & "		,WEL_CCLCLAVE	=	'" & ccl_clave & "' " & vbCrLf
	End If
	If die_clave <> "" Then
		SQL = SQL & "		,WEL_DIECLAVE	=	'" & die_clave & "' " & vbCrLf
	End If
	
	If allclave_ori <> -1 Then
		SQL = SQL & "		,WEL_ALLCLAVE_ORI	=	'" & allclave_ori & "' " & vbCrLf
	End If
	If allclave_dest <> -1 Then
		SQL = SQL & "		,WEL_ALLCLAVE_DEST	=	'" & allclave_dest & "' " & vbCrLf
	End If
	
	If i_ValorMercancia > 0 Then
		SQL = SQL & "		,WELIMPORTE		=	'" & i_ValorMercancia & "' " & vbCrLf
	Else
		SQL = SQL & "		,WELIMPORTE		=	0 " & vbCrLf
	End If
	If i_BultosTotales > 0 Then
		SQL = SQL & "		,WEL_CDAD_BULTOS	=	'" & i_BultosTotales & "' " & vbCrLf
	Else
		SQL = SQL & "		,WEL_CDAD_BULTOS	=	0 " & vbCrLf
	End If
	If i_Tarimas > 0 Then
		SQL = SQL & "		,WEL_CDAD_TARIMAS	=	'" & i_Tarimas & "' " & vbCrLf
	Else
		SQL = SQL & "		,WEL_CDAD_TARIMAS	=	0 " & vbCrLf
	End If
	If i_BultosConstitutivos > 0 Then
		SQL = SQL & "		,WEL_CAJAS_TARIMAS	=	'" & i_BultosConstitutivos & "' " & vbCrLf
	Else
		SQL = SQL & "		,WEL_CAJAS_TARIMAS	=	0 " & vbCrLf
	End If
	If i_BultosGranel > 0 Then
		SQL = SQL & "		,WELCDAD_CAJAS		=	'" & i_BultosGranel & "' " & vbCrLf
	Else
		SQL = SQL & "		,WELCDAD_CAJAS		=	0 " & vbCrLf
	End If

	SQL = SQL & " WHERE	WELCLAVE	=	'" & iNUI & "' " & vbCrLf
	Db_link_orfeo.Execute SQL
	
	
	Call registrar_tracking_stage_doc(iNUI,usuario)
	
	If i_Tarimas > 0 Then
		Call registrar_tarimas(iNUI, usuario, i_Tarimas, i_BultosConstitutivos)
	End If
	
	If i_BultosGranel > 0 Then
		Call registrar_bultos_granel(iNUI,usuario,i_BultosGranel)
	End If
	
	Call CHECK_VALID_LTL(iNUI)
	Call registrar_segundos_envios(iNUI, cliente, usuario)
	Call registrar_recol_domicilio(iNUI, cliente, usuario)
	Call registrar_bitacora(cliente, "DOC_MASIVA-SIN_FACTURA", iNUI, "LTL", usuario)
	
	Debug.Print "Termina documentacion del NUI: " & iNUI & vbCrLf
	
catch:
	documentar_nuevo_nui = iNUI
End Function
Function registrar_bitacora(ByVal cliente As String, ByVal modulo As String, ByVal nui As Double, ByVal tipo As String, ByVal usuario As String) As Boolean
	Dim SQL As String
	Dim Total As Double
	Dim ip_site As String
	Dim rst_bita As New ADODB.Recordset
	Dim res As Boolean
	On Error GoTo catch
	
	SQL = ""
	Total = 0
	ip_site = "192.168.100.4"
	res = false
	Set rst_bita = New ADODB.Recordset
	rst_bita.CursorLocation = adUseClient
	rst_bita.CursorType = adOpenForwardOnly
	rst_bita.LockType = adLockBatchOptimistic
	rst_bita.ActiveConnection = Db_link_orfeo
	
	If cliente <> "" And modulo <> "" And nui > 0 And tipo <> "" And usuario <> "" Then
		SQL = SQL & " SELECT	COUNT(*) CANTIDAD " & vbCrLf
		SQL = SQL & " FROM	WEB_BITA_DOCUMENTA " & vbCrLf
		SQL = SQL & " WHERE	WBD_CLICLEF		=	'" & cliente & "' " & vbCrLf
		SQL = SQL & " 	AND	WBD_USUARIO		=	'" & usuario & "' " & vbCrLf
		SQL = SQL & " 	AND	WBD_MODULO		=	'" & modulo & "' " & vbCrLf
		SQL = SQL & " 	-- limitado a un registro por minuto para cada cliente, ya que se trata de cargas masivas y se puede saturar la tabla (validar si es funcional): " & vbCrLf
		SQL = SQL & " 	AND	TO_CHAR(WBD_FECHA,'DD/MM/YYYY HH24MI')	=	TO_CHAR(SYSDATE,'DD/MM/YYYY HH24MI') " & vbCrLf
		
		rst_bita.Open SQL_monto_concept
		If Not rst_bita.EOF Then
			Total = rst_bita.Fields("CANTIDAD")
		End If
		rst_bita.Close
		
		If Total <= 0 Then
			SQL = ""
			SQL = SQL & " INSERT INTO	WEB_BITA_DOCUMENTA " & vbCrLf
			SQL = SQL & " ( " & vbCrLf
			SQL = SQL & " 	 WBD_ID_EVENTO " & vbCrLf
			SQL = SQL & " 	,WBD_FECHA " & vbCrLf
			SQL = SQL & " 	,WBD_CLICLEF " & vbCrLf
			SQL = SQL & " 	,WBD_MODULO " & vbCrLf
			SQL = SQL & " 	,WBD_USUARIO " & vbCrLf
			SQL = SQL & " 	,NUI " & vbCrLf
			SQL = SQL & " 	,TIPO " & vbCrLf
			SQL = SQL & " 	,WBD_IP_SERVIDOR " & vbCrLf
			SQL = SQL & " ) " & vbCrLf
			SQL = SQL & " 	VALUES " & vbCrLf
			SQL = SQL & " 	( " & vbCrLf
			SQL = SQL & " 		(SELECT MAX(NVL(WBD_ID_EVENTO,0)) + 1 FROM WEB_BITA_DOCUMENTA) " & vbCrLf
			SQL = SQL & " 		,SYSDATE " & vbCrLf
			SQL = SQL & " 		,'" & cliente & "' " & vbCrLf
			SQL = SQL & " 		,'" & modulo & "' " & vbCrLf
			SQL = SQL & " 		,'" & usuario & "' " & vbCrLf
			SQL = SQL & " 		,'" & nui & "' " & vbCrLf
			SQL = SQL & " 		,'" & tipo & "' " & vbCrLf
			SQL = SQL & " 		,'" & ip_site & "' " & vbCrLf
			SQL = SQL & " 	) " & vbCrLf
			
			Db_link_orfeo.Execute SQL_Reco
			res = True
		End If
	End If
catch:
	registrar_bitacora = res
end function 