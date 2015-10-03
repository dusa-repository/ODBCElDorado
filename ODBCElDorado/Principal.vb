Imports System.Xml
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data


Module Principal

    Public conexionString As String = "Data Source=localhost;Database=dusa_dorado;User ID=client;Password=123;"
    Public conexionStringPedidos As String = "Data Source=localhost;Database=dusa_dorado;User ID=client;Password=123;"
    Public comando As New SqlClient.SqlCommand
    Public conexion As SqlClient.SqlConnection


    Public nombrePartner As String
    Dim logger As StreamWriter
    Dim lineaLogger As String
    Dim prefijo As String
    Public cadenaAS400 As String
    Public cadenaAS400CTL As String

    

    Sub Main()
        'DUSA

        'procesos_IM()
        'procesar_OUTB("DO")
        'procesar_OUTB("DD")
        procesos_STO()
        'procesos_PO()
        'crearGoodReceiptsPO()
        'crearGoodReceiptsSTO()
        'crearStockReconciliation("LL1")
        'crearStockReconciliation("LL2")

    End Sub


    Private Sub procesos_IM()

        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String


        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()

            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_IM.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")
            'bk = diccionario.Item("backup")
            'crearIM()



        Catch oe As Exception
            escribirLog(oe.Message.ToString, "(STO) ")
        Finally

            logger.Close()
        End Try


        Try
            Dim linea As String

            Dim cnn2 As New Odbc.OdbcConnection(cadenaAS400)
            Dim rs2 As New Odbc.OdbcCommand("SELECT DGUKID,DGICU,DGAA18,DGAA04,DGCDCTYPE,DGLOTN,DGTRQT,DGORGU,DGGMSTS,DGEV02,DGEV01,DGGASTS FROM F55DM WHERE DGBCTK=0    ", cnn2)
            Dim reader2 As Odbc.OdbcDataReader

            cnn2.Open()
            reader2 = rs2.ExecuteReader
            While reader2.Read()

                conIn.ConnectionString = conexionString
                conIn.Open()

                Dim nombreArchivo As String

                nombreArchivo = nombrePartner & "_ZWHINV_" & obtenerFechaHora() & ".xml"

                Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombreArchivo)


                Dim nextn As String

                nextn = Right("00000000000000" & obtenerNextNumber(), 14)
                linea = ""
                linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                linea = linea & "<ZWMMBID2>"
                linea = linea & "<IDOC BEGIN=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<EDI_DC40>"
                linea = linea & "<TABNAM>EDI_DC40</TABNAM>"
                linea = linea & "<MANDT>600</MANDT>"
                linea = linea & "<DOCNUM>" & nextn & "</DOCNUM>"
                linea = linea & "<DIRECT />"
                linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>"
                linea = linea & "<CIMTYP>ZWMMBID2</CIMTYP>"
                linea = linea & "<MESTYP>ZWHINV</MESTYP>"
                linea = linea & "<SNDPOR />"
                linea = linea & "<SNDPRT>LS</SNDPRT>"
                linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>"
                linea = linea & "<RCVPOR />"
                linea = linea & "<RCVPRT>LS</RCVPRT>"
                linea = linea & "<RCVPRN></RCVPRN>"
                linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>"
                linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>"
                linea = linea & "</EDI_DC40>"
                linea = linea & "<E1MBXYH SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<BKTXT>" & nextn & "</BKTXT>"
                linea = linea & "<ZGREC01 SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<DOC_TYP>X</DOC_TYP>"
                linea = linea & "</ZGREC01>"
                linea = linea & "<ZSTKDATE SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<DATE>" & obtenerFecha() & "</DATE>"
                linea = linea & "<TIME>" & obtenerHora() & "</TIME>"
                linea = linea & "</ZSTKDATE>"
                linea = linea & "<E1MBXYI SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<MATNR>" & Right("000000000000000000" & Trim(reader2("DGAA18")), 18) & "</MATNR>"
                linea = linea & "<WERKS>" & Trim(reader2("DGAA04")) & "</WERKS>"
                linea = linea & "<LGORT>" & Trim(reader2("DGCDCTYPE")) & "</LGORT>"
                linea = linea & "<CHARG>" & Trim(reader2("DGLOTN")) & "</CHARG>"
                linea = linea & "<ERFMG>" & Math.Abs(reader2("DGTRQT")) / 10000 & "</ERFMG>"
                linea = linea & "<ERFME>" & Trim(reader2("DGORGU")) & "</ERFME>"
                linea = linea & "<UMLGO></UMLGO>"
                linea = linea & "<UMCHA></UMCHA>"
                linea = linea & "<ZINVMV1 SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<FROM_STOCK_STAT>" & Trim(reader2("DGGMSTS")) & "</FROM_STOCK_STAT>"
                linea = linea & "<MESSAGE_TYPE>" & Trim(reader2("DGEV02")) & "</MESSAGE_TYPE>"
                linea = linea & "</ZINVMV1>"
                linea = linea & "<ZGREC02 SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<DUTY_STATUS>" & Trim(reader2("DGEV01")) & "</DUTY_STATUS>"
                linea = linea & "<STOCK_STAT>" & Trim(reader2("DGGASTS")) & "</STOCK_STAT>"
                linea = linea & "</ZGREC02>"
                linea = linea & "</E1MBXYI>"
                linea = linea & "</E1MBXYH>"
                linea = linea & "</IDOC>"
                linea = linea & "</ZWMMBID2>"

                oSW.WriteLine(linea)
                oSW.Flush()
                oSW.Close()

                Dim cnn3 As New Odbc.OdbcConnection(cadenaAS400)
                Dim rs3 As New Odbc.OdbcCommand("UPDATE F55DM SET DGBCTK=1 WHERE DGUKID=" & reader2("DGUKID") & "  ", cnn3)
                Dim reader3 As Odbc.OdbcDataReader

                cnn3.Open()
                reader3 = rs3.ExecuteReader
                reader3.Close()
                cnn3.Close()


              

                System.Threading.Thread.Sleep(2000)

            End While
           
            cnn2.Close()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try





    End Sub

    Private Sub crearIM()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String

        Try


            Dim linea As String


            conEx.ConnectionString = conexionString
            conEx.Open()
            cmdEx.Connection = conEx
            cmdEx.CommandText = "SELECT * FROM INVENTORY_MOVEMENTS WHERE GENERADO='N'"
            Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()

            While lrdEx.Read()

                conIn.ConnectionString = conexionString
                conIn.Open()


                Dim nombreArchivo As String

                nombreArchivo = nombrePartner & "_ZWHINV_" & obtenerFechaHora() & ".xml"

                Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombreArchivo)


                Dim nextn As String

                nextn = Right("00000000000000" & obtenerNextNumber(), 14)
                linea = ""
                linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                linea = linea & "<ZWMMBID2>"
                linea = linea & "<IDOC BEGIN=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<EDI_DC40>"
                linea = linea & "<TABNAM>EDI_DC40</TABNAM>"
                linea = linea & "<MANDT>600</MANDT>"
                linea = linea & "<DOCNUM>" & nextn & "</DOCNUM>"
                linea = linea & "<DIRECT />"
                linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>"
                linea = linea & "<CIMTYP>ZWMMBID2</CIMTYP>"
                linea = linea & "<MESTYP>ZWHINV</MESTYP>"
                linea = linea & "<SNDPOR />"
                linea = linea & "<SNDPRT>LS</SNDPRT>"
                linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>"
                linea = linea & "<RCVPOR />"
                linea = linea & "<RCVPRT>LS</RCVPRT>"
                linea = linea & "<RCVPRN></RCVPRN>"
                linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>"
                linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>"
                linea = linea & "</EDI_DC40>"
                linea = linea & "<E1MBXYH SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<BKTXT>" & nextn & "</BKTXT>"
                linea = linea & "<ZGREC01 SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<DOC_TYP>X</DOC_TYP>"
                linea = linea & "</ZGREC01>"
                linea = linea & "<ZSTKDATE SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<DATE>" & obtenerFecha() & "</DATE>"
                linea = linea & "<TIME>" & obtenerHora() & "</TIME>"
                linea = linea & "</ZSTKDATE>"
                linea = linea & "<E1MBXYI SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<MATNR>" & Right("000000000000000000" & lrdEx.GetString(1), 18) & "</MATNR>"
                linea = linea & "<WERKS>" & lrdEx.GetString(2) & "</WERKS>"
                linea = linea & "<LGORT>" & lrdEx.GetString(3) & "</LGORT>"
                linea = linea & "<CHARG>" & lrdEx.GetString(4) & "</CHARG>"
                linea = linea & "<ERFMG>" & lrdEx.GetString(5) & "</ERFMG>"
                linea = linea & "<ERFME>" & lrdEx.GetString(6) & "</ERFME>"
                linea = linea & "<UMLGO>" & lrdEx.GetString(7) & "</UMLGO>"
                linea = linea & "<UMCHA>" & lrdEx.GetString(8) & "</UMCHA>"
                linea = linea & "<ZINVMV1 SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<FROM_STOCK_STAT>" & lrdEx.GetString(9) & "</FROM_STOCK_STAT>"
                linea = linea & "<MESSAGE_TYPE>" & lrdEx.GetString(11) & "</MESSAGE_TYPE>"
                linea = linea & "</ZINVMV1>"
                linea = linea & "<ZGREC02 SEGMENT=" & Chr(34) & Chr(34) & ">"
                linea = linea & "<DUTY_STATUS>P</DUTY_STATUS>"
                linea = linea & "<STOCK_STAT>" & lrdEx.GetString(10) & "</STOCK_STAT>"
                linea = linea & "</ZGREC02>"
                linea = linea & "</E1MBXYI>"
                linea = linea & "</E1MBXYH>"
                linea = linea & "</IDOC>"
                linea = linea & "</ZWMMBID2>"


                oSW.WriteLine(linea)
                oSW.Flush()
                oSW.Close()



                System.Threading.Thread.Sleep(2000)


                conIn.Close()


            End While

            actualizarGeneradoIM()

            conEx.Close()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


    End Sub

    Private Function actualizarGeneradoIM() As Integer

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim sqlstring As String

        Try

            conEx2.ConnectionString = conexionString
            conEx2.Open()

            sqlstring = ""
            sqlstring = "UPDATE INVENTORY_MOVEMENTS SET GENERADO='S'  "
            comando.Connection = conEx2
            comando.CommandText = " "
            comando.CommandText = sqlstring
            comando.ExecuteNonQuery()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try

        Return valor
    End Function

    Private Sub procesar_OUTB(ByVal proceso As String)

        Dim host As String
        Dim host_pedidos As String
        Dim database As String
        Dim database_pedidos As String
        Dim user As String
        Dim password As String
        Dim user_pedidos As String
        Dim password_pedidos As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_OUT.xml", FileMode.Open, FileAccess.Read)
            xmldoc = New XmlDataDocument()
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")

            prefijo = diccionario.Item("prefijo_archivo")
            host_pedidos = diccionario.Item("host_pedidos")
            database_pedidos = diccionario.Item("database_pedidos")
            user_pedidos = diccionario.Item("user_pedidos")
            password_pedidos = diccionario.Item("password_pedidos")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            conexionStringPedidos = "Data Source=" & host_pedidos & ";Database=" & database_pedidos & ";User ID=" & user_pedidos & ";Password=" & password_pedidos & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

            escribirLog(conexionString, "(OUTB) ")

            If proceso = "DO" Then
                procesar_OUTBOUND()
                crearDeliveryAcknowledgement()
            Else
                crearDeliveryDispatch()
            End If


        Catch oe As Exception
            escribirLog(oe.Message.ToString, "(OUTB) ")
        Finally

            logger.Close()
        End Try



    End Sub

    Private Sub procesar_OUTBOUND()

        Dim DOCNUM As String
        Dim CREDAT As String
        Dim CRETIM As String
        Dim VBELN As String
        Dim VSTEL As String
        Dim POSNR As String
        Dim MATNR As String
        Dim ARKTX As String
        Dim MATKL As String
        Dim WERKS As String
        Dim LGORT As String
        Dim LFIMG As String
        Dim VRKME As String
        Dim NTGEW As String
        Dim GEWEI As String
        Dim VOLUM As String
        Dim LGMNG As String
        Dim MEINS As String
        Dim PARTNER_ID As String
        Dim NAME1 As String
        Dim VGBEL As String
        Dim TIPODOCUMENTO As String

        Dim conexion As New SqlConnection
        Dim conexion_pedidos As New SqlConnection
        Dim myTrans As SqlTransaction
        Dim myTransPedidos As SqlTransaction
        Dim comando As New SqlClient.SqlCommand
        Dim comando_pedidos As New SqlClient.SqlCommand
        Dim sqlstring As String

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim listaNodos As New List(Of Nodo)
        Dim nombreArchivoProcesar As String

        Dim listaArchivosProcesar As New List(Of String)
        Dim listaDiccionario As New List(Of Dictionary(Of String, String))
        Dim diccionarioDetalle As New Dictionary(Of String, String)
        Dim iw As Integer
        Dim insertar As Boolean
        Dim SZEKCO As String
        Dim SZMCU As String
        Dim SZAN8 As String
        Dim ZFUSER As String
        Dim esDO As Boolean
        esDO = False


        LGORT = ""

        listaArchivosProcesar = obtenerNombreArchivo(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", ""), prefijo)
        For iw = 0 To listaArchivosProcesar.Count - 1

            Try
                nombreArchivoProcesar = listaArchivosProcesar.Item(iw)
                If nombreArchivoProcesar.Trim.CompareTo("") = 0 Then
                    escribirLog("No se encuentra el archivo necesitado con prefijo " & prefijo & " y extension .xml !", "(OUTB) ")
                Else

                    Dim fs As New FileStream(nombreArchivoProcesar, FileMode.Open, FileAccess.Read)
                    xmldoc = New XmlDataDocument()
                    xmldoc.Load(fs)
                    diccionario = obtenerNodosHijosDePadre("EDI_DC40", xmldoc)
                    DOCNUM = diccionario.Item("DOCNUM")

                    escribirLog("INICIO DE IMPORTACION DE DATOS DE DOCUMENTO:" & DOCNUM, "(OUTB) ")

                    CREDAT = diccionario.Item("CREDAT")
                    CRETIM = diccionario.Item("CRETIM")
                    diccionario.Clear()
                    diccionario = obtenerNodosHijosDePadre("E1EDL20", xmldoc)
                    VBELN = diccionario.Item("VBELN")
                    VSTEL = diccionario.Item("VSTEL")
                    diccionario.Clear()

                    diccionario = obtenerNodosHijosDePadre("E1ADRM1", xmldoc)
                    PARTNER_ID = diccionario.Item("PARTNER_ID")
                    NAME1 = diccionario.Item("NAME1")

                    conexion.ConnectionString = conexionString
                    conexion_pedidos.ConnectionString = conexionStringPedidos
                    Try
                        conexion.Open()
                        conexion_pedidos.Open()
                    Catch ex As Exception
                        escribirLog(ex.Message.ToString, "(OUTB) ")
                    End Try
                    myTrans = conexion.BeginTransaction()
                    myTransPedidos = conexion_pedidos.BeginTransaction()

                    sqlstring = ""
                    sqlstring = "DELETE FROM [CABECERA_OUTBOUND] WHERE DOCNUM='" & DOCNUM & "' "
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()

                    Dim id_pedido As Integer
                    id_pedido = obtenerIdPedido("D19")

                    sqlstring = ""
                    escribirLog("Se realizo la Conexion !!", "(OUTB) ")
                    sqlstring = "INSERT INTO [CABECERA_OUTBOUND] VALUES('" & DOCNUM & "','" & CREDAT & "','" & CRETIM & "','" & VBELN & "','" & VSTEL & "','" & PARTNER_ID & "','" & NAME1 & "','N','N')"
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()

                    
                    sqlstring = ""
                    sqlstring = "DELETE FROM [DETALLE_OUTBOUND] WHERE DOCNUM='" & DOCNUM & "' "
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()


                    Dim linea As Long
                    linea = 1000
                    listaDiccionario = obtenerNodosHijosDePadreLista("E1EDL24", xmldoc)
                    Dim contadorItem As Integer
                    Dim ix As Integer
                    contadorItem = 1
                    For ix = 0 To listaDiccionario.Count - 1

                        diccionarioDetalle = listaDiccionario.Item(ix)

                        POSNR = diccionarioDetalle.Item("POSNR")
                        MATNR = diccionarioDetalle.Item("MATNR")
                        ARKTX = diccionarioDetalle.Item("ARKTX")
                        MATKL = diccionarioDetalle.Item("MATKL")
                        WERKS = diccionarioDetalle.Item("WERKS")
                        LGORT = diccionarioDetalle.Item("LGORT")
                        LFIMG = diccionarioDetalle.Item("LFIMG")
                        VRKME = diccionarioDetalle.Item("VRKME")
                        LGMNG = diccionarioDetalle.Item("LGMNG")
                        MEINS = diccionarioDetalle.Item("MEINS")
                        NTGEW = diccionarioDetalle.Item("NTGEW")
                        GEWEI = diccionarioDetalle.Item("GEWEI")
                        VOLUM = diccionarioDetalle.Item("VOLUM")
                        VGBEL = diccionarioDetalle.Item("VGBEL")

                        If LGORT = "2000" Or LGORT = "2010" Then
                            Dim comandoAS400 As New Odbc.OdbcCommand
                            Dim cnnPO As New Odbc.OdbcConnection(cadenaAS400)
                            Dim sql As String
                            Dim fechaJuliana As String
                            Dim uom As String
                            Dim item As String
                            cnnPO.Open()

                            fechaJuliana = buscarFechaJuliana()

                            insertar = True
                            sqlstring = "INSERT INTO [DETALLE_OUTBOUND] VALUES('" & DOCNUM & "','" & POSNR & "','" & MATNR & "','" & ARKTX & "','" & MATKL & "','" & WERKS & "','" & LGORT & "','" & LFIMG & "','" & VRKME & "','" & LGMNG & "','" & MEINS & "','" & NTGEW & "','" & GEWEI & "','" & VOLUM & "','N','0')"
                            comando.Connection = conexion
                            comando.Transaction = myTrans
                            comando.CommandText = " "
                            comando.CommandText = sqlstring
                            If insertar Then
                                comando.ExecuteNonQuery()
                            Else
                                sqlstring = ""
                                sqlstring = "DELETE FROM [CABECERA_OUTBOUND] WHERE DOCNUM='" & DOCNUM & "' "
                                comando.Connection = conexion
                                comando.Transaction = myTrans
                                comando.CommandText = " "
                                comando.CommandText = sqlstring
                                comando.ExecuteNonQuery()
                                Exit For

                            End If
                            comando.Parameters.Clear()

                            comandoAS400.Connection = cnnPO
                            sql = "INSERT INTO  F4301Z1(SYEDUS    ,SYEDBT  ,SYEDTN,SYEDLN,SYTYTN     ,SYDRIN,SYKCOO ,SYDOCO   ,SYDCTO,SYMCU           ,SYOKCO ,SYAN8 ,SYSHAN  ,SYDRQJ,SYTRDJ,SYPDDJ,SYPTC ,SYEXR1,SYTXA1,SYATXT  ,SYANBY ,SYOTOT,SYAVCH,SYCORD,SYCRRM,SYCRCD,SYCRR,SYORBY,SYTNAC,SYVR01      ,SYVR02      ,SYSFXO,SYOPDJ,SYCRMD,SYFUF2,SYEDSP) VALUES " & _
                                                      "('EDELGADO','" & CLng(Right(DOCNUM, 8)) & "','0'   ," & CLng(POSNR) * 1000 & " ,'JDEINSAP' ,'1'   ,'00300',0        ,'OJ'  ,'" & buscarCampoItemF4104(CLng(MATNR), "IVDSC2") & "','00300',47301 ,30000099," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'1'   ,'V'   ,'IVA' ,''    ,11409  ,0,'N'   ,0     ,'D'   ,'VEF' ,0     ,'EDELGADO'  ,'02'  ,'" & VGBEL & "','" & CLng(DOCNUM) & "','000' ," & fechaJuliana & ",'2'   ,0     ,'N')"
                            comandoAS400.CommandText = Sql
                            comandoAS400.ExecuteNonQuery()

                            uom = buscarUOMF4101(CLng(MATNR))
                            item = buscarCampoItemF4104(CLng(MATNR), "IVLITM")
                            sql = "INSERT INTO  F4311Z1(SZEDUS    ,SZEDBT                 ,SZEDTN,SZEDLN                ,SZEDCT,SZTYTN     ,SZEDDT,SZDRIN,SZEDDL,SZEDSP,SZPNID,SZTNAC,SZKCOO ,SZDOCO,SZDCTO,SZLNID                ,SZMCU                                                ,SZCO   ,SZOKCO ,SZOORN                ,SZAN8  ,SZSHAN    ,SZDRQJ                      ,SZTRDJ                      ,SZPDDJ                      ,SZITM     ,SZLITM      ,SZAITM      ,SZDSC1                 ,SZLNTY,SZUOM,SZUORG    ,SZUOPN  ,SZTX,SZEXR1,SZTXA1,SZANBY ,SZPQOR    ,SZUOM2,SZSQOR    ,SZGLC ,SZVR01      ,SZVR02      ,SZSFXO,SZDGL ,SZNXTR,SZLTTR,SZCRMD,SZUOM1,SZAVCH,SZUNCD,SZCRCD) VALUES " & _
                                  "                    ('EDELGADO','" & CLng(Right(DOCNUM, 8)) & "' ,'0'   ," & CLng(POSNR) * 1000 & " ,''   ,'JDEINSAP'  ,'0'   ,'1'   ,'0'   ,'N'   ,''    ,'02'  ,'00300',0     ,'OJ'  ," & CLng(POSNR) * 1000 & "  ,'" & buscarCampoItemF4104(CLng(MATNR), "IVDSC2") & "','00300','00300','" & CLng(Right(DOCNUM, 8)) & "','47301','30000099'," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'" & buscarCampoItemF4104(CLng(MATNR), "IVITM") & "'   ,'" & item & "','" & item & "','" & buscarCampoItemF4104(CLng(MATNR), "IVDSC1") & "','C'  ,'" & uom & "' ," & (CDbl(LFIMG) / 1) * 10000 & "," & (CDbl(LFIMG) / 1) * 100 & ",'Y' ,'V'   ,'IVA' ,11409  ," & (CDbl(LFIMG) / 1) * 10000 & ",'" & uom & "'  ," & (CDbl(LFIMG) / 1) * 10000 & ",'ME00','" & VGBEL & "','" & CLng(DOCNUM) & "','000' ," & fechaJuliana & ",280   ,230   ,'2'   ,'" & uom & "'  ,'N'   ,'N'   ,'VEF')"
                            comandoAS400.CommandText = sql
                            comandoAS400.ExecuteNonQuery()

                            myTrans.Commit()
                            cnnPO.Close()

                        Else

                            If LGORT = "5000" Or LGORT = "5010" Then
                                Dim marca As String
                                Dim fechaJuliana As String
                                fechaJuliana = buscarFechaJuliana()

                                insertar = True
                                sqlstring = "INSERT INTO [DETALLE_OUTBOUND] VALUES('" & DOCNUM & "','" & POSNR & "','" & MATNR & "','" & ARKTX & "','" & MATKL & "','" & WERKS & "','" & LGORT & "','" & LFIMG & "','" & VRKME & "','" & LGMNG & "','" & MEINS & "','" & NTGEW & "','" & GEWEI & "','" & VOLUM & "','N','0')"
                                comando.Connection = conexion
                                comando.Transaction = myTrans
                                comando.CommandText = " "
                                comando.CommandText = sqlstring
                                If insertar Then
                                    comando.ExecuteNonQuery()
                                Else
                                    sqlstring = ""
                                    sqlstring = "DELETE FROM [CABECERA_OUTBOUND] WHERE DOCNUM='" & DOCNUM & "' "
                                    comando.Connection = conexion
                                    comando.Transaction = myTrans
                                    comando.CommandText = " "
                                    comando.CommandText = sqlstring
                                    comando.ExecuteNonQuery()
                                    Exit For

                                End If
                                comando.Parameters.Clear()

                                insertar = False
                                marca = buscarMarcaPt(CLng(MATNR))

                                If marca = "SM1" Or marca = "CAX" Then
                                    If WERKS = "LL2" Then
                                        insertar = True
                                    End If
                                Else
                                    If WERKS = "LL1" Then
                                        insertar = True
                                    End If

                                End If

                                diccionario.Clear()

                                Dim unidad As String

                                Dim cantidad As Double

                                'LL1 CACIQUE / LL2 SMIRNOFF
                                If WERKS = "LL1" Then
                                    SZEKCO = "00309"
                                    SZMCU = "    309A0005"
                                    SZAN8 = "30000099"
                                    ZFUSER = "DIAGEO"
                                    esDO = True
                                Else
                                    If WERKS = "LL2" Then
                                        SZEKCO = "00300"
                                        SZMCU = "    300A0005"
                                        'SZAN8 = "30001407"
                                        SZAN8 = buscarClientePt(buscarItemPt(CLng(MATNR)))
                                        ZFUSER = "DIAGEO"
                                        esDO = False
                                    End If

                                End If

                                unidad = buscarUOM(VRKME)

                                If unidad = "BT" Then
                                    unidad = "CA"
                                    cantidad = (CDbl(LFIMG) / 1) / buscarPackingSize(CLng(MATNR))
                                Else
                                    cantidad = CDbl(LFIMG) / 1
                                End If

                                sqlstring = ""
                                sqlstring = "INSERT INTO [dusa_pedidos].[dbo].[orders_details]([order_id],[product_id],[unit],[qty],[subtotal],[customer_id],[warehouse],[required_date]) VALUES ('" & id_pedido & "','" & buscarItemPt(CLng(MATNR)) & "','" & unidad & "'," & Replace(cantidad.ToString, ",", ".") & ",'1','" & SZAN8 & "','" & buscarAlmacenPt(CLng(MATNR)) & "','" & obtenerFecha() & "')"
                                comando.Connection = conexion_pedidos
                                comando.Transaction = myTransPedidos
                                comando.CommandText = " "
                                comando.CommandText = sqlstring
                                comando.ExecuteNonQuery()
                                comando.Parameters.Clear()

                                Dim cnn3 As New Odbc.OdbcConnection(cadenaAS400)

                                Dim rs3 As New Odbc.OdbcCommand("INSERT INTO F4011Z(SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZKCOO,SZDOCO,SZDCTO,SZLNID,SZMCU,SZAN8,SZSHAN,SZDRQJ,SZLITM,SZUOM,SZUORG,SZTORG,SZUSER,SZPID,SZJOBN,SZUPMJ,SZTDAY, SZVR01,   SZEDTY,SZEDSQ,SZEDST,SZEDFT,SZEDDT,SZEDER,SZEDDL,SZEDSP,SZPNID,SZSFXO,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO,SZRKCO,SZRORN,SZRCTO,SZRLLN,SZDMCT,SZDMCS,SZBALU,SZPA8,SZTRDJ,SZPDDJ,SZOPDJ,SZADDJ,SZIVD,SZCNDJ,SZDGL,SZRSDJ,SZPEFJ,SZPPDJ,SZPSDJ,SZVR02,SZITM,SZAITM,SZLOCN,SZLOTN,SZFRGD,SZTHGD,SZFRMP,SZTHRP,SZEXDP,SZDSC1,SZDSC2,SZLNTY,SZNXTR,SZLTTR,SZEMCU,SZRLIT,SZKTLN,SZCPNT,SZRKIT,SZKTP,SZSRP1,SZSRP2,SZSRP3,SZSRP4,SZSRP5,SZPRP1,SZPRP2,SZPRP3,SZPRP4,SZPRP5,SZSOQS,SZSOBK,SZSOCN,SZSONE,SZUOPN,SZQTYT,SZQRLV,SZCOMM,SZOTQY,SZUPRC,SZAEXP,SZAOPN,SZPROV,SZTPC,SZAPUM,SZLPRC,SZUNCS,SZECST,SZCSTO,SZTCST,SZINMG,SZPTC,SZRYIN,SZDTBS,SZTRDC,SZFUN2,SZASN,SZPRGR,SZCLVL,SZDSPR,SZDSFT,SZFAPP,SZCADC,SZKCO,SZDOC,SZDCT,SZODOC,SZODCT,SZOKC,SZPSN,SZDELN,SZTAX1,SZTXA1,SZEXR1,SZATXT,SZPRIO,SZRESL,SZBACK,SZSBAL,SZAPTS,SZLOB,SZEUSE,SZDTYS,SZNTR,SZVEND,SZANBY,SZCARS,SZMOT,SZCOT,SZROUT,SZSTOP,SZZON,SZCNID,SZFRTH,SZAFT,SZFUF1,SZFRTC,SZFRAT,SZRATT,SZSHCM,SZSHCN,SZSERN,SZUOM1,SZPQOR,SZUOM2,SZSQOR,SZUOM4,SZITWT,SZWTUM,SZITVL,SZVLUM,SZRPRC,SZORPR,SZORP,SZCMGP,SZCMGL,SZGLC,SZCTRY,SZFY,SZSTTS,SZSO01,SZSO02,SZSO03,SZSO04,SZSO05,SZSO06,SZSO07,SZSO08,SZSO09,SZSO10,SZSO11,SZSO12,SZSO13,SZSO14,SZSO15,SZACOM,SZCMCG,SZRCD,SZGRWT,SZGWUM,SZANI,SZAID,SZOMCU,SZOBJ,SZSUB,SZLT,SZSBL,SZSBLT,SZLCOD,SZUPC1,SZUPC2,SZUPC3,SZSWMS,SZUNCD,SZCRMD,SZCRCD,SZCRR,SZFPRC,SZFUP,SZFEA,SZFUC,SZFEC,SZURCD,SZURDT,SZURAT,SZURAB,SZURRF,SZIR01,SZIR02,SZIR03,SZIR04,SZIR05,SZSOOR,SZDEID,SZPSIG,SZRLNU,SZPMDT,SZRLTM,SZRLDJ,SZDRQT,SZADTM,SZOPTT,SZPDTT,SZPSTM,SZPMTN,SZBSC,SZCBSC,SZDVAN,SZRFRV,SZSHPN,SZPRJM,SZHOLD,SZPMTO,SZDUAL) VALUES ('" & SZEKCO & "'," & id_pedido & ",'SY'," & linea & ",'" & SZEKCO & "'," & id_pedido & ",'SY'," & linea & ",'" & SZMCU & "','" & SZAN8 & "','" & SZAN8 & "'," & buscarFechaJuliana() & ",'" & buscarItemPt(CLng(MATNR)) & "','" & unidad & "'," & cantidad * 10000 & ",'DIAGEO','DIAGEO','F55WEBDBT','S102F350'," & buscarFechaJuliana() & " ," & buscarFechaJuliana() & ",'" & CLng(Right(DOCNUM, 8)) & "','',0,'','',0,'',0,'','','','','','','',0,'','','',0,'',0,'',0,0,0,0,0,0,0,0,0,0,0,0,'',0,'','','','','',0,0,0,'','','',0,0,'','',0,0,0,0,'','','','','','','','','','',0,0,0,0,0,0,0,'','',0,0,0,'','','',0,0,0,'',0,'','','','',0,0,'','','',0,'','',0,'',0,'',0,'','',0,0,'','','','','','','','','','','','','',0,0,0,'','','','','','','','','','','','','','','','',0,'',0,'',0,'',0,'','','','','','','',0,0,'','','','','','','','','','','','','','','','','','','',0,'','','','','','','','','','','','','','','','','',0,0,0,0,0,0,'',0,0,0,'','','','','','',0,0,'','',0,0,0,0,0,0,0,0,'','','',0,'',0,0,'','','')  ", cnn3)
                                Dim reader3 As Odbc.OdbcDataReader

                                cnn3.Open()
                                reader3 = rs3.ExecuteReader
                                reader3.Close()
                                cnn3.Close()

                                linea = linea + 1000


                            End If

                        End If

                        

                    Next

                    If LGORT = "5000" Or LGORT = "5010" Then

                        sqlstring = ""
                        sqlstring = "INSERT INTO [dusa_pedidos].[dbo].[orders]([order_id],[order_date],[order_time],[salesman_id],[customer_id],[delivery_date],[status],[amount],[comments],[required_date],[status_changed_at],[status_changed_on])     VALUES('" & id_pedido & "','" & obtenerFecha() & "','" & obtenerHoraPedido() & "','D19','" & SZAN8 & "','" & obtenerFecha() & "','PRO','1','" & CLng(DOCNUM) & "','" & obtenerFecha() & "','','')"
                        comando_pedidos.Connection = conexion_pedidos
                        comando_pedidos.Transaction = myTransPedidos
                        comando_pedidos.CommandText = " "
                        comando_pedidos.CommandText = sqlstring
                        If insertar Then
                            comando_pedidos.ExecuteNonQuery()

                            Dim cnn4 As New Odbc.OdbcConnection(cadenaAS400)

                            Dim rs4 As New Odbc.OdbcCommand("INSERT INTO F4001Z(SYEKCO,SYEDOC,SYEDCT,SYKCOO,SYDOCO,SYDCTO,SYMCU,SYAN8,SYSHAN,SYDRQJ,SYTRDJ,SYUSER,SYPID,SYJOBN,SYUPMJ,SYTDAY,SYVR01,SYEDTY,SYEDSQ,SYEDLN,SYEDST,SYEDFT,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYPNID,SYOFRQ,SYNXDJ,SYSSDJ,SYSUN,SYMON,SYTUE,SYWED,SYTHR,SYFRI,SYSAT,SYSFXO,SYCO,SYOKCO,SYOORN,SYOCTO,SYRKCO,SYRORN,SYRCTO,SYPA8,SYPDDJ,SYOPDJ,SYADDJ,SYCNDJ,SYPEFJ,SYPPDJ,SYPSDJ,SYVR02,SYDEL1,SYDEL2,SYINMG,SYPTC,SYRYIN,SYASN,SYPRGP,SYTRDC,SYPCRT,SYTXA1,SYEXR1,SYTXCT,SYATXT,SYPRIO,SYBACK,SYSBAL,SYHOLD,SYPLST,SYINVC,SYNTR,SYANBY,SYCARS,SYMOT,SYCOT,SYROUT,SYSTOP,SYZON,SYCNID,SYFRTH,SYAFT,SYFUF1,SYFRTC,SYMORD,SYRCD,SYFUF2,SYOTOT,SYTOTC,SYWUMD,SYVUMD,SYAUTN,SYCACT,SYCEXP,SYSBLI,SYCRMD,SYCRRM,SYCRCD,SYCRR,SYLNGP,SYFAP,SYFCST,SYORBY,SYTKBY,SYURCD,SYURDT,SYURAT,SYURAB,SYURRF,SYIR01,SYIR02,SYIR03,SYIR04,SYIR05,SYVR03,SYSOOR,SYPMDT,SYRSDT,SYRQSJ,SYPSTM,SYPDTT,SYOPTT,SYDRQT,SYADTM,SYADLJ,SYPBAN,SYITAN,SYFTAN,SYDVAN,SYDOC1,SYDCT4,SYCORD,SYBSC,SYBCRC,SYAUFT,SYAUFI,SYOPBO,SYOPTC,SYOPLD,SYOPBK,SYOPSB,SYOPPS,SYOPPL,SYOPMS,SYOPSS,SYOPBA,SYOPLL) VALUES ('" & SZEKCO & "'," & id_pedido & ",'SY','" & SZEKCO & "'," & id_pedido & ",'SY','" & SZMCU & "','" & SZAN8 & "','" & SZAN8 & "'," & buscarFechaJuliana() & "," & buscarFechaJuliana() & ",'DIAGEO','F55WEBHBT','S102F350'," & buscarFechaJuliana() & "," & buscarFechaJuliana() & ",'" & CLng(Right(DOCNUM, 8)) & "','',0,0,'','',0,'',0,'','','',0,0,'','','','','','','','','','','','','','','',0,0,0,0,0,0,0,0,'','','','','','','','',0,0,'','','','','','','','','',0,'',0,0,'','','','','','','','','','','','','',0,0,'','','','',0,'','','','',0,'',0,0,'','','',0,0,0,'','','','','','','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'',0,'','','','','',0,0,0,0,'','','','','','')  ", cnn4)
                            Dim reader4 As Odbc.OdbcDataReader

                            cnn4.Open()
                            reader4 = rs4.ExecuteReader
                            reader4.Close()
                            cnn4.Close()

                            Dim tipoPedido As String

                            If esDO Then
                                tipoPedido = ""
                            Else
                                tipoPedido = "SM1"
                            End If

                            '
                            Dim cnn5 As New Odbc.OdbcConnection(cadenaAS400)

                            Dim rs5 As New Odbc.OdbcCommand("INSERT INTO F0041Z1(ZFCTID,ZFUSER,ZFTRNM,ZFTRNY,ZFTRNK,ZFPID,ZFSERK,ZFAPVC,ZFTSC,ZFCNF,ZFUPTY) VALUES ('S102F350','" & ZFUSER & "'," & buscarMaxF0041Z1() & ",'ZY','" & id_pedido & "SY" & SZEKCO & "','" & tipoPedido & "',0,'A','1','','')", cnn5)
                            Dim reader5 As Odbc.OdbcDataReader

                            cnn5.Open()
                            reader5 = rs5.ExecuteReader
                            reader5.Close()
                            cnn5.Close()

                            esDO = False


                        End If
                        comando_pedidos.Parameters.Clear()

                        myTrans.Commit()
                        myTransPedidos.Commit()

                    End If


                    If File.Exists(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar) Then
                        File.Delete(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar)
                    End If

                    fs.Close()
                    File.Copy(nombreArchivoProcesar, Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar)
                    File.Delete(nombreArchivoProcesar)

                    escribirLog("FINALIZADA IMPORTACION DE DATOS DE DOCUMENTO:" & DOCNUM, "(OUTB) ")
                    conexion.Close()
                    conexion_pedidos.Close()

                    End If
            Catch ex As Exception
                'myTrans.Rollback()
                escribirLog(ex.Message.ToString, "(OUTB) ")
            End Try

        Next

    End Sub



    Private Sub procesos_STO()

        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String


        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()

            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_STO.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

            procesar_STO_ADVICE()
            'crearGoodReceiptsSTO()

        Catch oe As Exception
            escribirLog(oe.Message.ToString, "(STO) ")
        Finally

            logger.Close()
        End Try

    End Sub

    Private Sub procesar_STO_ADVICE()

        Dim DOCNUM As String
        Dim CREDAT As String
        Dim CRETIM As String
        Dim VBELN As String
        Dim VSTEL As String
        Dim POSNR As String
        Dim MATNR As String
        Dim ARKTX As String
        Dim MATKL As String
        Dim WERKS As String
        Dim CHARG As String
        Dim LGORT As String
        Dim LFIMG As String
        Dim VRKME As String
        Dim NTGEW As String
        Dim GEWEI As String
        Dim VOLUM As String
        Dim PARTNER_ID As String
        Dim VGBEL As String
        Dim VGPOS As String

        Dim conexion As New SqlConnection
        Dim myTrans As SqlTransaction
        Dim comando As New SqlClient.SqlCommand
        Dim sqlstring As String

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim listaNodos As New List(Of Nodo)
        Dim nombreArchivoProcesar As String
        Dim listaArchivosProcesar As New List(Of String)
        Dim listaDiccionario As New List(Of Dictionary(Of String, String))
        Dim diccionarioDetalle As New Dictionary(Of String, String)
        Dim iw As Integer

        listaArchivosProcesar = obtenerNombreArchivo(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", ""), prefijo)
        For iw = 0 To listaArchivosProcesar.Count - 1

            Try

                ' nombreArchivoProcesar = obtenerNombreArchivo(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", ""), prefijo)
                nombreArchivoProcesar = listaArchivosProcesar.Item(iw)
                If nombreArchivoProcesar.Trim.CompareTo("") = 0 Then
                    escribirLog("No se encuentra el archivo necesitado con prefijo " & prefijo & " y extension .xml !", "(STO) ")
                Else

                    Dim fs As New FileStream(nombreArchivoProcesar, FileMode.Open, FileAccess.Read)
                    xmldoc = New XmlDataDocument()
                    xmldoc.Load(fs)
                    diccionario = obtenerNodosHijosDePadre("EDI_DC40", xmldoc)
                    DOCNUM = diccionario.Item("DOCNUM")

                    escribirLog("INICIO DE IMPORTACION DE DATOS DE DOCUMENTO:" & DOCNUM, "(STO) ")

                    CREDAT = diccionario.Item("CREDAT")
                    CRETIM = diccionario.Item("CRETIM")
                    diccionario.Clear()
                    diccionario = obtenerNodosHijosDePadre("E1EDL20", xmldoc)
                    VBELN = diccionario.Item("VBELN")
                    VSTEL = diccionario.Item("VSTEL")
                    diccionario.Clear()

                    diccionario = obtenerNodosHijosDePadre("E1ADRM1", xmldoc)
                    PARTNER_ID = diccionario.Item("PARTNER_ID")


                    conexion.ConnectionString = conexionString
                    conexion.Open()

                    sqlstring = ""
                    sqlstring = "DELETE FROM [CABECERA_STO_ADVICE] WHERE DOCNUM='" & DOCNUM & "' "
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()



                    myTrans = conexion.BeginTransaction()
                    sqlstring = ""
                    sqlstring = "INSERT INTO [CABECERA_STO_ADVICE] VALUES('" & DOCNUM & "','" & CREDAT & "','" & CRETIM & "','" & VBELN & "','" & VSTEL & "','" & PARTNER_ID & "','N')"
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()


                    sqlstring = ""
                    sqlstring = "DELETE FROM [DETALLE_STO_ADVICE] WHERE DOCNUM='" & DOCNUM & "' "
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()

                    Dim linea As Integer
                    Dim comandoAS400 As New Odbc.OdbcCommand
                    Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
                    Dim fechaJuliana As String
                    Dim uom As String
                    Dim item As String
                    Dim sql As String

                    cnn.Open()
                    linea = 1

                    fechaJuliana = buscarFechaJuliana()
                    comandoAS400.Connection = cnn




                    listaDiccionario = obtenerNodosHijosDePadreLista("E1EDL24", xmldoc)
                    Dim contadorItem As Integer
                    Dim ix As Integer
                    contadorItem = 1
                    For ix = 0 To listaDiccionario.Count - 1

                        diccionarioDetalle = listaDiccionario.Item(ix)


                        'diccionario = obtenerNodosHijosDePadre("E1EDL24", xmldoc)
                        POSNR = diccionarioDetalle.Item("POSNR")
                        MATNR = diccionarioDetalle.Item("MATNR")
                        ARKTX = diccionarioDetalle.Item("ARKTX")
                        MATKL = diccionarioDetalle.Item("MATKL")
                        WERKS = diccionarioDetalle.Item("WERKS")
                        'CHARG = diccionario.Item("CHARG")
                        LFIMG = diccionarioDetalle.Item("LFIMG")
                        VRKME = diccionarioDetalle.Item("VRKME")
                        NTGEW = diccionarioDetalle.Item("NTGEW")
                        GEWEI = diccionarioDetalle.Item("GEWEI")
                        LGORT = diccionarioDetalle.Item("LGORT")
                        VGBEL = diccionarioDetalle.Item("VGBEL")
                        VGPOS = diccionarioDetalle.Item("VGPOS")

                        diccionario.Clear()


                        If linea = 1 Then
                            'sql = "INSERT INTO  F4301Z1(SYEDUS    ,SYEDBT  ,SYEDTN,SYEDLN,SYTYTN     ,SYDRIN,SYKCOO ,SYDOCO   ,SYDCTO,SYMCU           ,SYOKCO ,SYAN8 ,SYSHAN  ,SYDRQJ,SYTRDJ,SYPDDJ,SYPTC ,SYEXR1,SYTXA1,SYATXT  ,SYANBY ,SYOTOT,SYAVCH,SYCORD,SYCRRM,SYCRCD,SYCRR,SYORBY,SYTNAC,SYVR01      ,SYVR02      ,SYSFXO,SYOPDJ,SYCRMD,SYFUF2,SYEDSP) VALUES " & _
                            '                  "('EDELGADO','" & Clng(DOCNUM) & "','0'   ," & linea * 1000 & " ,'JDEINSAP' ,'1'   ,'00300',0        ,'O5'  ,'" & buscarCampoItemF4104(Clng(MATNR), "IVDSC2") & "','00300',47301 ,30000099," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'1'   ,'V'   ,'IVA' ,''    ,11409  ,0,'N'   ,0     ,'D'   ,'VEF' ,0     ,'EDELGADO'  ,'02'  ,'" & BELNR & "','" & Clng(DOCNUM) & "','000' ," & fechaJuliana & ",'2'   ,0     ,'N')"

                            ' STO
                            sql = "INSERT INTO  F4301Z1(SYEDUS    ,SYEDBT  ,SYEDTN,SYEDLN,SYTYTN     ,SYDRIN,SYKCOO ,SYDOCO   ,SYDCTO,SYMCU           ,SYOKCO ,SYAN8 ,SYSHAN  ,SYDRQJ,SYTRDJ,SYPDDJ,SYPTC ,SYEXR1,SYTXA1,SYATXT  ,SYANBY ,SYOTOT,SYAVCH,SYCORD,SYCRRM,SYCRCD,SYCRR,SYORBY,SYTNAC,SYVR01      ,SYVR02      ,SYSFXO,SYOPDJ,SYCRMD,SYFUF2,SYEDSP) VALUES " & _
                                   "('EDELGADO','" & CLng(Right(DOCNUM, 8)) & "','0'   ," & linea * 1000 & " ,'JDEINSAP' ,'1'   ,'00300',0        ,'O7'  ,'" & buscarCampoItemF4104(CLng(MATNR), "IVDSC2") & "'  ,'00300',47301 ,30000099," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'1'   ,'V'   ,'IVA' ,''    ,11409  ,0,'N'   ,0     ,'D'   ,'VEF' ,0     ,'EDELGADO'  ,'02'  ,'" & VBELN & "','" & CLng(Right(DOCNUM, 8)) & "','000' ," & fechaJuliana & ",'2'   ,0     ,'N')"

                            escribirLog(sql, "(OUTB) ")

                            comandoAS400.CommandText = sql
                            comandoAS400.ExecuteNonQuery()
                        End If



                        sqlstring = "INSERT INTO [DETALLE_STO_ADVICE] VALUES('" & DOCNUM & "','" & POSNR & "','" & MATNR & "','" & ARKTX & "','" & MATKL & "','" & WERKS & "','" & LGORT & "','" & CHARG & "','" & LFIMG & "','" & VRKME & "','" & NTGEW & "','" & GEWEI & "','" & VOLUM & "','" & VGBEL & "','" & VGPOS & "')"
                        comando.Connection = conexion
                        comando.Transaction = myTrans
                        comando.CommandText = " "
                        comando.CommandText = sqlstring
                        comando.ExecuteNonQuery()
                        comando.Parameters.Clear()

                        uom = buscarUOMF4101(CLng(MATNR))
                        item = buscarCampoItemF4104(CLng(MATNR), "IVLITM")
                        'sql = "INSERT INTO  F4311Z1(SZEDUS    ,SZEDBT                 ,SZEDTN,SZEDLN                ,SZEDCT,SZTYTN     ,SZEDDT,SZDRIN,SZEDDL,SZEDSP,SZPNID,SZTNAC,SZKCOO ,SZDOCO,SZDCTO,SZLNID                ,SZMCU                                                ,SZCO   ,SZOKCO ,SZOORN                ,SZAN8  ,SZSHAN    ,SZDRQJ                      ,SZTRDJ                      ,SZPDDJ                      ,SZITM     ,SZLITM      ,SZAITM      ,SZDSC1                 ,SZLNTY,SZUOM,SZUORG    ,SZUOPN  ,SZTX,SZEXR1,SZTXA1,SZANBY ,SZPQOR    ,SZUOM2,SZSQOR    ,SZGLC ,SZVR01      ,SZVR02      ,SZSFXO,SZDGL ,SZNXTR,SZLTTR,SZCRMD,SZUOM1,SZAVCH,SZUNCD,SZCRCD) VALUES " & _
                        '      "                    ('EDELGADO','" & Clng(DOCNUM) & "' ,'0'   ," & linea * 1000 & " ,''   ,'JDEINSAP'  ,'0'   ,'1'   ,'0'   ,'N'   ,''    ,'02'  ,'00300',0     ,'O5'  ," & linea * 1000 & "  ,'" & buscarCampoItemF4104(Clng(IDTNR), "IVDSC2") & "','00300','00300','" & Clng(DOCNUM) & "','47301','30000099'," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'" & buscarCampoItemF4104(Clng(IDTNR), "IVITM") & "'   ,'" & item & "','" & item & "','" & buscarCampoItemF4104(Clng(IDTNR), "IVDSC1") & "','S5'  ,'" & uom & "' ," & (CDbl(MENGE) / 1000) * 10000 & "," & (CDbl(MENGE) / 1000) * 100 & ",'Y' ,'V'   ,'IVA' ,11409  ," & (CDbl(MENGE) / 1000) * 10000 & ",'" & uom & "'  ," & (CDbl(MENGE) / 1000) * 10000 & ",'CS08','" & BELNR & "','" & Clng(DOCNUM) & "','000' ," & fechaJuliana & ",280   ,230   ,'2'   ,'" & uom & "'  ,'N'   ,'N'   ,'VEF')"

                        ' STO
                        sql = "INSERT INTO  F4311Z1(SZEDUS    ,SZEDBT   ,SZEDTN,SZEDLN,SZEDCT,SZTYTN     ,SZEDDT,SZDRIN,SZEDDL,SZEDSP,SZPNID,SZTNAC,SZKCOO ,SZDOCO,SZDCTO,SZLNID,SZMCU         ,SZCO   ,SZOKCO ,SZOORN  ,SZAN8  ,SZSHAN    ,SZDRQJ,SZTRDJ,SZPDDJ,SZITM     ,SZLITM      ,SZAITM      ,SZDSC1                 ,SZLNTY,SZUOM,SZUORG    ,SZUOPN  ,SZTX,SZEXR1,SZTXA1,SZANBY ,SZPQOR    ,SZUOM2,SZSQOR    ,SZGLC ,SZVR01      ,SZVR02      ,SZSFXO,SZDGL ,SZNXTR,SZLTTR,SZCRMD,SZUOM1,SZAVCH,SZUNCD,SZCRCD) VALUES " & _
                        "                          ('EDELGADO','" & CLng(Right(DOCNUM, 8)) & "' ,'0'   ," & linea * 1000 & " ,''   ,'JDEINSAP'  ,'0'   ,'1'   ,'0'   ,'N'   ,''    ,'02'  ,'00300',0     ,'O7'  ," & linea * 1000 & "  ,'" & buscarCampoItemF4104(CLng(MATNR), "IVDSC2") & "','00300','00300','" & CLng(Right(DOCNUM, 8)) & "','47301','30000099'," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'" & buscarCampoItemF4104(CLng(MATNR), "IVITM") & "'   ,'" & item & "','" & item & "','" & buscarCampoItemF4104(CLng(MATNR), "IVDSC1") & "','S5'  ,'" & uom & "' ," & (CDbl(LFIMG) / 1) * 10000 & "," & (CDbl(LFIMG) / 1) * 100 & ",'Y' ,'V'   ,'IVA' ,11409  ," & (CDbl(LFIMG) / 1) * 10000 & ",'" & uom & "'  ," & (CDbl(LFIMG) / 1) * 10000 & ",'CS08','" & VBELN & "','" & CLng(Right(DOCNUM, 8)) & "','000' ," & fechaJuliana & ",280   ,230   ,'2'   ,'" & uom & "'  ,'N'   ,'N'   ,'VEF')"

                        escribirLog(sql, "(OUTB) ")

                        comandoAS400.CommandText = sql
                        comandoAS400.ExecuteNonQuery()
                        linea = linea + 1

                    Next

                    cnn.Close()
                    myTrans.Commit()

                    If File.Exists(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar) Then
                        File.Delete(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar)
                    End If

                    fs.Close()
                    File.Copy(nombreArchivoProcesar, Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar)
                    File.Delete(nombreArchivoProcesar)

                    escribirLog("FINALIZADA IMPORTACION DE DATOS DE DOCUMENTO:" & DOCNUM, "(STO) ")
                    conexion.Close()

                End If
            Catch ex As Exception
                'myTrans.Rollback()
                escribirLog(ex.Message.ToString, "(PO) ")
            End Try

        Next

    End Sub


    Private Function actualizarGenerado(ByVal tabla As String, ByVal docnum As String) As Integer

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim sqlstring As String

        Try

            conEx2.ConnectionString = conexionString
            conEx2.Open()

            sqlstring = ""
            sqlstring = "UPDATE " & tabla & " SET GENERADO='S' WHERE DOCNUM='" & docnum & "' "
            comando.Connection = conEx2
            comando.CommandText = " "
            comando.CommandText = sqlstring
            comando.ExecuteNonQuery()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try



        Return valor
    End Function

    Private Sub procesos_PO()

        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand


        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_PO.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

            procesar_PO_ADVICE()
            'crearGoodReceiptsPO()


        Catch oe As Exception
            escribirLog(oe.Message.ToString, "(PO) ")
        Finally

            logger.Close()
        End Try


    End Sub


    Private Sub procesar_PO_ADVICE()

        Dim RECIPNT_NO As String
        Dim BELNR As String
        Dim DOCNUM As String
        Dim CREDAT As String
        Dim CRETIM As String
        Dim POSEX As String
        Dim ACTION1 As String
        Dim MENGE As String
        Dim MENEE As String
        Dim NETWR As String
        Dim GEWEI As String
        Dim WERKS As String
        Dim LGORT As String
        Dim IDTNR As String
        Dim KTEXT As String
        Dim EDATU As String
        Dim NAME1 As String

        Dim conexion As New SqlConnection
        Dim myTrans As SqlTransaction
        Dim comando As New SqlClient.SqlCommand
        Dim sqlstring As String

        Dim diccionario As New Dictionary(Of String, String)
        Dim diccionarioDetalle As New Dictionary(Of String, String)
        Dim diccionarioItem As New Dictionary(Of String, String)
        Dim diccionarioArchivos As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim listaNodos As New List(Of Nodo)
        Dim nombreArchivoProcesar As String
        Dim listaDiccionario As New List(Of Dictionary(Of String, String))
        Dim listaDiccionarioItem As New List(Of Dictionary(Of String, String))
        Dim listaArchivosProcesar As New List(Of String)

        Dim iw As Integer

        listaArchivosProcesar = obtenerNombreArchivo(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", ""), prefijo)
        For iw = 0 To listaArchivosProcesar.Count - 1

            Try
                nombreArchivoProcesar = listaArchivosProcesar.Item(iw)
                If nombreArchivoProcesar.Trim.CompareTo("") = 0 Then
                    escribirLog("No se encuentra el archivo necesitado con prefijo " & prefijo & " y extension .xml !", "(PO) ")
                Else
                    Dim fs As New FileStream(nombreArchivoProcesar, FileMode.Open, FileAccess.Read)

                    xmldoc = New XmlDataDocument()
                    xmldoc.Load(fs)
                    diccionario = obtenerNodosHijosDePadre("EDI_DC40", xmldoc)
                    DOCNUM = diccionario.Item("DOCNUM")

                    escribirLog("INICIO DE IMPORTACION DE DATOS DE DOCUMENTO:" & DOCNUM, "(PO) ")

                    CREDAT = diccionario.Item("CREDAT")
                    CRETIM = diccionario.Item("CRETIM")
                    diccionario.Clear()
                    diccionario = obtenerNodosHijosDePadre("E1EDK01", xmldoc)
                    BELNR = diccionario.Item("BELNR")
                    RECIPNT_NO = diccionario.Item("RECIPNT_NO")
                    diccionario.Clear()

                    conexion.ConnectionString = conexionString
                    conexion.Open()
                    myTrans = conexion.BeginTransaction()

                    sqlstring = ""
                    sqlstring = "DELETE FROM [CABECERA_PO_ADVICE] WHERE DOCNUM='" & DOCNUM & "' "
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()

                    sqlstring = ""
                    sqlstring = "INSERT INTO [CABECERA_PO_ADVICE] VALUES('" & DOCNUM & "','" & CREDAT & "','" & CRETIM & "','" & BELNR & "','" & RECIPNT_NO & "','N')"
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()

                    sqlstring = "DELETE FROM [DETALLE_PO_ADVICE] WHERE DOCNUM='" & DOCNUM & "' "
                    comando.Connection = conexion
                    comando.Transaction = myTrans
                    comando.CommandText = " "
                    comando.CommandText = sqlstring
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()


                    Dim linea As Integer
                    Dim comandoAS400 As New Odbc.OdbcCommand
                    Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
                    Dim fechaJuliana As String
                    Dim uom As String
                    Dim item As String
                    Dim sql As String

                    cnn.Open()
                    linea = 1

                    fechaJuliana = buscarFechaJuliana()


                    listaDiccionario = obtenerNodosHijosDePadreLista("E1EDP01", xmldoc)
                    Dim contadorItem As Integer
                    Dim ix As Integer
                    contadorItem = 1
                    For ix = 0 To listaDiccionario.Count - 1

                        diccionarioDetalle = listaDiccionario.Item(ix)

                        POSEX = diccionarioDetalle.Item("POSEX")
                        ACTION1 = diccionarioDetalle.Item("ACTION")
                        MENGE = diccionarioDetalle.Item("MENGE")
                        MENEE = diccionarioDetalle.Item("MENEE")
                        NETWR = diccionarioDetalle.Item("NETWR")

                        Try
                            GEWEI = diccionarioDetalle.Item("GEWEI")
                        Catch ex As Exception
                            GEWEI = "NA"
                        End Try


                        WERKS = diccionarioDetalle.Item("WERKS")

                        Try
                            LGORT = diccionarioDetalle.Item("LGORT")
                        Catch ex As Exception
                            LGORT = "0"
                        End Try

                        diccionario.Clear()

                        listaDiccionarioItem = obtenerNodosHijosDePadreLista("E1EDP19", xmldoc)
                        Dim iy As Integer
                        Dim contadorInternoItem As Integer
                        contadorInternoItem = 0

                        For iy = 0 To listaDiccionarioItem.Count - 1
                            diccionarioItem = listaDiccionarioItem.Item(iy)

                            If diccionarioItem.Item("QUALF") = "001" Then
                                contadorInternoItem = contadorInternoItem + 1
                                If contadorItem = contadorInternoItem Then


                                    Try
                                        IDTNR = diccionarioItem.Item("IDTNR")
                                    Catch ex As Exception
                                        IDTNR = "0"
                                    End Try

                                    Try
                                        KTEXT = diccionarioItem.Item("KTEXT")
                                    Catch ex As Exception
                                        KTEXT = "NA"
                                    End Try
                                    contadorItem = contadorItem + 1
                                    Exit For
                                End If
                            End If

                        Next

                        If linea = 1 Then
                            comandoAS400.Connection = cnn
                            sql = "INSERT INTO  F4301Z1(SYEDUS    ,SYEDBT  ,SYEDTN,SYEDLN,SYTYTN     ,SYDRIN,SYKCOO ,SYDOCO   ,SYDCTO,SYMCU           ,SYOKCO ,SYAN8 ,SYSHAN  ,SYDRQJ,SYTRDJ,SYPDDJ,SYPTC ,SYEXR1,SYTXA1,SYATXT  ,SYANBY ,SYOTOT,SYAVCH,SYCORD,SYCRRM,SYCRCD,SYCRR,SYORBY,SYTNAC,SYVR01      ,SYVR02      ,SYSFXO,SYOPDJ,SYCRMD,SYFUF2,SYEDSP) VALUES " & _
                                                      "('EDELGADO','" & CLng(Right(DOCNUM, 8)) & "','0'   ," & linea * 1000 & " ,'JDEINSAP' ,'1'   ,'00300',0        ,'O5'  ,'" & buscarCampoItemF4104(CLng(IDTNR), "IVDSC2") & "','00300',47301 ,30000099," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'1'   ,'V'   ,'IVA' ,''    ,11409  ,0,'N'   ,0     ,'D'   ,'VEF' ,0     ,'EDELGADO'  ,'02'  ,'" & BELNR & "','" & CLng(Right(DOCNUM, 8)) & "','000' ," & fechaJuliana & ",'2'   ,0     ,'N')"

                            comandoAS400.CommandText = sql
                            comandoAS400.ExecuteNonQuery()
                        End If

                        diccionario = obtenerNodosHijosDePadre("E1EDP20", xmldoc)
                        EDATU = diccionario.Item("EDATU")

                        Dim nodName1 As Nodo
                        nodName1 = buscarNodo("NAME1", obtenerListaNodosHijosDePadre("E1EDKA1", xmldoc))
                        If nodName1.sName <> "NULL" Then
                            NAME1 = nodName1.sInner
                        End If

                        sqlstring = "INSERT INTO [DETALLE_PO_ADVICE] VALUES('" & DOCNUM & "','" & POSEX & "','" & ACTION1 & "','" & MENGE & "','" & MENEE & "','" & NETWR & "','" & GEWEI & "','" & WERKS & "','" & IDTNR & "','" & KTEXT & "','" & EDATU & "','" & NAME1 & "','" & LGORT & "')"
                        comando.Connection = conexion
                        comando.Transaction = myTrans
                        comando.CommandText = " "
                        comando.CommandText = sqlstring
                        comando.ExecuteNonQuery()
                        comando.Parameters.Clear()

                        uom = buscarUOMF4101(CLng(IDTNR))
                        item = buscarCampoItemF4104(CLng(IDTNR), "IVLITM")
                        sql = "INSERT INTO  F4311Z1(SZEDUS    ,SZEDBT                 ,SZEDTN,SZEDLN                ,SZEDCT,SZTYTN     ,SZEDDT,SZDRIN,SZEDDL,SZEDSP,SZPNID,SZTNAC,SZKCOO ,SZDOCO,SZDCTO,SZLNID                ,SZMCU                                                ,SZCO   ,SZOKCO ,SZOORN                ,SZAN8  ,SZSHAN    ,SZDRQJ                      ,SZTRDJ                      ,SZPDDJ                      ,SZITM     ,SZLITM      ,SZAITM      ,SZDSC1                 ,SZLNTY,SZUOM,SZUORG    ,SZUOPN  ,SZTX,SZEXR1,SZTXA1,SZANBY ,SZPQOR    ,SZUOM2,SZSQOR    ,SZGLC ,SZVR01      ,SZVR02      ,SZSFXO,SZDGL ,SZNXTR,SZLTTR,SZCRMD,SZUOM1,SZAVCH,SZUNCD,SZCRCD) VALUES " & _
                              "                    ('EDELGADO','" & CLng(Right(DOCNUM, 8)) & "' ,'0'   ," & linea * 1000 & " ,''   ,'JDEINSAP'  ,'0'   ,'1'   ,'0'   ,'N'   ,''    ,'02'  ,'00300',0     ,'O5'  ," & linea * 1000 & "  ,'" & buscarCampoItemF4104(CLng(IDTNR), "IVDSC2") & "','00300','00300','" & CLng(Right(DOCNUM, 8)) & "','47301','30000099'," & fechaJuliana & "," & fechaJuliana & "," & fechaJuliana & ",'" & buscarCampoItemF4104(CLng(IDTNR), "IVITM") & "'   ,'" & item & "','" & item & "','" & buscarCampoItemF4104(CLng(IDTNR), "IVDSC1") & "','S5'  ,'" & uom & "' ," & (CDbl(MENGE) / 1) * 10000 & "," & (CDbl(MENGE) / 1) * 100 & ",'Y' ,'V'   ,'IVA' ,11409  ," & (CDbl(MENGE) / 1) * 10000 & ",'" & uom & "'  ," & (CDbl(MENGE) / 1) * 10000 & ",'CS08','" & BELNR & "','" & CLng(Right(DOCNUM, 8)) & "','000' ," & fechaJuliana & ",280   ,230   ,'2'   ,'" & uom & "'  ,'N'   ,'N'   ,'VEF')"
                        comandoAS400.CommandText = sql
                        comandoAS400.ExecuteNonQuery()
                        linea = linea + 1

                    Next

                    cnn.Close()
                    myTrans.Commit()

                    If File.Exists(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar) Then
                        File.Delete(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar)
                    End If

                    fs.Close()
                    File.Copy(nombreArchivoProcesar, Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\BK\" & nombreArchivoProcesar)
                    File.Delete(nombreArchivoProcesar)

                    escribirLog("FINALIZADA IMPORTACION DE DATOS DE DOCUMENTO:" & DOCNUM, "(PO) ")
                    conexion.Close()

                End If
            Catch ex As Exception
                'myTrans.Rollback()
                escribirLog(ex.Message.ToString, "(PO) ")
            End Try

        Next



    End Sub

    Public Function IsAñoBisiesto(ByVal YYYY As Integer) As Boolean
        Return YYYY Mod 4 = 0 _
                    And (YYYY Mod 100 <> 0 Or YYYY Mod 400 = 0)
    End Function

    Private Function buscarCampoItemF4104(ByVal codigo As String, ByVal campo As String) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT " & campo & " FROM F4104 WHERE IVXRT='DO' AND IVCITM='" & codigo & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader(campo))
        End While

        If campo = "IVDSC2" Then
            valor = "    " & Trim(valor)
        End If

        cnn.Close()

        Return valor
    End Function


    Private Function buscarItemDiageo(ByVal codigo As String) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT IVCITM FROM F4104 WHERE IVXRT='DO' AND IVLITM='" & codigo & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("IVCITM"))
        End While

        cnn.Close()

        Return valor
    End Function


    Private Function buscarMaxF0041Z1() As Double

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT MAX(DEC(ZFTRNM))as ZFTRNM FROM F0041Z1 ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As Double
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = reader("ZFTRNM")
            valor = valor + 1
        End While

        cnn.Close()

        Return valor
    End Function




    Private Function buscarCantidadLotesUtilizados(ByVal docnum As Long) As Integer

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT Count(*) as cantidad FROM F55DD WHERE  DGVR01='" & docnum & "' group by dglotn ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As Integer
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            'valor = Trim(reader("cantidad"))
            valor = valor + 1
        End While


        cnn.Close()

        Return valor
    End Function


    Private Function buscarCantidadLotesUtilizadosItem(ByVal docnum As Long, ByVal idItem As Long) As Integer

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT Count(*) as cantidad FROM F55DD WHERE  DGVR01='" & docnum & "' AND DGAA18='" & idItem & "' group by dglotn ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As Integer
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            'valor = Trim(reader("cantidad"))
            valor = valor + 1
        End While


        cnn.Close()

        Return valor
    End Function



    Private Function buscarLotesUtilizado(ByVal docnum As Long, ByVal idItem As Long) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT DGLOTN  FROM F55DD WHERE  DGVR01='" & docnum & "' AND DGAA18='" & idItem & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = ""
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("DGLOTN"))
        End While


        cnn.Close()

        Return valor
    End Function


    Private Function buscarCantidadUtilizada(ByVal docnum As Long, ByVal lote As String, ByVal unidad As String, ByVal pt As String) As Double


        Dim unidad_utilizada As String
        unidad_utilizada = buscarUOMUtilizado(docnum)

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT sum(DGTRQT) as DGTRQT FROM F55DD WHERE  DGVR01='" & docnum & "' and DGLOTN='" & lote & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As Double
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Math.Abs(CDbl(Trim(reader("DGTRQT"))) / 10000)
        End While

        cnn.Close()

        If unidad_utilizada = "CA" Then

            If Trim(unidad_utilizada) <> buscarUOM(Trim(unidad)) Then
                valor = valor * buscarPackingSize(pt)
            End If

        End If
        Return valor
    End Function


    Private Function buscarUOMUtilizado(ByVal docnum As Long) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT DGTRUM  FROM F55DD WHERE  DGVR01='" & docnum & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = ""
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("DGTRUM"))
        End While

        cnn.Close()

        Return valor
    End Function




    Private Function buscarUOM(ByVal UOM As String) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400CTL)
        Dim rs As New Odbc.OdbcCommand("SELECT DRDL01 FROM F0005 WHERE DRSY='55' AND DRRT='I7' AND DRKY='     " & UOM & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("DRDL01"))
        End While

        cnn.Close()

        Return valor
    End Function


    Private Function buscarUOMF4101(ByVal item As String) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("select imuom1 from F4101 where imitm in (SELECT ivitm FROM F4104  where ivxrt='DO' and ivcitm='" & item & "')   ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("IMUOM1"))
        End While

        cnn.Close()

        Return valor
    End Function



    Private Function buscarUOMSRF(ByVal UOM As String) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400CTL)
        Dim rs As New Odbc.OdbcCommand("SELECT DRKY FROM F0005 WHERE DRSY='55' AND DRRT='I7' AND DRDL01='" & UOM & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("DRKY"))
        End While

        cnn.Close()

        Return valor
    End Function

    Private Sub crearGoodReceiptsPO()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String


        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand


        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_PO.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


        Try

            Dim cnn2 As New Odbc.OdbcConnection(cadenaAS400)
            Dim rs2 As New Odbc.OdbcCommand("SELECT DRBCTK,DRUREC,DRLITM,DRLOTN,DRMCU,DRURRF FROM F55DR WHERE DRCDCTYPE='PO' AND DRURCD=''   ", cnn2)
            Dim reader2 As Odbc.OdbcDataReader

            cnn2.Open()
            reader2 = rs2.ExecuteReader
            While reader2.Read()

                Dim linea As String

                conEx.ConnectionString = conexionString
                conEx.Open()
                cmdEx.Connection = conEx
                cmdEx.CommandText = "SELECT * FROM CABECERA_PO_ADVICE WHERE DOCNUM ='" & Right("0000000000000000" & Trim(reader2.GetString(0)), 16) & "'"
                Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()

                While lrdEx.Read()

                    conIn.ConnectionString = conexionString
                    conIn.Open()
                    docnum = lrdEx.GetString("0")

                    Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombrePartner & "_ZWHGRC_" & obtenerFechaHora() & ".xml")

                    linea = ""
                    linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                    linea = linea & "<ZWMMBID2>" & vbNewLine
                    linea = linea & "<IDOC BEGIN=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<EDI_DC40 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<TABNAM>EDI_DC40</TABNAM>" & vbNewLine
                    linea = linea & "<MANDT>600</MANDT>" & vbNewLine
                    linea = linea & "<DOCNUM>" & obtenerNextNumber() & "</DOCNUM>" & vbNewLine
                    linea = linea & "<DOCREL>701</DOCREL>" & vbNewLine
                    linea = linea & "<STATUS>30</STATUS>" & vbNewLine
                    linea = linea & "<DIRECT>1</DIRECT>" & vbNewLine
                    linea = linea & "<OUTMOD>2</OUTMOD>" & vbNewLine
                    linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>" & vbNewLine
                    linea = linea & "<CIMTYP>ZWMMBID2</CIMTYP>" & vbNewLine
                    linea = linea & "<MESTYP>ZWHGRC</MESTYP>" & vbNewLine
                    linea = linea & "<STDMES>ZWHGRC</STDMES>" & vbNewLine
                    linea = linea & "<SNDPOR>SAPGU3</SNDPOR>" & vbNewLine
                    linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
                    linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>" & vbNewLine
                    linea = linea & "<RCVPOR>PI_XML</RCVPOR>" & vbNewLine
                    linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
                    linea = linea & "<RCVPRN>STPG</RCVPRN>" & vbNewLine
                    linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>" & vbNewLine
                    linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>" & vbNewLine
                    linea = linea & "</EDI_DC40>" & vbNewLine
                    linea = linea & "<E1MBXYH SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<BUDAT>" & obtenerFecha() & "</BUDAT>" & vbNewLine
                    linea = linea & "<BKTXT>" & Trim(reader2.GetString(5)) & "</BKTXT>" & vbNewLine
                    linea = linea & "<ZGREC01 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<DOC_TYP>PO</DOC_TYP>" & vbNewLine
                    linea = linea & "</ZGREC01>" & vbNewLine
                    linea = linea & "<ZSTKDATE SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<DATE>" & obtenerFecha() & "</DATE>" & vbNewLine
                    linea = linea & "<TIME>" & obtenerHora() & "</TIME>" & vbNewLine
                    linea = linea & "</ZSTKDATE>" & vbNewLine

                    Dim cmdIn As New SqlCommand
                    cmdIn.Connection = conIn
                    cmdIn.CommandText = "SELECT * FROM DETALLE_PO_ADVICE WHERE DOCNUM='" & lrdEx.GetString(0) & "' AND IDTNR='" & Right("000000000000000000" & buscarItemDiageo(Trim(reader2.GetString(2))), 18) & "'"
                    Dim lrdIn As SqlDataReader = cmdIn.ExecuteReader()

                    While lrdIn.Read()

                        linea = linea & "<E1MBXYI SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                        linea = linea & "<MATNR>" & lrdIn.GetString(8) & "</MATNR>" & vbNewLine
                        linea = linea & "<WERKS>" & lrdIn.GetString(7) & "</WERKS>" & vbNewLine
                        linea = linea & "<LGORT>" & lrdIn.GetString(12) & "</LGORT>" & vbNewLine

                        If reader2.GetString(4) = "    300AMP01" Then
                            linea = linea & "<CHARG>" & Trim(reader2.GetString(3)) & "</CHARG>" & vbNewLine
                        Else
                            linea = linea & "<CHARG></CHARG>" & vbNewLine
                        End If




                        'Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
                        'Dim rs As New Odbc.OdbcCommand("SELECT DRUREC FROM F55DR WHERE DRBCTK='" & CLng(docnum) & "'  ", cnn)
                        'Dim reader As Odbc.OdbcDataReader
                        'Dim valor As String
                        'valor = 0
                        'cnn.Open()
                        'reader = rs.ExecuteReader
                        'While reader.Read()

                        'End While
                        'cnn.Close()

                        linea = linea & "<ERFMG>" & CDbl(Replace(reader2.GetDecimal(1), "D", "")) / 10000 & "</ERFMG>" & vbNewLine

                        linea = linea & "<ERFME>" & lrdIn.GetString(4) & "</ERFME>" & vbNewLine
                        linea = linea & "<EBELN>" & lrdEx.GetString(3) & "</EBELN>" & vbNewLine
                        linea = linea & "<EBELP>" & Right(lrdIn.GetString(1), 5) & "</EBELP>" & vbNewLine
                        linea = linea & "<ZGREC02 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine

                        If lrdIn.GetString(12).Substring(lrdIn.GetString(12).Length - 2, 2) = "10" Or lrdIn.GetString(12).Substring(lrdIn.GetString(12).Length - 2, 2) = "11" Then
                            linea = linea & "<DUTY_STAT>P</DUTY_STAT>" & vbNewLine
                        Else
                            linea = linea & "<DUTY_STAT>S</DUTY_STAT>" & vbNewLine
                        End If

                        linea = linea & "<STOCK_STAT>FREE</STOCK_STAT>" & vbNewLine
                        linea = linea & "<SPECIAL>1</SPECIAL>" & vbNewLine
                        linea = linea & "<USRTXT2>" & Trim(reader2.GetString(5)) & "</USRTXT2>" & vbNewLine
                        linea = linea & "</ZGREC02>" & vbNewLine
                        linea = linea & "</E1MBXYI>" & vbNewLine

                        'End While


                        Dim cnn3 As New Odbc.OdbcConnection(cadenaAS400)
                        Dim rs3 As New Odbc.OdbcCommand("UPDATE F55DR SET DRURCD='Y' WHERE DRCDCTYPE='PO' AND DRBCTK=" & docnum & "  ", cnn3)
                        Dim reader3 As Odbc.OdbcDataReader

                        cnn3.Open()
                        reader3 = rs3.ExecuteReader
                        reader3.Close()
                        cnn3.Close()

                    End While

                    linea = linea & "</E1MBXYH>" & vbNewLine
                    linea = linea & "</IDOC>" & vbNewLine
                    linea = linea & "</ZWMMBID2>"
                    oSW.WriteLine(linea)
                    oSW.Flush()
                    oSW.Close()

                    actualizarGenerado("CABECERA_PO_ADVICE", docnum)

                    System.Threading.Thread.Sleep(2000)




                    conIn.Close()


                End While

                conEx.Close()



            End While
            cnn2.Close()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


    End Sub

    Private Sub crearGoodReceiptsSTO()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String


        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand


        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_STO.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


        Try

            Dim cnn2 As New Odbc.OdbcConnection(cadenaAS400)
            Dim rs2 As New Odbc.OdbcCommand("SELECT DRBCTK,DRUREC,DRLITM,DRURRF FROM F55DR WHERE DRCDCTYPE='STO' AND DRURCD=''  ", cnn2)
            Dim reader2 As Odbc.OdbcDataReader

            cnn2.Open()
            reader2 = rs2.ExecuteReader
            While reader2.Read()




                Dim linea As String


                conEx.ConnectionString = conexionString
                conEx.Open()
                cmdEx.Connection = conEx
                cmdEx.CommandText = "SELECT * FROM CABECERA_STO_ADVICE WHERE DOCNUM ='" & Right("0000000000000000" & Trim(reader2.GetString(0)), 16) & "'"
                Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()

                While lrdEx.Read()

                    conIn.ConnectionString = conexionString
                    conIn.Open()
                    docnum = lrdEx.GetString("0")

                    Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombrePartner & "_ZWHGRC_" & obtenerFechaHora() & ".xml")

                    linea = ""
                    linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                    linea = linea & "<ZWMMBID2>" & vbNewLine
                    linea = linea & "<IDOC BEGIN=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<EDI_DC40 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<TABNAM>EDI_DC40</TABNAM>" & vbNewLine
                    linea = linea & "<MANDT>600</MANDT>" & vbNewLine
                    linea = linea & "<DOCNUM>" & obtenerNextNumber() & "</DOCNUM>" & vbNewLine
                    linea = linea & "<DOCREL>701</DOCREL>" & vbNewLine
                    linea = linea & "<STATUS>30</STATUS>" & vbNewLine
                    linea = linea & "<DIRECT>1</DIRECT>" & vbNewLine
                    linea = linea & "<OUTMOD>2</OUTMOD>" & vbNewLine
                    linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>" & vbNewLine
                    linea = linea & "<CIMTYP>ZWMMBID2</CIMTYP>" & vbNewLine
                    linea = linea & "<MESTYP>ZWHGRC</MESTYP>" & vbNewLine
                    linea = linea & "<STDMES>ZWHGRC</STDMES>" & vbNewLine
                    linea = linea & "<SNDPOR>SAPGU3</SNDPOR>" & vbNewLine
                    linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
                    linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>" & vbNewLine
                    linea = linea & "<RCVPOR>PI_XML</RCVPOR>" & vbNewLine
                    linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
                    linea = linea & "<RCVPRN>STPG</RCVPRN>" & vbNewLine
                    linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>" & vbNewLine
                    linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>" & vbNewLine
                    linea = linea & "</EDI_DC40>" & vbNewLine
                    linea = linea & "<E1MBXYH SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<BUDAT>" & obtenerFecha() & "</BUDAT>" & vbNewLine
                    linea = linea & "<BKTXT>" & Trim(reader2.GetString(3)) & "</BKTXT>" & vbNewLine
                    linea = linea & "<ZGREC01 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<DOC_TYP>PO</DOC_TYP>" & vbNewLine
                    linea = linea & "</ZGREC01>" & vbNewLine
                    linea = linea & "<ZSTKDATE SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<DATE>" & obtenerFecha() & "</DATE>" & vbNewLine
                    linea = linea & "<TIME>" & obtenerHora() & "</TIME>" & vbNewLine
                    linea = linea & "</ZSTKDATE>" & vbNewLine

                    Dim cmdIn As New SqlCommand
                    cmdIn.Connection = conIn
                    cmdIn.CommandText = "SELECT * FROM DETALLE_STO_ADVICE WHERE DOCNUM='" & lrdEx.GetString(0) & "' AND MATNR='" & Right("000000000000000000" & buscarItemDiageo(Trim(reader2.GetString(2))), 18) & "'"
                    Dim lrdIn As SqlDataReader = cmdIn.ExecuteReader()

                    While lrdIn.Read()

                        linea = linea & "<E1MBXYI SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                        linea = linea & "<MATNR>" & lrdIn.GetString(2) & "</MATNR>" & vbNewLine
                        linea = linea & "<WERKS>" & lrdEx.GetString(5) & "</WERKS>" & vbNewLine
                        linea = linea & "<LGORT>" & lrdIn.GetString(6) & "</LGORT>" & vbNewLine
                        linea = linea & "<CHARG></CHARG>" & vbNewLine


                        'Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
                        'Dim rs As New Odbc.OdbcCommand("SELECT DRUREC FROM F55DR WHERE DRBCTK='" & CLng(docnum) & "'  ", cnn)
                        'Dim reader As Odbc.OdbcDataReader
                        'Dim valor As String
                        'valor = 0
                        'cnn.Open()
                        'reader = rs.ExecuteReader
                        'While reader.Read()

                        'End While
                        'cnn.Close()

                        ' linea = linea & "<ERFMG>" & CDbl(Replace(reader2.GetDecimal(1), "D", "")) / 10000 & "</ERFMG>" & vbNewLine

                        'linea = linea & "<ERFME>" & lrdIn.GetString(4) & "</ERFME>" & vbNewLine

                        linea = linea & "<ERFMG>" & lrdIn.GetString(8) & "</ERFMG>" & vbNewLine

                        linea = linea & "<ERFME>" & lrdIn.GetString(9) & "</ERFME>" & vbNewLine

                        linea = linea & "<EBELN>" & lrdIn.GetString(13) & "</EBELN>" & vbNewLine
                        linea = linea & "<EBELP>" & Right(lrdIn.GetString(14), 5) & "</EBELP>" & vbNewLine
                        linea = linea & "<ZGREC02 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine

                        If lrdIn.GetString(6).Substring(lrdIn.GetString(6).Length - 2, 2) = "10" Or lrdIn.GetString(6).Substring(lrdIn.GetString(6).Length - 2, 2) = "11" Then
                            linea = linea & "<DUTY_STAT>P</DUTY_STAT>" & vbNewLine
                        Else
                            linea = linea & "<DUTY_STAT>S</DUTY_STAT>" & vbNewLine
                        End If

                        linea = linea & "<STOCK_STAT>FREE</STOCK_STAT>" & vbNewLine
                        linea = linea & "<SPECIAL>1</SPECIAL>" & vbNewLine
                        linea = linea & "<USRTXT2>" & Trim(reader2.GetString(3)) & "</USRTXT2>" & vbNewLine
                        linea = linea & "</ZGREC02>" & vbNewLine
                        linea = linea & "</E1MBXYI>" & vbNewLine

                        'End While


                        Dim cnn3 As New Odbc.OdbcConnection(cadenaAS400)
                        Dim rs3 As New Odbc.OdbcCommand("UPDATE F55DR SET DRURCD='Y' WHERE DRCDCTYPE='STO' AND DRBCTK=" & docnum & "  ", cnn3)
                        Dim reader3 As Odbc.OdbcDataReader

                        cnn3.Open()
                        reader3 = rs3.ExecuteReader
                        reader3.Close()
                        cnn3.Close()


                    End While

                    linea = linea & "</E1MBXYH>" & vbNewLine
                    linea = linea & "</IDOC>" & vbNewLine
                    linea = linea & "</ZWMMBID2>"
                    oSW.WriteLine(linea)
                    oSW.Flush()
                    oSW.Close()

                    actualizarGenerado("CABECERA_PO_ADVICE", docnum)

                    System.Threading.Thread.Sleep(1000)




                    conIn.Close()


                End While

                conEx.Close()



            End While
            cnn2.Close()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


    End Sub



    Private Function obtenerIdPedido(ByVal usuario As String) As Integer

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim comando As New SqlClient.SqlCommand
        Dim sqlstring As String

        Try
            conEx2.ConnectionString = conexionStringPedidos
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT (last_serial + 1) as id FROM salesmen WHERE salesman_id='" & usuario & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetInt32("0")
            End While

            lrdEx2.Close()

            sqlstring = ""
            sqlstring = " UPDATE salesmen set last_serial=last_serial+1 WHERE salesman_id='" & usuario & "'"
            comando.Connection = conEx2
            comando.CommandText = " "
            comando.CommandText = sqlstring
            comando.ExecuteNonQuery()
            comando.Parameters.Clear()

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function


    Private Function obtenerNextNumber() As Integer

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim sqlstring As String

        Try

            conEx2.ConnectionString = conexionString
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT * FROM DOCNUMS "
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetInt32("0")
            End While

            lrdEx2.Close()
            sqlstring = ""
            sqlstring = "UPDATE DOCNUMS SET CONTADOR='" & valor + 1 & "' "
            comando.Connection = conEx2
            comando.CommandText = " "
            comando.CommandText = sqlstring
            comando.ExecuteNonQuery()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try

        Return valor
    End Function

    Private Function buscarItemPtOLD(ByVal pt As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String

        Try
            conEx2.ConnectionString = conexionStringPedidos
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT product_id FROM productsdg WHERE PRODID_SAP='" & CDbl(pt) & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetString("0")
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function



    Private Function buscarItemPt(ByVal pt As String) As String

        
        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT IVLITM FROM F4104 WHERE IVXRT='DO' AND IVCITM='" & CDbl(pt) & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("IVLITM"))
        End While

        cnn.Close()

        Return valor

    End Function

    Private Function buscarClientePt(ByVal pt As String) As String


        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT IMCARP FROM F4101 WHERE IMLITM='" & pt & "' ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("IMCARP"))
        End While

        cnn.Close()

        Return valor

    End Function




    Private Function buscarMarcaPtOld(ByVal pt As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String

        Try
            conEx2.ConnectionString = conexionStringPedidos
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT brand FROM productsdg WHERE PRODID_SAP='" & CDbl(pt) & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetString("0")
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function


    Private Function buscarMarcaPt(ByVal pt As String) As String

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT IMSRP7 FROM F4101 WHERE  IMITM IN ( SELECT IVITM FROM F4104 WHERE IVXRT='DO' AND IVCITM='" & CDbl(pt) & "') ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = Trim(reader("IMSRP7"))
        End While

        cnn.Close()

        Return valor
    End Function


   


    Private Function buscarCustomerPt(ByVal pt As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String

        Try
            conEx2.ConnectionString = conexionStringPedidos
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT customer_id FROM productsdg WHERE PRODID_SAP='" & CDbl(pt) & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetString("0")
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function



    Private Function buscarPackingSizeOLD(ByVal pt As String) As Integer

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer

        valor = 1
        Try
            conEx2.ConnectionString = conexionStringPedidos
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT conversion_rate FROM productsdg WHERE PRODID_SAP='" & CDbl(pt) & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = Int32.Parse(lrdEx2.GetString("0"))
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function


    Private Function buscarPackingSize(ByVal pt As String) As Integer



        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT UMCONV  FROM F41002 WHERE  UMUM='CA' AND UMRUM='BT' AND UMITM IN ( SELECT IVITM FROM F4104 WHERE IVXRT='DO' AND IVCITM='" & pt & "') ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As Integer
        valor = 0
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            valor = CDbl(Trim(reader("UMCONV"))) / 10000000
        End While

        cnn.Close()

        Return valor

    End Function





    Private Function buscarAlmacenPtOLD(ByVal pt As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String

        Try
            conEx2.ConnectionString = conexionStringPedidos
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT warehouse FROM productsdg WHERE PRODID_SAP='" & pt & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetString("0")
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function


    Private Function buscarAlmacenPt(ByVal pt As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String

        Try

            valor = buscarCampoItemF4104(pt, "IVDSC2")

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()
        End Try


        Return valor
    End Function



    Private Function buscarEstadoDeliveryDispatch(ByVal storage As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String

        Try
            conEx2.ConnectionString = conexionString
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT * FROM STATUS_DELIVERY_DISPATCH WHERE STORAGE_LOCATION='" & storage & "'"
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = lrdEx2.GetString("1")
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function

    Private Function buscarFechaJuliana() As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As String
        Dim bisiesto As String

        Try

            If IsAñoBisiesto(Year(CDate(Now))) Then
                bisiesto = "S"
            Else
                bisiesto = "N"
            End If

            conEx2.ConnectionString = conexionString
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT DIA FROM CONVERSION_FECHA WHERE ESBISIESTO='" & bisiesto & "' AND  fecha like '%" & Right("00" & Month(CDate(Now)), 2) & Right("00" & Day(CDate(Now)), 2) & "%'  "
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                valor = "1" & Right(Year(CDate(Now)), 2) & lrdEx2.GetString("0")
            End While

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try


        Return valor
    End Function


    Private Sub crearDeliveryAcknowledgement()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim comando As New SqlCommand

        Try


            conIn.ConnectionString = conexionString

            Try
                conIn.Open()

            Catch ex As Exception
                escribirLog(ex.Message.ToString, "(OUTB) ")
            End Try


            Dim sqlstring As String
            

            Dim linea As String

            conEx.ConnectionString = conexionString
            conEx.Open()
            cmdEx.Connection = conEx
            cmdEx.CommandText = "SELECT * FROM CABECERA_OUTBOUND WHERE STATUS_GENERADO='N'"
            Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()

            While lrdEx.Read()

                Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\STATUS\" & nombrePartner & "_STATUS_" & obtenerFechaHora() & ".xml")

                linea = ""
                linea = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                linea = linea & "<SYSTAT01>" & vbNewLine
                linea = linea & "<IDOC>" & vbNewLine
                linea = linea & "<EDI_DC40>" & vbNewLine
                linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
                linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>" & vbNewLine
                linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
                linea = linea & "<RCVPRN></RCVPRN>" & vbNewLine
                linea = linea & "</EDI_DC40>" & vbNewLine
                linea = linea & "<E1STATS>" & vbNewLine
                linea = linea & "<MANDT>600</MANDT>" & vbNewLine
                linea = linea & "<DOCNUM>" & lrdEx.GetString(0) & "</DOCNUM>" & vbNewLine
                linea = linea & "<LOGDAT>" & obtenerFecha() & "</LOGDAT>" & vbNewLine
                linea = linea & "<LOGTIM>" & obtenerHora() & "</LOGTIM>" & vbNewLine
                linea = linea & "<STATUS>41</STATUS>" & vbNewLine
                linea = linea & "<STACOD>41</STACOD>" & vbNewLine
                linea = linea & "<STATYP>I</STATYP>" & vbNewLine
                linea = linea & "</E1STATS>" & vbNewLine
                linea = linea & "</IDOC>" & vbNewLine
                linea = linea & "</SYSTAT01>"
                oSW.WriteLine(linea)
                oSW.Flush()
                oSW.Close()

                sqlstring = ""
                sqlstring = "UPDATE [CABECERA_OUTBOUND] set STATUS_GENERADO='S' WHERE DOCNUM='" & lrdEx.GetString(0) & "' "
                comando.Connection = conIn

                comando.CommandText = " "
                comando.CommandText = sqlstring
                comando.ExecuteNonQuery()

                System.Threading.Thread.Sleep(2000)

            End While

            conEx.Close()

        Catch oe As Exception
            escribirLog(oe.Message.ToString, "(OUTB) ")
        End Try

    End Sub

    Private Sub crearDeliveryDispatch()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim conexion As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String
        Dim cantidadLotes As Integer
        Dim sqlstring As String

        cantidadLotes = 0

        Try

            conEx.ConnectionString = conexionString
            conEx.Open()
            cmdEx.Connection = conEx
            cmdEx.CommandText = "SELECT * FROM CABECERA_OUTBOUND WHERE GENERADO='N'"
            Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()

            While lrdEx.Read()

                cantidadLotes = buscarCantidadLotesUtilizados(lrdEx.GetString(0))

                If cantidadLotes > 0 Then

                    Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombrePartner & "_ZWHDES_" & obtenerFechaHora() & ".xml")

                    Dim linea As String
                    linea = ""

                    docnum = lrdEx.GetString("0")
                    linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                    linea = linea & "<ZDELIVRY>" & vbNewLine
                    linea = linea & "<IDOC>" & vbNewLine
                    linea = linea & "<EDI_DC40>" & vbNewLine
                    linea = linea & "<TABNAM>EDI_DC40</TABNAM>" & vbNewLine
                    linea = linea & "<MANDT>600</MANDT>" & vbNewLine
                    linea = linea & "<DOCNUM>" & lrdEx.GetString(0) & "</DOCNUM>" & vbNewLine
                    linea = linea & "<DIRECT />" & vbNewLine
                    linea = linea & "<IDOCTYP>DELVRY01</IDOCTYP>" & vbNewLine
                    linea = linea & "<CIMTYP>ZDELIVRY</CIMTYP>" & vbNewLine
                    linea = linea & "<MESTYP>ZWHDES</MESTYP>" & vbNewLine
                    linea = linea & "<SNDPOR />" & vbNewLine
                    linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
                    linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>" & vbNewLine
                    linea = linea & "<RCVPOR />" & vbNewLine
                    linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
                    linea = linea & "<RCVPRN></RCVPRN>" & vbNewLine
                    linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>" & vbNewLine
                    linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>" & vbNewLine
                    linea = linea & "</EDI_DC40>" & vbNewLine
                    linea = linea & "<E1EDL20 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<VBELN>" & lrdEx.GetString(3) & "</VBELN>" & vbNewLine
                    linea = linea & "<Z1WBIN1 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<DATE>" & obtenerFecha() & "</DATE>" & vbNewLine
                    linea = linea & "<TIME>" & obtenerHora() & "</TIME>" & vbNewLine
                    linea = linea & "</Z1WBIN1>" & vbNewLine
                    linea = linea & "<E1ADRM1 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<PARTNER_Q>WE</PARTNER_Q>" & vbNewLine
                    linea = linea & "<PARTNER_ID>" & lrdEx.GetString(5) & "</PARTNER_ID>" & vbNewLine
                    linea = linea & "</E1ADRM1>" & vbNewLine
                    linea = linea & "<E1EDT13 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<QUALF>006</QUALF>" & vbNewLine
                    linea = linea & "<NTANF>" & obtenerFecha() & "</NTANF>" & vbNewLine
                    linea = linea & "<NTANZ>000000</NTANZ>" & vbNewLine
                    linea = linea & "</E1EDT13>" & vbNewLine


                    conIn.ConnectionString = conexionString
                    conIn.Open()
                    Dim cmdIn As New SqlCommand
                    cmdIn.Connection = conIn
                    cmdIn.CommandText = "SELECT * FROM DETALLE_OUTBOUND WHERE DOCNUM='" & lrdEx.GetString(0) & "' AND HIPOS='0' "
                    Dim lrdIn As SqlDataReader = cmdIn.ExecuteReader()

                    While lrdIn.Read()

                        conexion.ConnectionString = conexionString
                        conexion.Open()

                        sqlstring = "DELETE FROM [DETALLE_OUTBOUND] WHERE DOCNUM='" & lrdEx.GetString(0) & "' and cast(posnr as int) > 900000 "
                        comando.Connection = conexion
                        comando.CommandText = " "
                        comando.CommandText = sqlstring
                        comando.ExecuteNonQuery()
                        comando.Parameters.Clear()


                        cantidadLotes = buscarCantidadLotesUtilizadosItem(lrdEx.GetString(0), lrdIn.GetString(2))

                        linea = linea & "<E1EDL24 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                        linea = linea & "<POSNR>" & lrdIn.GetString(1) & "</POSNR>" & vbNewLine
                        linea = linea & "<MATNR>" & lrdIn.GetString(2) & "</MATNR>" & vbNewLine
                        linea = linea & "<WERKS>" & lrdIn.GetString(5) & "</WERKS>" & vbNewLine
                        linea = linea & "<LGORT>" & lrdIn.GetString(6) & "</LGORT>" & vbNewLine

                        If cantidadLotes = 1 Then
                            linea = linea & "<CHARG>" & Trim(buscarLotesUtilizado(lrdEx.GetString(0), CLng(lrdIn.GetString(2)))) & "</CHARG>" & vbNewLine
                        Else
                            linea = linea & "<CHARG></CHARG>" & vbNewLine
                        End If

                        linea = linea & "<LFIMG>" & lrdIn.GetString(7) & "</LFIMG>" & vbNewLine
                        linea = linea & "<VRKME>" & lrdIn.GetString(8) & "</VRKME>" & vbNewLine

                        If cantidadLotes = 1 Then

                        Else
                            linea = linea & "<HIPOS>00000</HIPOS>" & vbNewLine
                        End If


                        linea = linea & "<Z1WDES1 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                        linea = linea & "<DUTY_STATUS>" & buscarEstadoDeliveryDispatch(lrdIn.GetString(6).Trim) & "</DUTY_STATUS>" & vbNewLine
                        linea = linea & "</Z1WDES1>" & vbNewLine
                        linea = linea & "</E1EDL24>" & vbNewLine


                        If cantidadLotes = 1 Then

                            sqlstring = "DELETE FROM [DETALLE_OUTBOUND] WHERE DOCNUM='" & docnum & "' AND MATNR ='" & lrdIn.GetString(2) & "'"
                            comando.Connection = conexion
                            comando.CommandText = " "
                            comando.CommandText = sqlstring
                            comando.ExecuteNonQuery()
                            comando.Parameters.Clear()

                            sqlstring = "INSERT INTO [DETALLE_OUTBOUND] VALUES('" & docnum & "','" & lrdIn.GetString(1) & "','" & lrdIn.GetString(2) & "','" & lrdIn.GetString(3) & "','" & lrdIn.GetString(4) & "','" & lrdIn.GetString(5) & "','" & lrdIn.GetString(6) & "','" & lrdIn.GetString(7) & "','" & lrdIn.GetString(8) & "','" & lrdIn.GetString(9) & "','" & lrdIn.GetString(8) & "','" & lrdIn.GetString(11) & "','" & lrdIn.GetString(12) & "','" & lrdIn.GetString(13) & "','" & Trim(buscarLotesUtilizado(CLng(lrdEx.GetString(0)), CLng(lrdIn.GetString(2)))) & "','0')"
                            comando.Connection = conexion
                            comando.CommandText = " "
                            comando.CommandText = sqlstring
                            comando.ExecuteNonQuery()
                            comando.Parameters.Clear()

                        End If


                        If cantidadLotes > 1 Then

                            Dim POSNR_COUNT As Integer
                            Dim POSNR_BASE As Integer
                            POSNR_BASE = 900000
                            POSNR_COUNT = 1



                            Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
                            Dim rs As New Odbc.OdbcCommand("select dglotn,dgtrum,sum(dgtrqt) from F55DD WHERE  DGVR01='" & CLng(docnum) & "' AND DGAA18='" & CLng(lrdIn.GetString(2)) & "' group by dglotn,dgtrum  ", cnn)
                            Dim reader As Odbc.OdbcDataReader
                            Dim valor As Integer
                            valor = 0
                            cnn.Open()

                            reader = rs.ExecuteReader
                            While reader.Read()
                                linea = linea & "<E1EDL24 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                                linea = linea & "<POSNR>" & POSNR_BASE + POSNR_COUNT & "</POSNR>" & vbNewLine
                                linea = linea & "<MATNR>" & lrdIn.GetString(2) & "</MATNR>" & vbNewLine
                                linea = linea & "<WERKS>" & lrdIn.GetString(5) & "</WERKS>" & vbNewLine
                                linea = linea & "<LGORT>" & lrdIn.GetString(6) & "</LGORT>" & vbNewLine
                                linea = linea & "<CHARG>" & Trim(reader("DGLOTN")) & "</CHARG>" & vbNewLine
                                Dim cantidadUtilizada As String
                                cantidadUtilizada = buscarCantidadUtilizada(lrdIn.GetString(0), Trim(reader("DGLOTN")), lrdIn.GetString(8), CLng(lrdIn.GetString(2)))
                                linea = linea & "<LFIMG>" & Replace(cantidadUtilizada, ",", ".") & "</LFIMG>" & vbNewLine
                                linea = linea & "<VRKME>" & lrdIn.GetString(8) & "</VRKME>" & vbNewLine
                                linea = linea & "<HIPOS>" & Right(lrdIn.GetString(1), 5) & "</HIPOS>" & vbNewLine
                                linea = linea & "<Z1WDES1 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
                                linea = linea & "<DUTY_STATUS>" & buscarEstadoDeliveryDispatch(lrdIn.GetString(6).Trim) & "</DUTY_STATUS>" & vbNewLine
                                linea = linea & "</Z1WDES1>" & vbNewLine
                                linea = linea & "</E1EDL24>" & vbNewLine


                                sqlstring = "INSERT INTO [DETALLE_OUTBOUND] VALUES('" & docnum & "','" & POSNR_BASE + POSNR_COUNT & "','" & lrdIn.GetString(2) & "','" & lrdIn.GetString(3) & "','" & lrdIn.GetString(4) & "','" & lrdIn.GetString(5) & "','" & lrdIn.GetString(6) & "','" & Replace(cantidadUtilizada, ",", ".") & "','" & lrdIn.GetString(8) & "','" & lrdIn.GetString(9) & "','" & lrdIn.GetString(8) & "','" & lrdIn.GetString(11) & "','" & lrdIn.GetString(12) & "','" & lrdIn.GetString(13) & "','" & Trim(reader("DGLOTN")) & "','" & Right(lrdIn.GetString(1), 5) & "')"
                                comando.Connection = conexion
                                comando.CommandText = " "
                                comando.CommandText = sqlstring
                                comando.ExecuteNonQuery()
                                comando.Parameters.Clear()

                                POSNR_COUNT = POSNR_COUNT + 1

                            End While

                            cnn.Close()

                        End If

                        conexion.Close()

                    End While


                    linea = linea & "</E1EDL20>" & vbNewLine
                    linea = linea & "</IDOC>" & vbNewLine
                    linea = linea & "</ZDELIVRY>"


                    oSW.WriteLine(linea)
                    oSW.Flush()
                    oSW.Close()

                    actualizarGenerado("CABECERA_OUTBOUND", docnum)

                    System.Threading.Thread.Sleep(2000)


                    conIn.Close()


                End If

                cantidadLotes = 0


            End While
        Catch oe As Exception
            MsgBox(oe.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub crearGoodReceiptsSTOOLD()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String


        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand


        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_STO.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try



        Try

            Dim linea As String

            conEx.ConnectionString = conexionString
            conEx.Open()
            cmdEx.Connection = conEx
            cmdEx.CommandText = "SELECT * FROM CABECERA_STO_ADVICE WHERE GENERADO='N'"
            Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()

            While lrdEx.Read()

                conIn.ConnectionString = conexionString
                conIn.Open()

                docnum = lrdEx.GetString("0")

                Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombrePartner & "_ZWHGRC_" & obtenerFechaHora() & ".xml")

                linea = ""
                linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
                linea = linea & "<ZWMMBID2>" & vbNewLine
                linea = linea & "<IDOC BEGIN=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                linea = linea & "<EDI_DC40 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                linea = linea & "<TABNAM>EDI_DC40</TABNAM>" & vbNewLine
                linea = linea & "<MANDT>600</MANDT>" & vbNewLine
                linea = linea & "<DOCNUM>" & obtenerNextNumber() & "</DOCNUM>" & vbNewLine
                linea = linea & "<DOCREL>701</DOCREL>" & vbNewLine
                linea = linea & "<STATUS>30</STATUS>" & vbNewLine
                linea = linea & "<DIRECT>1</DIRECT>" & vbNewLine
                linea = linea & "<OUTMOD>2</OUTMOD>" & vbNewLine
                linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>" & vbNewLine
                linea = linea & "<CIMTYP>ZWMMBID2</CIMTYP>" & vbNewLine
                linea = linea & "<MESTYP>ZWHGRC</MESTYP>" & vbNewLine
                linea = linea & "<STDMES>ZWHGRC</STDMES>" & vbNewLine
                linea = linea & "<SNDPOR>SAPGU3</SNDPOR>" & vbNewLine
                linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
                linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>" & vbNewLine
                linea = linea & "<RCVPOR>PI_XML</RCVPOR>" & vbNewLine
                linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
                linea = linea & "<RCVPRN>STPG</RCVPRN>" & vbNewLine
                linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>" & vbNewLine
                linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>" & vbNewLine
                linea = linea & "</EDI_DC40>" & vbNewLine
                linea = linea & "<E1MBXYH SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                linea = linea & "<BUDAT>" & obtenerFecha() & "</BUDAT>" & vbNewLine
                linea = linea & "<ZGREC01 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                linea = linea & "<DOC_TYP>PO</DOC_TYP>" & vbNewLine
                linea = linea & "</ZGREC01>" & vbNewLine
                linea = linea & "<ZSTKDATE SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                linea = linea & "<DATE>" & obtenerFecha() & "</DATE>" & vbNewLine
                linea = linea & "<TIME>" & obtenerHora() & "</TIME>" & vbNewLine
                linea = linea & "</ZSTKDATE>" & vbNewLine

                Dim cmdIn As New SqlCommand
                cmdIn.Connection = conIn
                cmdIn.CommandText = "SELECT * FROM DETALLE_STO_ADVICE WHERE DOCNUM='" & lrdEx.GetString(0) & "'"
                Dim lrdIn As SqlDataReader = cmdIn.ExecuteReader()

                While lrdIn.Read()

                    linea = linea & "<E1MBXYI SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine
                    linea = linea & "<MATNR>" & lrdIn.GetString(2) & "</MATNR>" & vbNewLine
                    linea = linea & "<WERKS>" & lrdEx.GetString(5) & "</WERKS>" & vbNewLine
                    linea = linea & "<LGORT>" & lrdIn.GetString(6) & "</LGORT>" & vbNewLine
                    linea = linea & "<CHARG></CHARG>" & vbNewLine
                    linea = linea & "<ERFMG>" & lrdIn.GetString(8) & "</ERFMG>" & vbNewLine
                    linea = linea & "<ERFME>" & lrdIn.GetString(9) & "</ERFME>" & vbNewLine
                    linea = linea & "<EBELN>" & lrdIn.GetString(13) & "</EBELN>" & vbNewLine
                    linea = linea & "<EBELP>" & lrdIn.GetString(14) & "</EBELP>" & vbNewLine
                    linea = linea & "<ZGREC02 SEGMENT=" & Chr(34) & "1" & Chr(34) & ">" & vbNewLine

                    If lrdIn.GetString(6).Substring(lrdIn.GetString(6).Length - 2, 2) = "10" Or lrdIn.GetString(6).Substring(lrdIn.GetString(6).Length - 2, 2) = "11" Then
                        linea = linea & "<DUTY_STAT>P</DUTY_STAT>" & vbNewLine
                    Else
                        linea = linea & "<DUTY_STAT>S</DUTY_STAT>" & vbNewLine
                    End If

                    linea = linea & "<STOCK_STAT>FREE</STOCK_STAT>" & vbNewLine
                    linea = linea & "<SPECIAL>1</SPECIAL>" & vbNewLine
                    linea = linea & "<USRTXT2>" & lrdIn.GetString(13) & "</USRTXT2>" & vbNewLine
                    linea = linea & "</ZGREC02>" & vbNewLine
                    linea = linea & "</E1MBXYI>" & vbNewLine

                End While

                linea = linea & "</E1MBXYH>" & vbNewLine
                linea = linea & "</IDOC>" & vbNewLine
                linea = linea & "</ZWMMBID2>"
                oSW.WriteLine(linea)
                oSW.Flush()
                oSW.Close()

                actualizarGenerado("CABECERA_STO_ADVICE", docnum)

                System.Threading.Thread.Sleep(2000)

                conIn.Close()


            End While

            conEx.Close()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


    End Sub


    Public Sub escribirLog(ByVal mensaje As String, ByVal proceso As String)

        Dim time As DateTime = DateTime.Now
        Dim format As String = "dd/MM/yyyy HH:mm "

        lineaLogger = proceso & time.ToString(format) & ":" & mensaje & vbNewLine
        logger.WriteLine(lineaLogger)
        logger.Flush()

    End Sub

    Private Function obtenerFecha() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyyMMdd"
        Return time.ToString(format)

    End Function

    Private Function obtenerFechaHora() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyyMMddHHmmss"
        Return time.ToString(format)

    End Function

    Private Function obtenerHora() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "HHmmss"
        Return time.ToString(format)

    End Function

    Private Function obtenerHoraPedido() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        Return time.ToString(format)

    End Function




    Public Function obtenerNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As Dictionary(Of String, String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
        Next
        Return diccionario
    End Function


    Public Function obtenerNodosHijosDePadreLista(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As List(Of Dictionary(Of String, String))

        Dim listaDiccionario As New List(Of Dictionary(Of String, String))
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            diccionario = New Dictionary(Of String, String)
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
            listaDiccionario.Add(diccionario)
        Next
        Return listaDiccionario

    End Function

    Public Function obtenerListaNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As List(Of Nodo)

        Dim listaNodos As New List(Of Nodo)
        Dim nodo As Nodo
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                nodo = New Nodo()
                nodo.sName = nodoPadre(i).ChildNodes.Item(h).Name.Trim()
                nodo.sInner = nodoPadre(i).ChildNodes.Item(h).InnerText.Trim()
                listaNodos.Add(nodo)
            Next
        Next
        Return listaNodos

    End Function


    Private Function buscarNodo(ByVal name As String, ByVal listaNodos As List(Of Nodo)) As Nodo

        Dim nodo As Nodo = New Nodo()
        Dim encontrado As Boolean
        Dim nodos_enumerator As IEnumerator
        nodos_enumerator = listaNodos.GetEnumerator()
        encontrado = False

        nodo.sName = "NULL"

        Do While (nodos_enumerator.MoveNext) And Not encontrado
            nodo = CType(nodos_enumerator.Current, Nodo)
            If nodo.sName.CompareTo(name) = 0 Then
                encontrado = True
            Else
                nodo.sName = "NULL"
            End If

        Loop

        buscarNodo = nodo

    End Function

    Private Function obtenerNombreArchivo(ByVal directorio As String, ByVal nombreBase As String) As List(Of String)


        Dim listaArchivos As New List(Of String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim strFileSize As String = ""
        Dim nombreArchivo As String
        nombreArchivo = ""

        Try
            Dim di As New IO.DirectoryInfo(directorio)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
            Dim fi As IO.FileInfo

            For Each fi In aryFi
                If InStr(fi.Name, nombreBase) > 0 Then
                    diccionario = New Dictionary(Of String, String)
                    nombreArchivo = fi.Name
                    nombreArchivo.Concat(".xml")

                    listaArchivos.Add(nombreArchivo)
                    'Exit For
                Else
                    nombreArchivo = ""
                End If

            Next

        Catch ex As Exception

            lineaLogger = "Línea de texto " & vbNewLine & "Otra linea de texto"
            logger.WriteLine(lineaLogger)
            logger.Flush()

        Finally

        End Try
        Return listaArchivos

    End Function


    Private Sub crearInventoryMovements()

        Try
            Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\DUSA_ZWHINV_20120726122345.xml")

            Dim linea As String
            linea = ""


            linea = linea & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "no" & Chr(34) & " ?>" & vbNewLine
            linea = linea & "<ZWMMBID2>" & vbNewLine
            linea = linea & "<IDOC BEGIN=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<EDI_DC40>" & vbNewLine
            linea = linea & "<TABNAM>EDI_DC40</TABNAM>" & vbNewLine
            linea = linea & "<MANDT>600</MANDT>" & vbNewLine
            linea = linea & "<DOCNUM>00000000000076</DOCNUM>" & vbNewLine
            linea = linea & "<DIRECT />" & vbNewLine
            linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>" & vbNewLine
            linea = linea & "<CIMTYP>ZWMMBID2</CIMTYP>" & vbNewLine
            linea = linea & "<MESTYP>ZWHINV</MESTYP>" & vbNewLine
            linea = linea & "<SNDPOR />" & vbNewLine
            linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
            linea = linea & "<SNDPRN>DUSA</SNDPRN>" & vbNewLine
            linea = linea & "<RCVPOR />" & vbNewLine
            linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
            linea = linea & "<RCVPRN></RCVPRN>" & vbNewLine
            linea = linea & "<CREDAT>20111227</CREDAT>" & vbNewLine
            linea = linea & "<CRETIM>183131</CRETIM>" & vbNewLine
            linea = linea & "</EDI_DC40>" & vbNewLine
            linea = linea & "<E1MBXYH SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<BKTXT>00000000000076</BKTXT>" & vbNewLine
            linea = linea & "</E1MBXYH>" & vbNewLine
            linea = linea & "<ZGREC01 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<DOC_TYP>X</DOC_TYP>" & vbNewLine
            linea = linea & "</ZGREC01>" & vbNewLine
            linea = linea & "<ZSTKDATE SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<DATE>20111227</DATE>" & vbNewLine
            linea = linea & "<TIME>183131</TIME>" & vbNewLine
            linea = linea & "</ZSTKDATE>" & vbNewLine
            linea = linea & "<E1MBXYI SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<MATNR>000000000000618693</MATNR>" & vbNewLine
            linea = linea & "<WERKS>CF1</WERKS>" & vbNewLine
            linea = linea & "<LGORT>5011</LGORT>" & vbNewLine
            linea = linea & "<CHARG>L2386-PC</CHARG>" & vbNewLine
            linea = linea & "<ERFMG>000000000000025</ERFMG>" & vbNewLine
            linea = linea & "<ERFME>CX</ERFME>" & vbNewLine
            linea = linea & "<UMLGO>5012</UMLGO>" & vbNewLine
            linea = linea & "<UMCHA>TESTR19</UMCHA>" & vbNewLine
            linea = linea & "<ZINVMV1 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<FROM_STOCK_STAT>QUAR</FROM_STOCK_STAT>" & vbNewLine
            linea = linea & "<MESSAGE_TYPE>3</MESSAGE_TYPE>" & vbNewLine
            linea = linea & "</ZINVMV1>" & vbNewLine
            linea = linea & "<ZGREC02 SEGMENT=" & Chr(34) & "" & Chr(34) & ">" & vbNewLine
            linea = linea & "<DUTY_STATUS>P</DUTY_STATUS>" & vbNewLine
            linea = linea & "<STOCK_STAT>FREE</STOCK_STAT>" & vbNewLine
            linea = linea & "</ZGREC02>" & vbNewLine
            linea = linea & "</E1MBXYI>" & vbNewLine
            linea = linea & "</IDOC>" & vbNewLine
            linea = linea & "</ZWMMBID2>"


            oSW.WriteLine(linea)
            oSW.Flush()
            oSW.Close()


        Catch oe As Exception
            MsgBox(oe.Message, MsgBoxStyle.Critical)
        End Try

    End Sub






    Private Sub crearStockReconciliation(ByVal planta As String)


        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim docnum As String

        Try

            nombrePartner = "DUSA"

            llenarStockReconciliation(planta)

            Dim linea As String

            conIn.ConnectionString = conexionString
            conIn.Open()



            Dim oSW As New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\" & nombrePartner & "_ZWHSRI_" & obtenerFechaHora() & ".xml")

            linea = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & " ?>" & vbNewLine
            linea = linea & "<ZWMMBID1>" & vbNewLine
            linea = linea & "<IDOC>" & vbNewLine
            linea = linea & "<EDI_DC40>" & vbNewLine
            linea = linea & "<TABNAM>EDI_DC40</TABNAM>" & vbNewLine
            linea = linea & "<MANDT>600</MANDT>" & vbNewLine
            linea = linea & "<DOCNUM>" & obtenerNextNumber() & "</DOCNUM>" & vbNewLine
            linea = linea & "<DIRECT />" & vbNewLine
            linea = linea & "<IDOCTYP>WMMBID01</IDOCTYP>" & vbNewLine
            linea = linea & "<CIMTYP>ZWMMBID1</CIMTYP>" & vbNewLine
            linea = linea & "<MESTYP>ZWHSRI</MESTYP>" & vbNewLine
            linea = linea & "<SNDPOR />" & vbNewLine
            linea = linea & "<SNDPRT>LS</SNDPRT>" & vbNewLine
            linea = linea & "<SNDPRN>" & nombrePartner & "</SNDPRN>" & vbNewLine
            linea = linea & "<RCVPOR />" & vbNewLine
            linea = linea & "<RCVPRT>LS</RCVPRT>" & vbNewLine
            linea = linea & "<RCVPRN>GD1600</RCVPRN>" & vbNewLine
            linea = linea & "<CREDAT>" & obtenerFecha() & "</CREDAT>" & vbNewLine
            linea = linea & "<CRETIM>" & obtenerHora() & "</CRETIM>" & vbNewLine
            linea = linea & "</EDI_DC40>" & vbNewLine
            linea = linea & "<E1MBXYH>" & vbNewLine
            linea = linea & "<BLDAT>" & obtenerFecha() & "</BLDAT>" & vbNewLine
            linea = linea & "<ZSTKDATE>" & vbNewLine
            linea = linea & "<DATE>" & obtenerFecha() & "</DATE>" & vbNewLine
            linea = linea & "<TIME>" & obtenerHora() & "</TIME>" & vbNewLine
            linea = linea & "</ZSTKDATE>" & vbNewLine

            Dim cmdIn As New SqlCommand
            cmdIn.Connection = conIn
            cmdIn.CommandText = "SELECT * FROM STOCK_FILE"
            Dim lrdIn As SqlDataReader = cmdIn.ExecuteReader()

            While lrdIn.Read()

                linea = linea & "<E1MBXYI>" & vbNewLine
                linea = linea & "<MATNR>" & Right("000000000000000000" & Trim(lrdIn.GetString("0")), 18) & "</MATNR>" & vbNewLine
                linea = linea & "<WERKS>" & Trim(lrdIn.GetString("3")) & "</WERKS>" & vbNewLine
                linea = linea & "<LGORT>" & Trim(lrdIn.GetString("4")) & "</LGORT>" & vbNewLine
                linea = linea & "<CHARG>" & Trim(lrdIn.GetString("2")) & "</CHARG>" & vbNewLine
                linea = linea & "<Z1WSTK1>" & vbNewLine
                linea = linea & "<FREE>" & Trim(lrdIn.GetDouble("6")).Replace(",", ".") & "</FREE>" & vbNewLine
                linea = linea & "<QUAR>" & Trim(lrdIn.GetDouble("7")).Replace(",", ".") & "</QUAR>" & vbNewLine
                linea = linea & "<RETURN>" & Trim(lrdIn.GetDouble("8")).Replace(",", ".") & "</RETURN>" & vbNewLine
                linea = linea & "<HELD>" & Trim(lrdIn.GetDouble("9")).Replace(",", ".") & "</HELD>" & vbNewLine
                linea = linea & "<TOTAL>" & Trim(lrdIn.GetDouble("10")).Replace(",", ".") & "</TOTAL>" & vbNewLine

                If Trim(lrdIn.GetString("4")).Substring(Trim(lrdIn.GetString("4")).Length - 2, 2) = "10" Or Trim(lrdIn.GetString("4")).Substring(Trim(lrdIn.GetString("4")).Length - 2, 2) = "11" Then
                    linea = linea & "<DUTYSTAT>P</DUTYSTAT>" & vbNewLine
                Else
                    linea = linea & "<DUTYSTAT>S</DUTYSTAT>" & vbNewLine
                End If

                linea = linea & "<VENDOR>000000000000000</VENDOR>" & vbNewLine
                linea = linea & "</Z1WSTK1>" & vbNewLine
                linea = linea & "</E1MBXYI>" & vbNewLine

            End While

            linea = linea & "</E1MBXYH>" & vbNewLine
            linea = linea & "</IDOC>" & vbNewLine
            linea = linea & "</ZWMMBID1>"
            oSW.WriteLine(linea)
            oSW.Flush()
            oSW.Close()

            System.Threading.Thread.Sleep(2000)

            conIn.Close()



        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try


    End Sub



    Private Function llenarStockReconciliation(ByVal planta1 As String)


        Dim conInsertar As New SqlConnection
        Dim free As Double
        Dim quar As Double
        Dim returnn As Double
        Dim held As Double
        Dim total As Double
        Dim planta As String
        Dim storage As String
        Dim uom As String
        Dim codigo As String
        Dim lote As String

        Dim host As String
        Dim host_pedidos As String
        Dim database As String
        Dim database_pedidos As String
        Dim user As String
        Dim password As String
        Dim user_pedidos As String
        Dim password_pedidos As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String




        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion_SRF.xml", FileMode.Open, FileAccess.Read)
            xmldoc = New XmlDataDocument()
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")

            prefijo = diccionario.Item("prefijo_archivo")
            host_pedidos = diccionario.Item("host_pedidos")
            database_pedidos = diccionario.Item("database_pedidos")
            user_pedidos = diccionario.Item("user_pedidos")
            password_pedidos = diccionario.Item("password_pedidos")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            conexionStringPedidos = "Data Source=" & host_pedidos & ";Database=" & database_pedidos & ";User ID=" & user_pedidos & ";Password=" & password_pedidos & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

            escribirLog(conexionString, "(OUTB) ")
        Catch oe As Exception
            escribirLog(oe.Message.ToString, "(OUTB) ")
        Finally

            logger.Close()
        End Try

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("SELECT * FROM F55DE2 WHERE DRAA04='" & planta1 & "'", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As String
        valor = 0

        Try
            conInsertar.ConnectionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            conInsertar.Open()
            Dim sqlstring As String
            sqlstring = " delete from [dusa_dorado].[dbo].[STOCK_FILE]"
            comando.Connection = conInsertar
            comando.CommandText = " "
            comando.CommandText = sqlstring
            comando.ExecuteNonQuery()
            comando.Parameters.Clear()
            conInsertar.Close()


            cnn.Open()

            reader = rs.ExecuteReader
            While reader.Read()
                codigo = Trim(reader("DRAA18"))
                planta = Trim(reader("DRAA04"))
                storage = Trim(reader("DRCDCTYPE"))
                free = Trim(reader("DRPQOH")) / 10000
                quar = Trim(reader("DRFCOM")) / 10000
                returnn = Trim(reader("DRSCMS")) / 10000
                held = Trim(reader("DRPCOM")) / 10000
                'uom = buscarUOMSRF(Trim(reader("DRUOM1")))
                uom = Trim(reader("DRUOM1"))
                lote = Right(Trim(reader("DRLOTN")), 10)

                conInsertar.Open()

                total = free + quar + returnn + held
                sqlstring = " INSERT INTO [dusa_dorado].[dbo].[STOCK_FILE]([codigo],[descripcion],[lote],[planta],[storage],[uom],[free],[quar],[retur],[held],[total]) VALUES ('" & codigo & "','','" & lote & "','" & planta & "','" & storage & "','" & uom & "'," & Replace(free, ",", ".") & "," & Replace(quar, ",", ".") & "," & Replace(returnn, ",", ".") & "," & Replace(held, ",", ".") & "," & Replace(total, ",", ".") & ") "
                comando.Connection = conInsertar
                comando.CommandText = " "
                comando.CommandText = sqlstring
                comando.ExecuteNonQuery()
                comando.Parameters.Clear()

                free = 0
                quar = 0
                returnn = 0
                held = 0
                total = 0

                conInsertar.Close()


            End While

            cnn.Close()

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally


        End Try







    End Function




    Private Sub leerXML()
        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer
        Dim str As String
        Dim fs As New FileStream("products.xml", FileMode.Open, FileAccess.Read)

        xmldoc.Load(fs)
        xmlnode = xmldoc.GetElementsByTagName("Product")
        For i = 0 To xmlnode.Count - 1
            'xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
            str = xmlnode(i).ChildNodes.Item(0).Name.Trim() & " | " & xmlnode(i).ChildNodes.Item(1).InnerText.Trim() & " | " & xmlnode(i).ChildNodes.Item(2).InnerText.Trim()
            MsgBox(str)
        Next
    End Sub



    Private Sub crearXML()
        Dim writer As New XmlTextWriter("product.xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("Table")
        createNode(1, "Product 1", "1000", writer)
        createNode(2, "Product 2", "2000", writer)
        createNode(3, "Product 3", "3000", writer)
        createNode(4, "Product 4", "4000", writer)
        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub


    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal pPrice As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Product")
        writer.WriteStartElement("Product_id")
        writer.WriteString(pID)
        writer.WriteEndElement()
        writer.WriteStartElement("Product_name")
        writer.WriteString(pName)
        writer.WriteEndElement()
        writer.WriteStartElement("Product_price")
        writer.WriteString(pPrice)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub buscar()
        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create("Product.xml", New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)

        dv = New DataView(ds.Tables(0))
        dv.Sort = "Product_Name"
        Dim index As Integer = dv.Find("Product 2")

        If index = -1 Then
            MsgBox("Item Not Found")
        Else
            MsgBox(dv(index)("Product_Name").ToString() & "  " & dv(index)("Product_Price").ToString())
        End If
    End Sub

    Private Sub filtrar()

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create("Product.xml", New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)
        dv = New DataView(ds.Tables(0), "Product_price > = 3000", "Product_Name", DataViewRowState.CurrentRows)
        dv.ToTable().WriteXml("Result.xml")
        MsgBox("Done")
    End Sub

    Private Sub buscarPorTag()


        ' Open the XML file
        Dim xmlDocContinents As New XmlDocument
        xmlDocContinents.Load("Product.xml")

        ' Get a list of elements whose names are Continent
        Dim lstContinents As XmlNodeList = xmlDocContinents.GetElementsByTagName("Product")

        ' Retrieve the name of each continent and put it in the combo box
        Dim i As Integer
        For i = 0 To lstContinents.Count Step 1

            MsgBox(lstContinents(i).Attributes("Product_name").InnerText)

        Next
    End Sub



End Module
