'----**** 
'----**** MyBusiness POS V20
'----**** Version del script: 1.0
'----**** 19/02/2020
'----**** 
Sub Main()    
                                               
    If EsArticuloParaTiempoAire Then
		CancelaProceso = True
		txtFields(4).SetFocus()
		Exit Sub
	ElseIf EsArticuloOtrosServicios Then
		CancelaProceso = True
		txtFields(4).SetFocus()
		Exit Sub                         
	ElseIf EsArticuloParaPagoServicios Then
		CancelaProceso = True
		txtFields(4).SetFocus()
		Exit Sub
	End If
     
    'Call incrementaProducto()
                       
    ' Articulo es una variable que entrega el dato que se capturo en
    ' el punto de venta
    If Trim( Ucase(Articulo) ) = "ADMISION" Then 

       If Me.Venta = 0 Then
          MyMessage "Es necesario capturar al menos de un producto"
          CancelaProceso = True
          Exit Sub
       End If
 
       Script.RunForm "MIFORMA", Me, Ambiente,  , True
       'PlaySound Ambiente.Path & "\sounds\s03.wav"       
       CancelaProceso = True
    End If
        
    Me.usuarioRequerido = 0
                  
    If clAt( "V", Articulo ) = 1 Then

		If Not ExisteArticuloTiempoAireOEspecial Then

           nVenta = Val2( Mid( Articulo, 2 ) ) 
           If nVenta = 0 Then
              Exit Sub
           End If

 		   Call recuperaVenta()
		   CancelaProceso = True

   		End If            
                 
    End If                 

    If Ucase(Trim(Articulo)) = "EXR" Then
       CancelaProceso = True
       Script.RunForm "EXISTENCIAREMOTA", Me, Ambiente,, True
    End If                   

    Version2005     

    'Call calculaPrecioDecaja()

End Sub  
       

Function EsArticuloOtrosServicios()
	Dim sCadena

	EsArticuloOtrosServicios = False	

	If UCase(Trim(Articulo)) = "RSATFEMYB" Then
		sCadena = InformacionDeSKUs_OS

		If sCadena = "" Then
        	Exit Function
		End If

    	CreaProducto UCase(Trim(Articulo)), "RECARGA DE TIMBRES MYCFDI"
		VentaOtrosServicios UCase(Trim(Articulo)), UCase(Mid(Trim(Articulo), 2)), sCadena
		CancelaProceso = True
		EsArticuloOtrosServicios = True
	End If
End Function
            

Function InformacionDeSKUs_OS()
	Dim Informacion
        
	On Error Resume Next
	Set Informacion = CreateObject("Servicios.Servicios")
	If Err.Number <> 0 Then
    	InformacionDeSKUs_OS = ""
	Else
		InformacionDeSKUs_OS = Informacion.LlamaAlServicio("SKUS_OS")
	End If        
End Function


Sub VentaOtrosServicios(articulo_OS, proveedor_OS, sCadena)
	Select Case articulo_OS
           Case "RSATFEMYB" 
				VentaTimbresMyCFDI articulo_OS, proveedor_OS, sCadena
	End Select
End Sub
          

Sub VentaTimbresMyCFDI(articulo_OS, proveedor_OS, sCadena)
	Dim TimbresMyCFDI, sMensaje, rstArt, iMonto                       
                
	On Error Resume Next
	Set TimbresMyCFDI = CreateObject("Servicios.Servicios")
	sMensaje = TimbresMyCFDI.LlamaAlServicio("RECARGATIMBRES", articulo_OS, (Me.Venta), Ambiente.Uid)        
	                       
	If Err.Number <> 0 Then
	    MyMessage "Verifique que tenga instalado Servicios TAE en su equipo"  & vbcrLf & Err.Number & ": " & Err.Description
	ElseIf ClAt("<Code>0</Code>", sMensaje) > 0 Then
		Set rstArt = Rst("SELECT prods.*, impuestos.valor FROM prods, impuestos WHERE prods.impuesto = impuestos.impuesto " & _
						 "AND prods.articulo = '" & articulo_OS & "'", Ambiente.Connection)
	   
		iMonto = Val2(Monto(sCadena, DatoXML((sMensaje), "Sku")))
		NuevaPartida rstArt, iMonto / (1 + (Val2(rstArt("valor"))/100))
		GuardaDatos_OS sMensaje, proveedor_OS, iMonto
	ElseIf ClAt("<Code>", sMensaje) > 0 Then
		MyMessage("Error: " & DatoXML((sMensaje), "Code") & vbCrLf & DatoXML((sMensaje), "DCode"))
	Else
		MyMessage(sMensaje)
	End If  
End Sub         


Sub GuardaDatos_OS(Cadena, proveedorTAE, iMonto)
	Dim Query, id

	Set Query = NewQuery()
	Set Query.Connection = Ambiente.Connection

	id = TraeSiguiente("recargastimbres", Ambiente.Connection)          

	Query.Reset
	Query.strState = "INSERT"
	Query.AddField "recargastimbres","ID", id
    Query.AddField "recargastimbres","SKU", DatoXML((Cadena), "Sku")
    Query.AddField "recargastimbres","RFC", DatoXML((Cadena), "RFC")
    Query.AddField "recargastimbres","monto", iMonto 
    Query.AddField "recargastimbres","folios", Val2(Replace(DatoXML((Cadena), "Sku"), "RSATFE", ""))
	Query.AddField "recargastimbres","serie", DatoXML((Cadena), "Serie")
    Query.AddField "recargastimbres","venta", Me.Venta
    Query.AddField "recargastimbres","usuario", Ambiente.Uid
    Query.AddField "recargastimbres","usuFecha", Date
    Query.AddField "recargastimbres","usuHora", Formato(Time, "HH:mm:ss")
    Query.AddField "recargastimbres","codigoRespuesta", DatoXML((Cadena), "Code")
    Query.AddField "recargastimbres","descripCodigoRespuesta", DatoXML((Cadena), "DCode")
    Query.AddField "recargastimbres","folioCarrier", DatoXML((Cadena), "IDTrans")
    Query.AddField "recargastimbres","xmlDeRespuesta", sCadena
    Query.Exec            

	Script.ImprimeFormato "TICKETTIMBRES", (id), (Ambiente), Me, False
End Sub
                

''Inicia Tiempo Aire
Function EsArticuloParaTiempoAire()

	EsArticuloParaTiempoAire = False	

	If UCase(Trim(Articulo)) = "RTELCEL" OR UCase(Trim(Articulo)) = "RMOVISTAR" OR UCase(Trim(Articulo)) = "RIUSACELL" OR _
	   UCase(Trim(Articulo)) = "RUNEFON" OR UCase(Trim(Articulo)) = "RNEXTEL" OR UCase(Trim(Articulo)) = "RVIRGIN" OR _
	   UCase(Trim(Articulo)) = "RTELCELINT" OR UCase(Trim(Articulo)) = "RTELCELPAQ" OR UCase(Trim(Articulo)) = "RALO" Then

    	CreaProducto UCase(Trim(Articulo)), "RECARGA " & Mid(UCase(Trim(Articulo)), 2)

		VentaTiempoAire UCase(Trim(Articulo))

		CancelaProceso = True
		EsArticuloParaTiempoAire = True

	End If         

End Function                                                               


''Pago de servicios
Function EsArticuloParaPagoServicios()

	EsArticuloParaPagoServicios = False

	If Ucase(Trim(Articulo)) = "RSERVICIOS" Then
                                                                         
		CreaProducto Ucase(Trim(Articulo)), "Pago de servicios"
		PagoDeServicios Ucase(Trim(Articulo))

		CancelaProceso = True
		EsArticuloParaPagoServicios = True

	End If

End Function

                
Sub CreaProducto(articuloTAE, descripcionTAE)
	Dim rstP, rstImp

	Set rstP = Rst("SELECT articulo FROM prods WHERE articulo = '" & articuloTAE & "'", Ambiente.Connection)
	Set rstImp = Rst("SELECT valor FROM impuestos WHERE impuesto = 'IVA'", Ambiente.Connection)    

	If rstP.EOF Then
    	Dim Query

		Set Query = NewQuery()
		Set Query.Connection = Ambiente.Connection
                          
		Query.Reset
		Query.strState = "INSERT"
             
		Query.AddField "prods","ARTICULO", articuloTAE
     	Query.AddField "prods","DESCRIP", descripcionTAE
     	Query.AddField "prods","LINEA", "SYS"
     	Query.AddField "prods","MARCA", "SYS"

		If Val2(Ambiente.rstEstacion("conimpuesto")) <> 0 Then
     		Query.AddField "prods","PRECIO1", (1/(1 + (Val2(rstImp("valor"))/100)))
			Query.AddField "prods","IMPUESTO", "IVA"
		Else
        	Query.AddField "prods","PRECIO1", 1
			Query.AddField "prods","IMPUESTO", "SYS"
		End If

     	Query.AddField "prods","PRECIO2", 0
     	Query.AddField "prods","PRECIO3", 0
     	Query.AddField "prods","PRECIO4", 0
     	Query.AddField "prods","PRECIO5", 0
     	Query.AddField "prods","PRECIO6", 0
     	Query.AddField "prods","PRECIO7", 0
     	Query.AddField "prods","PRECIO8", 0
     	Query.AddField "prods","PRECIO9", 0
    	Query.AddField "prods","PRECIO10", 0
     	Query.AddField "prods","PRECIOUSD", 0
     	Query.AddField "prods","EXISTENCIA", 0
     	Query.AddField "prods","COSTO_U", 0
     	Query.AddField "prods","COSTOUSD", 0 
		If Val2(Ambiente.rstEstacion("conimpuesto")) <> 0 Then
     		Query.AddField "prods","COSTO", (1/(1 + (Val2(rstImp("valor"))/100)))
		Else
        	Query.AddField "prods","COSTO", 1
		End If
     	  
     	Query.AddField "prods","KIT", 0
     	Query.AddField "prods","SERIE", 0
     	Query.AddField "prods","LOTE", 0 
     	Query.AddField "prods","INVENT", 1
     	Query.AddField "prods","IMAGEN", ""
     	Query.AddField "prods","PARAVENTA", 1
     	Query.AddField "prods","URL", "" 
     	Query.AddField "prods","USUARIO", Ambiente.Uid
     	Query.AddField "prods","USUHORA", Formato(Time, "HH:mm:ss")
     	Query.AddField "prods","USUFECHA", Date 
     	Query.AddField "prods","Granel", 0 
     	Query.Exec	

		Query.Reset
		Query.strState = "INSERT"
		Query.AddField "existenciaalmacen","almacen", 1
     	Query.AddField "existenciaalmacen","articulo", articuloTAE
     	Query.AddField "existenciaalmacen","existencia", 0
		Query.Exec

	End If

	rstP.Close
	Set rstP = Nothing

	rstImp.Close
	Set rstImp = Nothing
End Sub                                                


Sub PagodeServicios(Servicio)
                  
	Set PSrv = CreateObject("TiempoAire.Servicios")

	sMensaje = PSrv.LlamaAlServicio("PAGOSERVICIOS", Servicio, Me.Venta, Ambiente.UId)

	If Err.Number <> 0 Then

	    MyMessage "Verifique que tenga instalado Servicios TAE en su equipo" & vbcrLf & Err.Number & ": " & Err.Description

	End If

	If clAt("errorcode", sMensaje) > 0 Then

		If DatosXML(sMensaje, "errorcode") = 0 Then

			nMonto = DatosXML(sMensaje, "monto")

			Set rstArt = Rst("SELECT prods.*, impuestos.valor FROM prods, impuestos WHERE prods.impuesto = impuestos.impuesto " & _
							 "AND prods.articulo = '" & Servicio & "'", Ambiente.Connection)

			NuevaPartida rstArt, nMonto / (1 + (Val2(rstArt("valor"))/100))

			GuardaDatosSRV sMensaje, nMonto

	    ElseIf DatosXML(sMensaje, "errorcode") > 0 Then

			MyMessage("Error: " & DatosXML(sMensaje, "errorcode") & vbCrLf & DatosXML(sMensaje, "responsemessage"))

		Else
                                            
  			MyMessage("Error: " & sMensaje)                    

		End If

	Else

		If Trim(sMensaje) <> "Operación cancelada por el usuario" Then
			MyMessage("Error: " & sMensaje)
		End If

		CancelaProceso = True

	End If

End Sub
                     

Sub VentaTiempoAire(articuloTAE)
	Dim TiempoAire, sMensaje, rstArt, nMonto                       

	On Error Resume Next
	Set TiempoAire = CreateObject("TiempoAire.Servicios")
                          
	sMensaje = TiempoAire.LlamaAlServicio("RECARGA", Mid(articuloTAE, 2), (Me.Venta), Ambiente.Uid)      


	If Err.Number <> 0 Then
                              
	    MyMessage "Verifique que tenga instalado Servicios TAE en su equipo" & vbcrLf & Err.Number & ": " & Err.Description

	End If
                              
	If clAt("errorcode", sMensaje) > 0 Then

		If DatosXML(sMensaje, "errorcode") = 0 Then

			nMonto = DatosXML(sMensaje, "monto")

			Set rstArt = Rst("SELECT prods.*, impuestos.valor FROM prods, impuestos WHERE prods.impuesto = impuestos.impuesto " & _
							 "AND prods.articulo = '" & articuloTAE & "'", Ambiente.Connection)

			NuevaPartida rstArt, nMonto / (1 + (Val2(rstArt("valor"))/100))

			GuardaDatos sMensaje, nMonto   

	    ElseIf DatosXML(sMensaje, "errorcode") > 0 Then
                                           
			MyMessage("Error: " & DatosXML(sMensaje, "errorcode") & vbCrLf & DatosXML(sMensaje, "responsemessage"))     

		Else

      		MyMessage("Error: " & sMensaje)

		End If   
	
	Else
                                       
		If Trim(sMensaje) <> "Operación cancelada por el usuario" Then
			MyMessage("Error: " & sMensaje)
		End If

		CancelaProceso = True

	End If

End Sub                              


Sub GuardaDatos(Cadena, nMonto)   
	Dim Query, id

	Set Query = NewQuery()
	Set Query.Connection = Ambiente.Connection

	id = TraeSiguiente("recargastae", Ambiente.Connection)          

	Query.Reset
	Query.strState = "INSERT"
	Query.AddField "recargastae","id", id
    Query.AddField "recargastae","SKU", DatosXML(Cadena, "sku")
    Query.AddField "recargastae","telefono", DatosXML(Cadena, "numero")
    Query.AddField "recargastae","proveedor", DatosXML(Cadena, "carrier")
    Query.AddField "recargastae","monto", nMonto
    Query.AddField "recargastae","venta", Me.Venta
    Query.AddField "recargastae","usuario", Ambiente.Uid
    Query.AddField "recargastae","usufecha", Date
    Query.AddField "recargastae","usuhora", Formato(Time, "HH:mm:ss")
    Query.AddField "recargastae","codigoRespuesta", DatosXML(Cadena, "errorcode")
    Query.AddField "recargastae","descripCodigoRespuesta", DatosXML(Cadena, "responsemessage")
    Query.AddField "recargastae","folioCarrier", DatosXML(Cadena, "authcode")
    Query.AddField "recargastae","xmlDeRespuesta", Cadena
	Query.Exec

	Script.ImprimeFormato "TICKETTAE", (id), (Ambiente), Me, False 

End Sub


Sub GuardaDatosSRV(Cadena, nMonto)   
	Dim Query, id

	Set Query = NewQuery()
	Set Query.Connection = Ambiente.Connection

	id = TraeSiguiente("pagoservicios", Ambiente.Connection)          

	Query.Reset
	Query.strState = "INSERT"
	Query.AddField "pagodeservicios", "id", id
    Query.AddField "pagodeservicios", "SKU", DatosXML(Cadena, "sku")
    Query.AddField "pagodeservicios", "referencia", DatosXML(Cadena, "numero")
    Query.AddField "pagodeservicios", "servicio", DatosXML(Cadena, "carrier")
    Query.AddField "pagodeservicios", "monto", nMonto
    Query.AddField "pagodeservicios", "venta", Me.Venta
    Query.AddField "pagodeservicios", "usuario", Ambiente.Uid
    Query.AddField "pagodeservicios", "usufecha", Date
    Query.AddField "pagodeservicios", "usuhora", Formato(Time, "HH:mm:ss")
    Query.AddField "pagodeservicios", "codigoRespuesta", DatosXML(Cadena, "errorcode")
    Query.AddField "pagodeservicios", "descripCodigoRespuesta", DatosXML(Cadena, "responsemessage")
    Query.AddField "pagodeservicios", "folioCarrier", DatosXML(Cadena, "authcode")
    Query.AddField "pagodeservicios", "xmlDeRespuesta", Cadena
	Query.Exec

	Script.ImprimeFormato "TICKETSRV", (id), (Ambiente), Me, False 

End Sub                        


Function DatosXML(operacion, datoABuscar)
    Set xml = CreateObject("Msxml2.DOMDocument.3.0")
    xml.loadXML((operacion))
                                                                             
    Set xVenta = xml.documentElement
                                                          
	DatosXML = xVenta.getAttribute((datoABuscar))

End Function                    


Function ExisteArticuloTiempoAireOEspecial()
	Dim rstPV

	Set rstPV = Rst("SELECT articulo FROM partvta WHERE venta = " & Me.Venta & " AND (articulo = 'RMOVISTAR' OR " & _
					"articulo = 'RTELCEL' OR articulo = 'RIUSACELL' OR articulo = 'RUNEFON' OR articulo = 'RNEXTEL' OR " & _
					"articulo = 'RSATFEMYB' OR articulo = 'RSERVICIOS' OR articulo = 'RVIRGIN')", Ambiente.Connection)

	If Not rstPV.EOF Then
    	ExisteArticuloTiempoAireOEspecial = True
	Else
		ExisteArticuloTiempoAireOEspecial = False
	End If

	rstPV.Close
	Set rstPV = Nothing	
End Function
                                        

Sub NuevaPartida(rstArticulo, nPrecio)
    
	Set rstVenta = Rst("SELECT venta FROM ventas WHERE venta = " & Me.Venta, Ambiente.Connection)
    
	If rstVenta.EOF Then
		SalvaVenta False
	End If
      
	Query.Reset
	Query.strState = "INSERT"
	Query.AddField "partvta", "clave", ""
	Query.AddField "partvta", "prcantidad", 1
	Query.AddField "partvta", "venta", Me.Venta
	Query.AddField "partvta", "articulo", rstArticulo("Articulo")
	Query.AddField "partvta", "cantidad", 1
	Query.AddField "partvta", "precio", nPrecio / Me.tipoDeCambio
	Query.AddField "partvta", "preciobase", rstArticulo("precio1") / Me.tipoDeCambio
	Query.AddField "partvta", "descuento", 0
	Query.AddField "partvta", "impuesto", rstArticulo("valor")
	Query.AddField "partvta", "observ", Trim(rstArticulo("Descrip"))
	Query.AddField "partvta", "partida", 0
	Query.AddField "partvta", "usuario", Me.Ambiente.uId
	Query.AddField "partvta", "usufecha", Date
	Query.AddField "partvta", "usuhora", Formato(Time, "HH:mm:ss")
	Query.AddField "partvta", "almacen", 1
	Query.AddField "partvta", "lista", 1
	Query.AddField "partvta", "caja", Me.Ambiente.Estacion
	Query.AddField "partvta", "estado", "PE"
	Query.AddField "partvta", "devconf", 0
	Query.AddField "partvta", "iespecial", rstArticulo("iEspecial")
	Query.AddField "partvta", "kit", rstArticulo("kit")
	Query.AddField "partvta", "costo_u", rstArticulo("costo_u")
	Query.AddField "partvta", "costo", rstArticulo("costo")
	Query.AddField "partvta", "donativo", rstArticulo("Donativo") / Me.tipoDeCambio
	Query.AddField "partvta", "invent", 1
	Query.Exec

	nPartida = LastInsertedId(Ambiente.Connection)              

	fg2.Redraw = False
        
	For n = Me.UltimaPartida To fg2.Rows - fg2.FixedRows
		If clEmpty(fg2.TextMatrix(n, 0)) Then
			fg2.TextMatrix(n, 10) = 1
			fg2.TextMatrix(n, 0) = rstArticulo("Articulo")
			fg2.TextMatrix(n, 1) = 1
			fg2.TextMatrix(n, 2) = nPrecio / Me.tipoDeCambio
			fg2.TextMatrix(n, 3) = Formato(0, formatoDeDinero)
			fg2.TextMatrix(n, 5) = Formato(rstArticulo("Valor"), Ambiente.FDinero)
			fg2.TextMatrix(n, 6) = Trim(rstArticulo("Descrip"))
			fg2.TextMatrix(n, 14) = Formato(rstArticulo("iEspecial"), Ambiente.FDinero)
			fg2.Row = n
			fg2.Col = 6
			fg2.CellFontBold = True
			fg2.Col = 1
			fg2.TextMatrix(n, 7) = nPartida
			fg2.TextMatrix(n, 8) = 1
			fg2.TextMatrix(n, 9) = 1
			fg2.TextMatrix(n, 12) = Val2(rstArticulo("costo_u"))
			fg2.TextMatrix(n, 17) = Val2(rstArticulo("Donativo")) / Me.tipoDeCambio
          
			Me.UltimaPartida = n
			AjustaPartida (n)
           
			Exit For
           
		End If
	Next
    
	Me.CantidadDeArticulos = 0
	Me.Peso = 0
    
	CalculaImportes
	fg2.Redraw = True
    
	Me.Peso = 0
End Sub
  

Sub calculaPrecioDeCaja()
    Dim rstClaveAdd          
    Dim rstArticulo

    Set rstClaveAdd = CreaRecordSet( "SELECT articulo, cantidad FROM clavesadd WHERE clave = '" & Me.Articulo & "'", Ambiente.Connection )

    If rstClaveAdd.EOF Then 
       Exit Sub 
    End If
               
    If rstClaveAdd("cantidad") = 1 Then
       Exit Sub
    End If

    Set rstArticulo = CreaRecordSet( "SELECT precio3 FROM prods WHERE articulo = '" & Trim(rstClaveAdd("articulo")) & "'", Ambiente.Connection )
                                                    
    If Not rstArticulo.EOF Then
       Me.PrecioEspecial = rstArticulo("precio3")
    End If

End Sub                  


Sub recuperaVenta()
    Dim rstVenta, nVenta

    nVenta = Val2( Mid( Articulo, 2 ) ) 
    If nVenta = 0 Then
       Exit Sub
    End If

    Set rstVenta = CreaRecordSet( _
    "SELECT * FROM ventas WHERE venta = " & nVenta, _
    Ambiente.Connection )

    If rstVenta.EOF Then
       ' El campo 3 despliega los mensajes en el punto de venta
       txtFields(3) = "Venta no encontrada"
       Exit Sub
    End If    
                                          
    If rstVenta("estado") <> "PE" Then
       txtFields(3) = "Venta no valida"
       Exit Sub
    End If  
                          
    If rstVenta("ocupado") <> 0 Then
       MyMessage "La venta ya esta ocupada por otro usuario: " & rstVenta("usuario")
       Exit Sub
    End If
    
    If Parent.Venta > 0 Then
       Ambiente.Connection.Execute "UPDATE ventas SET ocupado = 0, usuario = '" & Ambiente.uId & "' WHERE venta = " & Parent.Venta
    End If
      
    Parent.txtFields(0).Enabled = False
    Set Parent.rstVenta = rstVenta
    Set Parent.rstPartidas = CreaRecordSet("SELECT * FROM partvta WHERE venta = " & nVenta, Ambiente.Connection)
	Parent.Venta = nVenta
    Parent.ReLoad = True
    Parent.ReiniciaVenta
    Eventos
    Parent.txtFields(0).Enabled = True
    Parent.CalculaImportes
                                            
End Sub          



Sub Version2005       
    Dim rstEmpleado   
    Dim cEmpleado
    Dim Query
    Dim nPos
    Dim cCodigo
    Dim cPeso
    Dim rstAsistencia                 
    Dim Contenido                      
                      
    ' Articulo es una variable que entrega el punto de venta
    ' Esto proceso todos los comandos del punto de venta
    Call ValidaComando
             
    'validaNumeroDePartidas              

    If UCase(Trim(Articulo)) = "CENEFASXMARCA" Then
       Script.RunForm "CENEFASXMARCA", Me, Ambiente,, True
       CancelaProceso = True
    End If

    If UCase(Trim(Articulo)) = "CENEFAS" Then
       Script.RunForm "CENEFAS", Me, Ambiente,, True
       CancelaProceso = True
    End If

    If Ucase(Trim(Articulo)) = "BORDADO" Then
       Script.RunForm "BORDADO", Me, Ambiente,, True   
    End If
 
    If Ucase(Trim(Articulo)) = "LONAS" Then
       Script.RunForm "LONAS", Me, Ambiente,, True   
    End If

    
    If Ucase(Trim(Articulo)) = "ESTAMPADO01" Then
       Script.RunForm "ESTAMPADO", Me, Ambiente,, True   
    End If

    If Ucase(Trim(Articulo)) = "ESTAMPADO02" Then
       Script.RunForm "ESTAMPADO02", Me, Ambiente,, True   
    End If

    If Ucase(Trim(Articulo)) = "ESTAMPADO03" Then
       Script.RunForm "ESTAMPADO03", Me, Ambiente,, True   
    End If


    'SumaCantidades
    ' Articulo es una variable publica que contiene el dato que leyo el lector
    ' o que tecleo el usuario hasta el momento de presionar un enter
    nPos = clAt( "*", Articulo  )

    ' cantidadDeArticulos es una variable publica que indica la cantidad
    ' de productos que va a aceptar el punto de venta
    If nPos > 0 Then
       cantidadDeArticulos = Val2( Mid( Articulo, 1, nPos - 1 ) )
       Articulo = Mid( Articulo, nPos + 1 )
    End If
                              
 
    ' Para articulos con peso
    If clAt( "ARTICULOPARAPESO", Articulo ) = 1 Then
       cCodigo = Mid( Articulo, 3, 5 )
       cPeso   = Mid( Articulo, 8 ) 
       Articulo = cCodigo
       CantidadDeArticulos = (Val2(cPeso) / 10000)
    End if                              
                                
    nPos = clAt( "CLI", Articulo )

    ' La varable cancelaProceso termina el flujo de programa interno de MyBusiness
    ' Parecido a Exit Sub 
    If nPos = 1 Then
       txtFields(0) = Mid( Articulo, 4 )
       txtFields(4) = "" 
       CancelaProceso = True          
       ChecaSaldo txtFields(0)
       ColocaFoto txtFields(0)       
    End If    

    If clAt( "EMP", Articulo ) = 1 Then
       cEmpleado = UCase(Mid( Articulo, 4 ))
       Set rstEmpleado = CreaRecordSet( "SELECT * FROM empleados WHERE empleado = '" & cEmpleado & "'", Ambiente.Connection )

       If rstEmpleado.EOF Then
          txtFields(3) = "Empleado no existe"
          cancelaProceso = True
          txtFields(4) = ""
          PlaySound Ambiente.Path & "\sounds\error.wav"
          Exit Sub
       End If                           

       'Set rstEmpFecha = CreaRecordSet( "SELECT * FROM asistencia WHERE empleado = '" & cEmpleado & "' AND fecha = " & fechaSQL( Date, Ambiente.Connection ), Ambiente.Connection )

       'If Not rstEmpFecha.EOF Then             
       '   txtFields(3) = "Usted ya fue registrado"
       '   cancelaProceso = True
       '   txtFields(4) = ""
       '   PlaySound Ambiente.Path & "\sounds\error.wav"
       '   Exit Sub
       'End If

       'PlaySound Ambiente.Path & "\sounds\ready.wav"

       Set Query = NewQuery()
       Set Query.Connection = Ambiente.Connection
  
       Query.strState = "INSERT"

       Query.AddField "asistencia", "id", TraeSiguiente( "asistencia", Ambiente.Connection )
       Query.AddField "asistencia", "fechahora", Formato( Date, "dd-MM-yyyy" ) & ":" & Formato( Time, "hh:mm:ss" )
       Query.AddField "asistencia", "retardo", Retardo( rstEmpleado("horaentrada") )       
       Query.AddField "asistencia", "empleado", cEmpleado
       Query.AddField "asistencia", "fecha", Date
       Query.CreateQuery
       Query.Execute       

       txtFields(3) = rstEmpleado("nombre") & " " & Formato( Date, "dd-MM-yyyy" ) & ":" & Formato( Time, "hh:mm:ss" )
       CancelaProceso = True
       txtFields(4) = ""

       If Not clEmpty( rstEmpleado("imagen") ) Then
          CreaHTML "", "<img src='" & Trim(rstEmpleado("imagen")) & "'>"
       End If

    End If

    If Ucase(Trim(Articulo)) = "ARTICULOS PROMOCIONALES" Then
       Script.RunProcess "GALAXIA", Me.Parent, Ambiente
       CancelaProceso = True
       txtFields(4) = ""
    End If            

    If Mid(UCase(Trim(Articulo)),1,6) = "MONEDA" Then
       cambiaMonedaDeLaVenta Trim(Mid( Articulo, 7 ))
       cancelaProceso = True
    End If                                         

    'Call ValidaExistencia()

End Sub                    


Sub ValidaExistencia()
    Dim rstProd

    Set rstProd = CreaRecordSet( "SELECT alm" & Ambiente.Almacen & " FROM prods WHERE articulo = '" & articulo & "'", _
    Ambiente.Connection )

    If Not rstProd.EOF Then
       If rstProd(0) <= 0 Then
          Me.OperacionBloqueada = True
          txtFields(3) = "Existencia insuficiente, operación bloqueada" 
          CancelaProceso = True                   
          PlaySound Ambiente.Path & "\sounds\Existencia.wav"
       End If 
    End If

End Sub
                                               

Sub cambiaMonedaDeLaVenta( cMoneda )
    Dim rstVenta, rstMoneda, Query, rstMonedaOriginal

    Set rstVenta = CreaRecordSet( _
    "SELECT * FROM ventas WHERE venta = " & Me.Venta, Ambiente.Connection )

    Set rstMoneda = CreaRecordSet( _
    "SELECT * FROM monedas WHERE moneda = '" & cMoneda & "'", Ambiente.Connection )

    If Not rstMoneda.EOF Then
       Me.Moneda = cMoneda
       Me.tipoDeCambio = rstMoneda("tc")        
    Else
       Exit Sub
    End If

    If rstVenta.EOF Then
       Exit Sub
    End If                        

    If Ucase(Trim(rstVenta("moneda"))) = UCase(Trim(cMoneda)) Then
       Exit Sub
    End If     

    Ambiente.Connection.Execute _
    "UPDATE partvta SET precio = precio * " & FormatoDecimal( rstVenta("tipo_cam") ) & ", " & _
    "preciobase = preciobase * " & FormatoDecimal( rstVenta("tipo_cam") ) & ", " & _
    "donativo = donativo * " & FormatoDecimal( rstVenta("tipo_cam") ) & " " & _
    "WHERE venta = " & Me.Venta


    Ambiente.Connection.Execute _
    "UPDATE partvta SET precio = precio / " & FormatoDecimal( Me.tipoDeCambio ) & ", " & _
    "preciobase = preciobase / " & FormatoDecimal( Me.tipoDeCambio ) & ", " & _ 
    "donativo = donativo / " & FormatoDecimal( Me.tipoDeCambio ) & " " & _
    "WHERE venta = " & Me.Venta

    Set Query = NewQuery()
    Set Query.Connection = Ambiente.Connection          

    Query.Reset
    Query.strState = "UPDATE"
    Query.Condition = "venta = " & Me.Venta
    Query.AddField "ventas", "moneda", Me.Moneda
    Query.AddField "ventas", "tipo_cam", Me.tipoDeCambio
    Query.CreateQuery
    Query.Execute

    Parent.ReLoad = True
    ReiniciaVenta 
    CalculaImportes       

End Sub


Sub ChecaSaldo( Cliente ) 
    Dim rstSaldo, rstCobranza, Html
    
    Set rstSaldo = CreaRecordSet( "SELECT * FROM clients WHERE cliente = '" & Cliente & "'", _
        Ambiente.Connection )
    
    If rstSaldo.EOF Then
       Exit Sub
    End If

    CreaHTML "",""

    If rstSaldo("Saldo") <= 0 Then
       Exit Sub
    End If    

    PlaySound "c:\saldo.wav"

    Set rstCobranza = CreaRecordSet( _
    "SELECT * FROM cobranza WHERE cliente = '" & cliente & "' AND saldo > 0", _
    Ambiente.Connection )
                              
    html = ""   

    While Not rstCobranza.EOF
          html = html & "<p>"
          html = html & "Documento " & rstCobranza("tipo_doc") & _
          rstCobranza("no_referen") & " " & rstCobranza("Saldo")
          html = html & "</p>"
          rstCobranza.MoveNext
    Wend          

    CreaHTML "",(Html)

End Sub


Function Retardo( strHora )
    Dim intMinutos 
    Dim intMinutosActual 
    Dim nPos
    Dim strHoraActual 
    Dim tolerancia

    tolerancia = 10
  
    nPos = clAt( ":", strHora )    
    intMinutos = Val2(  Mid( strHora, 1, nPos - 1)  ) * 60
    intMinutos = intMinutos + Val2( Mid( strHora, nPos + 1 ) )

    strHoraActual = Formato( Time(), "hh:mm" )
    nPos = clAt( ":", strHoraActual )    
    intMinutosActual = Val2(  Mid( strHoraActual, 1, nPos - 1)  ) * 60
    intMinutosActual = intMinutosActual + Val2( Mid( strHoraActual, nPos + 1 ) )

    If intMinutosActual > (intMinutos + tolerancia) Then       
       Retardo = intMinutosActual - intMinutos
    Else
       Retardo = 0
    End If

End Function



Sub ColocaFoto( cCliente )
    Dim cHtml               
 
    cHtml = "<img src=c:\fotos\" & Trim( cCliente ) & ".jpg width=100%>"

    CreaHTML "", (cHtml)

End Sub



Sub SumaCantidades()

    ' Buscamos si el artículo ya esta en el GRID 
    For n = 1 to fg2.Rows - 1 
        If clEmpty( fg2.TextMatrix( n, 0 ) ) Then 
           Exit For 
        End If 

		Set rstArtAux = Rst("SELECT tiempoaire FROM prods WHERE articulo = '" & Trim(Articulo) & "'", Ambiente.Connection)     

        If Trim(fg2.TextMatrix( n, 0 )) = Trim(Articulo) AND Val2(rstArtAux("tiempoaire")) <> 0 Then 
           Ambiente.Connection.Execute "UPDATE partvta SET cantidad = cantidad + 1 WHERE id_salida = " & fg2.TextMatrix( n, 7 ) 
           fg2.TextMatrix( n, 1 ) = Val2( fg2.TextMatrix( n, 1 ) ) + 1 
           txtFields(4) = "" 
           CancelaProceso = True
           Exit Sub
        End If 
    Next

End Sub




Sub validaNumeroDePartidas
    Dim rstPartidas
    
    Set rstPartidas = CreaRecordSet( _
        "SELECT COUNT( * ) FROM partvta WHERE venta = " & Me.Venta, _
        Ambiente.Connection )

    If Val2( rstPartidas(0) ) >= 4 Then
       MyMessage "No es posible capturar mas de 4 partidas"
       CancelaProceso = True
    End If    

End Sub

Sub ValidaComando()
                              
    If clAt( "/", Articulo ) > 0 And clAt( "//", Articulo ) = 0 Then
       Call cantidaporPrecio()
       Exit Sub
    End If


    If Trim(UCase(Articulo)) = "COMPRAS" Then
       Set Compras = CreateObject( "My2016BCompras.Compras" )
       Set Compras.Ambiente = Ambiente
       Compras.NuevaCompra True
    End If
                                    
    If Len( Trim( Articulo ) ) <> 4 Then
       Exit Sub
    End If

    If clAt( "Z", UCase(Articulo) ) <> 1 Then
       Exit Sub
    End If

    Select Case UCase(Articulo)
           Case "Z001"
                Script.RunFormSG "ALTACLIENTE", Me, Ambiente,, True    

           Case "Z002"
                Me.OperacionBloqueada = False
                txtFields(3) = "Operación Reactivada"

           Case "Z003"
                Me.FinalizaOperacion

           Case "Z004"
                Script.RunFormSG "ALTARAPIDA", Me, Ambiente,, True

           Case "Z005"              

                If Not ExisteArticuloTiempoAireOEspecial Then
	                If Question(Mensaje(642, Ambiente)) Then
	                   Me.BorraVenta = True
	                   Me.FinalizaOperacion
	                End If
       			End If

          Case "Z006"

                If Not ExisteArticuloTiempoAireOEspecial Then
					Me.RecuperaVentaDeCliente
       			End If               

          Case "Z007"

               If txtFields(0).Enabled Then
                  txtFields(0).SetFocus
               End If

          Case "Z008"

               AplicaDescuento 

          Case "Z009"

               Script.RunProcess "CORTEX", Me, Ambiente
                
          Case "Z010"
                            
               Script.RunProcess "CORTEZ", Me, Ambiente

          Case "Z011"

               Script.RunFormSG "PAGOEFECTIVO", Me, Ambiente,, True

          Case "Z012"

               Script.RunFormSG "COBROENEFECTIVO", Me, Ambiente,, True
       
          Case "Z013"
                                                   
               Set rstUsuventas = CreaRecordSet( "SELECT * FROM usuventas WHERE usuario = '" & Ambiente.Uid & "'", Ambiente.Connection )
                              
               If rstUsuventas.EOF Then
                  Script.RunForm "DEVOLUCIONES", Me, Ambiente, True
                  Exit Sub
               End If

               If Val2( rstUsuventas("devpunto") ) <> 0 Then
                  Script.RunForm "DEVOLUCIONES", Me, Ambiente, True
                  'Exit Sub              
               Else
                  MyMessage "No tiene derecho a hacer devoluciones, solicite el permiso con su supervisor" 
                  'Exit Sub
               End If  

          Case "Z014"

               Script.RunFormSG "TICKETAFACTURA", Me, Ambiente,, True

          'Case "Z016"
          '                                                      
          '     Me.ActivaCobranza

          Case "Z015"

               Script.RunProcess "PUNTOV063", Me, Me.Ambiente

          Case "Z016"
                                                                
               Script.RunFormSG "RETICKET", Me, Ambiente,, True

          'Case "Z019"
          '
          '     MyMessage "CONTROL + F12 Editar datos cliente" & vbCrLf & "CONTROL + F7 Editar guión" & vbCrLf & "SHIFT + F3 Colocar cursor en campo repartidor" & vbCrLf & "SHIFT + F4 Alta de repartidor"

          Case "Z017"
                                                                
               Script.RunFormSG "FCAJA", Me, Ambiente,, True

          Case "Z018"

               Script.RunForm "CAMBIOUSUARIO", Me, Ambiente,, True 

          'Case "Z022" 
          '             
          '     Script.RunProcess "VENTASCOLECTOR", Me, Me.Ambiente               

          Case "Z019" 
                    
				CalculadoraPuntoDeVenta
 
          Case "Z020"

               Script.RunProcess "ABRECAJON", Me, Me.Ambiente                    

          Case "Z021"
                                                                   
               If Not ExisteArticuloTiempoAireOEspecial Then
               		Script.RunProcess "ELIMINAVENTA", Me, Me.Ambiente                    
               End If

          Case "Z022"

               Script.RunFormSG "VENTASOBSERV", Me, Ambiente,, True

          Case "Z023"

               Script.RunFormSG "VENTASDESC", Me, Ambiente,, True

          Case "Z024"

               Script.RunForm "CAMBIOVALEEFECTIVO", Me, Me.Ambiente,, True

          Case "Z025"

             Script.RunformSG "FACTURADECIERRE", Me, Ambiente,, False

          Case "Z026"

             Script.RunHuellaForm "REGISTROACCESO", Me, Ambiente,, True
          
          Case "Z028"
             Script.RunNetScript "ROLVISITAPROVEEDORES", "", Ambiente, "?Usuario=" & Ambiente.Uid & "?Estacion=" & Ambiente.Estacion & "?TodasFechas=No?TodosProveedores=Si?DiaInicial=" & Date & "?DiaFinal=" & Date
             'Script.Runform "LISTADEPEDIDOS", Me, Ambiente,, True
          Case "Z029"
             Script.Runform "CAMBIODETURNO", Me, Ambiente,, True            

          Case "Z030"

             Script.RunNetScript "LISTADEVENTASPORFACTURAR", "", Ambiente, ""    

		 Case "Z100"

                 ShellCommand "C:\MyBusinessPOS20\pservicios\CredicapitalWinForm.exe",1  

    End Select                    

    CancelaProceso = True
    txtFields(4) = ""

End Sub



Sub llenaPuntoDeVenta
    Dim cLinea, dondeEstaLaComa, articulo
    
    ' CloseFile cierra un manejador de archivo
    CloseFile 1                               
    ' OpenFile Abre un archivo de texto para lectura
    OpenFile "c:\articulos.txt", 1

    ' FileEOF regresa true en caso de que el archivo
    ' de texto llegue a su final
    While Not FileEOF( 1 )      
          ' ReadLine lee una cadena de texto
          ' hasta que encuentra un retorno de carro
          cLinea = ReadLine( 1 )
          dondeEstaLaComa = clAt( ",",  (cLinea) )
          articulo = Mid(cLinea, 1, dondeEstaLaComa -1 )
          
          ' cantidadDeArticulos es una variable que
          ' entrega el punto de venta e indica la 
          ' cantidad de productos que se va a ingresar
          ' en el punto de venta
          cantidadDeArticulos = _
          Mid( cLinea, dondeEstaLaComa + 1 ) 
          
          ' llenaPartidad Es un metodo que ingresa
          ' producto en el punto de venta
          llenaPartida (articulo)
    Wend
                               
    CloseFile 1

End Sub



Sub calculaPeso()
    Dim rstArticulo, cArticulo, nPesos
                         
    nPos = clAt( "--", Articulo )
    cArticulo = Mid( Articulo, nPos + 2 )
    nPesos = Val2(Mid( Articulo, 1, nPos - 1 ))

    Set rstArticulo = CreaRecordSet( _
    "SELECT * FROM prods WHERE articulo = '" & cArticulo & "'", _
    Ambiente.Connection )

    If Not rstArticulo.EOF Then
       Articulo = cArticulo
       cantidadDeArticulos =  nPesos /  rstArticulo("precio1")
    End If

End Sub
                                                                     

Function Permisos( permisoSolicitado )

    Set rstUsuario = CreaRecordSet( _ 
    "SELECT * FROM usuarios WHERE usuario = '" & Ambiente.uid & "'", Ambiente.connection )
    
    If rstUsuario("supervisor") Then       
	   cmdZ001.Caption = "Z001 Cambia o da de alta un cliente"
	   cmdZ002.Caption = "Z002 Quita bloqueo por error código de producto"
	   cmdZ003.Caption = "Z003 Confirma la venta o muestra ventana de cobro"
	   cmdZ004.Caption = "Z004 Da de alta o modifica un artículo"
	   cmdZ005.Caption = "Z005 Deja como pendiente la venta actual"
	   cmdZ006.Caption = "Z006 Muestra lista de ventas pendientes"
	   cmdZ007.Caption = "Z007 Posicionar el cursor en el campo de cliente"
	   cmdZ008.Caption = "Z008 Pantalla de descuentos"
	   cmdZ009.Caption = "Z009 Corte parcial X"
	   cmdZ010.Caption = "Z010 Corte total Z"
	   cmdZ011.Caption = "Z011 Permite capturar un ingreso de dinero a caja"
	   cmdZ012.Caption = "Z012 Permite capturar una salida de dinero a caja"
	   cmdZ013.Caption = "Z013 <ticket> <estación> Realiza una devolución de mercancia en caja"
	   cmdZ014.Caption = "Z014 <ticket> <estacion> Convierte un ticket en factura"
	   cmdZ015.Caption = "Z015 <devolución> Realiza el pago de la devolucion"
	   cmdZ016.Caption = "Z016 Cobranza a clientes"
	   cmdZ017.Caption = "Z017 Muestra información del próducto"
	   cmdZ018.Caption = "Z018 Re imprimir ticket"
	   cmdZ019.Caption = "Z019 Functiones ventas por teléfono (Informativo)"
	   cmdZ020.Caption = "Z020 Arqueo de efectivo"
	   cmdZ021.Caption = "Z021 Cambio de usuario"
	   cmdZ022.Caption = "Z022 Recupera ventas del colector"
	   cmdZ023.Caption = "Z023 Muestra la calculadora"
	   cmdZ024.Caption = "Z024 Abre el cajón de dinero"
	   cmdZ025.Caption = "Z025 Elimina la ultima venta"
	   cmdZ026.Caption = "Z026 Observaciones a la venta"     
       cmdZ027.Caption = "Z027 Descuento por importe"
    Else  
	   cmdZ001.Caption = "Z001 Cambia o da de alta un cliente"
	   cmdZ002.Caption = "Z002 Quita bloqueo por error código de producto"
	   cmdZ003.Caption = "Z003 Confirma la venta o muestra ventana de cobro"
	   cmdZ004.Caption = "Z005 Deja como pendiente la venta actual"
	   cmdZ005.Caption = "Z006 Muestra lista de ventas pendientes"
	   cmdZ006.Caption = "Z007 Posicionar el cursor en el campo de cliente"
	   cmdZ007.Caption = "Z011 Permite capturar un ingreso de dinero a caja"
	   cmdZ008.Caption = "Z012 Permite capturar una salida de dinero a caja"
	   cmdZ009.Caption = "Z015 <devolución> Realiza el pago de la devolucion"
	   cmdZ010.Caption = "Z020 Arqueo de efectivo"
	   cmdZ011.Caption = "Z021 Cambio de usuario"
	   cmdZ012.Caption = "Z024 Abre el cajón de dinero"
	   cmdZ013.Caption = "Z025 Elimina la ultima venta"
	   cmdZ014.Caption = "Z026 Observaciones a la venta"
       cmdZ015.Caption = "Z027 Descuento por importe"
    End If


End Function



Sub cantidaporPrecio()
    Dim nPos, cCantidad, cArticulo, rstProd
    Dim nPrecio
            
    nPos = clAt( "/", Articulo )
    cCantidad = Mid( Articulo,1, nPos - 1 )
    cArticulo = Trim( Mid( Articulo, nPos + 1 )  )
                                   
    Set rstProd = CreaRecordSet( "SELECT * FROM prods WHERE articulo = '" & cArticulo & "'", Ambiente.Connection )

    If rstProd.EOF Then
       Exit Sub
    End If   

    nCantidad = Round( Val2( cCantidad ) / rstProd("precio1"), 4 )

    If Not rstProd.EOF Then            
       cantidadDeArticulos = nCantidad
       llenaPartida rstProd("Articulo") 
    End If

    CancelaProceso = True
End Sub                      



Sub incrementaProducto()
    Dim rstArticulo, Query 

    Set rstArticulo = CreaRecordSet( _
    "SELECT articulo FROM prods WHERE articulo = '" & Me.Articulo & "' AND granel = 0 AND speso = 0", _
    Ambiente.Connection ) 

    Set Query = NewQuery()
    Set Query.Connection = Ambiente.Connection

    If rstArticulo.EOF Then
       Exit Sub
    End If     

    For n = 0 To fg2.Rows - 1

        If clEmpty( fg2.TextMatrix( n, 0 ) ) Then
           Exit For
        End If

        If Ucase(  Trim( fg2.TextMatrix(n,0) )  ) = _
           UCase( Trim(Me.Articulo) ) Then

           Query.SQL = _
        "UPDATE partvta SET cantidad = cantidad + 1 WHERE id_salida = " & _
           fg2.TextMatrix(n,7)
           Query.Exec

           fg2.TextMatrix(n,1) = Val2(fg2.TextMatrix(n,1)) + 1 
           CancelaProceso = True
           Me.CalculaImportes
           Exit For
        End If  

    Next


End Sub