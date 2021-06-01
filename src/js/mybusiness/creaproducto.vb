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