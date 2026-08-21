Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.Data
Imports DarkSoft
Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports System.Text.RegularExpressions

Public Class Generador
  Private PageState As CustomPageState
  'Public con As New SQL.Conexion("CRM", "201.158.105.66,2332", "C0muN1K10", "Va4MI1lA3t4y")
  'Public con As New SQL.Conexion("CRM", "localhost", "C0muN1K10", "Va4MI1lA3t4y")
  'Public con As New SQL.Conexion("CRM", "localhost", "4dm1n", "C0n3ktR8m6*")
  Public con As New SQL.Conexion("CRM", "DESKTOP-MM7GFSK\SQLEXPRESS", "njl", "nico123*")

  Private Function registrarEstado(ByVal id_cliente As Integer, ByVal id_contrato As Integer, ByVal estatus As Integer) As Integer
    Dim msj As String = ""
    Dim periodos As Object = getPeriodos(id_contrato, estatus)
    Dim paquete As Object = getPaquete(id_contrato)
    Dim excedentes_mat As Double = 0
    Dim excedentes_tel As Double = 0
    Dim otros_cobros As Double = getOtrosCobros(id_contrato)
    Dim dtCargos As DataTable = getDataCargos(id_contrato)
    Dim descuentos As Double = getDescuentos(id_contrato)
    Dim periodoA As Date = periodos.PeriodoA
    Dim periodoB As Date = periodos.PeriodoB
    Dim saldo_pendiente As Double = getSaldoPendiente(id_contrato)
    Dim id_paquete As Integer = paquete.idPaquete
    Dim mensualidad As Double = paquete.Mensualidad
    Dim id_estado_cuenta As Integer = 0
    Dim grantotal As Double = mensualidad + excedentes_mat + excedentes_tel + otros_cobros
    grantotal -= descuentos
    grantotal += saldo_pendiente

    Dim sql As String = $"insert into ESTADOS_CUENTA values(" & id_cliente & "," & id_contrato & "," & id_paquete & "," & mensualidad & "," & FormatNumber(excedentes_mat, 2).Replace(",", "") & "," & FormatNumber(excedentes_tel, 2).Replace(",", "") & "," & FormatNumber(otros_cobros, 2).Replace(",", "") & "," & FormatNumber(descuentos, 2).Replace(",", "") & "," & FormatNumber(grantotal, 2).Replace(",", "") & ",'" & periodoA & "','" & periodoB & "',1,getdate()," & FormatNumber(saldo_pendiente, 2).Replace(",", "") & ");SELECT @@IDENTITY as Id;"
    'Dim sql As String = $"insert into ESTADOS_CUENTA values(" & id_cliente & "," & id_contrato & "," & id_paquete & "," & mensualidad & "," & FormatNumber(excedentes_mat, 2).Replace(",", "") & "," & FormatNumber(excedentes_tel, 2).Replace(",", "") & "," & FormatNumber(otros_cobros, 2).Replace(",", "") & "," & FormatNumber(descuentos, 2).Replace(",", "") & "," & FormatNumber(auxTotalBill, 2).Replace(",", "") & ",'" & periodoA & "','" & periodoB & "',1,getdate()," & FormatNumber(saldo_pendiente, 2).Replace(",", "") & ");SELECT @@IDENTITY as Id;"
    'MsgBox(sql)
    Dim dt As DataTable = con.ConsultarDT(sql)

    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      id_estado_cuenta = Val(dt(0)("Id").ToString)
      Dim sqlupd As String = "update CONTRATOS set fecha_edo_cta=DATEADD(MONTH,1,fecha_edo_cta) where id_contrato=" & id_contrato
      con.ModRegEli(sqlupd)
    End If

    If id_estado_cuenta > 0 And (dtCargos IsNot Nothing AndAlso dtCargos.Rows.Count > 0) Then
      For i = 0 To dtCargos.Rows.Count - 1
        Dim sqlins As String = "insert into DET_EDOS_CARGOS values(" & id_estado_cuenta & "," & Val(dtCargos(i)(0).ToString) & ")"
        con.ModRegEli(sqlins)
        Dim sqlUpdate As String = "update OTROS_CARGOS set estatus=0 where id_otro_cargo=" & Val(dtCargos(i)(0).ToString)
        con.ModRegEli(sqlUpdate)
      Next
    End If

    Return id_estado_cuenta
  End Function

  Private Function registerBill(ByVal id_cliente As Integer, ByVal id_contrato As Integer, ByVal estatus As Integer) As Integer
    Dim msj As String = ""
    Dim periodos As Object = getPeriodos(id_contrato, estatus)
    Dim paquete As Object = getPaquete(id_contrato)
    Dim excedentes_mat As Double = 0
    Dim excedentes_tel As Double = 0
    Dim otros_cobros As Double = getOtrosCobros(id_contrato)
    Dim dtCargos As DataTable = getDataCargos(id_contrato)
    Dim descuentos As Double = getDescuentos(id_contrato)
    Dim periodoA As Date = periodos.PeriodoA
    Dim periodoB As Date = periodos.PeriodoB
    Dim saldo_pendiente As Double = getSaldoPendiente(id_contrato)
    Dim id_paquete As Integer = paquete.idPaquete
    Dim mensualidad As Double = paquete.Mensualidad
    Dim id_estado_cuenta As Integer = 0
    Dim grantotal As Double = mensualidad + excedentes_mat + excedentes_tel + otros_cobros
    grantotal -= descuentos
    'grantotal += saldo_pendiente

    Dim balance As Double = 0
    Dim sqlBalance As String = "select top 1 coalesce(b.balance,0) as balance from CONTRATOS c WITH (NOLOCK) left join CONTRACTS_BALANCES b WITH (NOLOCK) on c.id_contrato=b.id_contrato " &
        "where c.id_contrato=" & id_contrato & " order by b.id desc;"
    Dim dtBalance As DataTable = con.ConsultarDT(sqlBalance)

    If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
      balance = Val(dtBalance(0)("balance").ToString)
    End If

    Dim auxTotal As Double = grantotal + balance
    Dim auxTotalBill = 0

    If auxTotal >= 0 Then
      auxTotalBill = auxTotal
    End If

    Dim statusBill As Integer = 1

    If auxTotalBill <= 0 Then
      statusBill = 0
    End If

    'Dim sql As String = $"insert into ESTADOS_CUENTA values(" & id_cliente & "," & id_contrato & "," & id_paquete & "," & mensualidad & "," & FormatNumber(excedentes_mat, 2).Replace(",", "") & "," & FormatNumber(excedentes_tel, 2).Replace(",", "") & "," & FormatNumber(otros_cobros, 2).Replace(",", "") & "," & FormatNumber(descuentos, 2).Replace(",", "") & "," & FormatNumber(grantotal, 2).Replace(",", "") & ",'" & periodoA & "','" & periodoB & "',1,getdate()," & FormatNumber(saldo_pendiente, 2).Replace(",", "") & ");SELECT @@IDENTITY as Id;"
    Dim sql As String = $"insert into ESTADOS_CUENTA values(" & id_cliente & "," & id_contrato & "," & id_paquete & "," & mensualidad & "," & FormatNumber(excedentes_mat, 2).Replace(",", "") & "," & FormatNumber(excedentes_tel, 2).Replace(",", "") & "," & FormatNumber(otros_cobros, 2).Replace(",", "") & "," & FormatNumber(descuentos, 2).Replace(",", "") & "," & FormatNumber(auxTotalBill, 2).Replace(",", "") & ",convert(date,'" & periodoA & "',103),convert(date,'" & periodoB & "',103)," & statusBill & ",getdate()," & FormatNumber(saldo_pendiente, 2).Replace(",", "") & ");SELECT @@IDENTITY as Id;"
    'MsgBox(sql)
    Console.WriteLine(sql)
    Dim dt As DataTable = con.ConsultarDT(sql)

    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      'MsgBox("Entra")
      id_estado_cuenta = Val(dt(0)("Id").ToString)
      Dim sqlupd As String = "update CONTRATOS set fecha_edo_cta=DATEADD(MONTH,1,fecha_edo_cta) where id_contrato=" & id_contrato
      con.ModRegEli(sqlupd)
    End If

    If id_estado_cuenta > 0 Then
      'Get balance.
      Dim sqlUpdateBalance = "INSERT INTO BALANCES values(" & id_contrato & "," & balance & ",0," & grantotal & "," & auxTotal & ",getdate(),5)"
      con.ModRegEli(sqlUpdateBalance)

      sqlUpdateBalance = "UPDATE CONTRACTS_BALANCES set balance=" & auxTotal & ",last_update=getdate() where id_contrato=" & id_contrato & ";"
      con.ModRegEli(sqlUpdateBalance)

      If auxTotal <= 0 Then
        'Search payment
        Dim sqlSearchPayment As String = "select top 1 * from PAGOS where id_contrato=" & id_contrato & " and DATEDIFF(Day, periodoA,convert(date,'" & periodoA & "',103)) = 0 and DATEDIFF(day,periodoB,convert(date,'" & periodoB & "',103)) = 0 " &
        "order by id_pago desc;"
        'Console.WriteLine("Consulta para buscar el pago")
        'Console.WriteLine(sqlSearchPayment)
        Dim dtPayment As DataTable = con.ConsultarDT(sqlSearchPayment)
        If dtPayment IsNot Nothing AndAlso dtPayment.Rows.Count > 0 Then
          Dim idPayment As Integer = Val(dtPayment(0)("id_pago").ToString())
          Dim sqlUpdatePayment = "UPDATE PAGOS set id_estado_cuenta=" & id_estado_cuenta & "where id_pago=" & idPayment
          con.ModRegEli(sqlUpdatePayment)
        Else
          Dim sqlCreatePayment As String = "insert into PAGOS values(" & id_contrato & "," & id_estado_cuenta & ",0,5,getdate(),0,convert(date,'" & periodoA & "',103),convert(date,'" & periodoB & "',103),0,0,1,1,0,0);"
          con.ModRegEli(sqlCreatePayment)
          Console.WriteLine(sqlCreatePayment)
        End If
      End If
    End If

    If id_estado_cuenta > 0 And (dtCargos IsNot Nothing AndAlso dtCargos.Rows.Count > 0) Then
      For i = 0 To dtCargos.Rows.Count - 1
        Dim sqlins As String = "insert into DET_EDOS_CARGOS values(" & id_estado_cuenta & "," & Val(dtCargos(i)(0).ToString) & ")"
        con.ModRegEli(sqlins)
        Dim sqlUpdate As String = "update OTROS_CARGOS set estatus=0 where id_otro_cargo=" & Val(dtCargos(i)(0).ToString)
        con.ModRegEli(sqlUpdate)
      Next
    End If

    Return id_estado_cuenta
  End Function

  Private Function getPaquete(ByVal id_contrato As Integer) As Object
    Dim id_paquete As Integer = 0
    Dim mensualidad_paquete As Double = 0
    Dim sql As String = "select c.id_paquete,p.mensualidad from CONTRATOS c INNER JOIN PAQUETES p" &
" ON p.id_paquete=c.id_paquete where id_contrato=" & id_contrato
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      id_paquete = Val(dt(0)("id_paquete").ToString)
      mensualidad_paquete = Val(dt(0)("mensualidad").ToString)
    End If
    Return New With {.idPaquete = id_paquete, .Mensualidad = mensualidad_paquete}
  End Function
  Private Function getSaldoPendiente(ByVal id_Contrato As Integer) As Double
    Dim saldo_pendiente As Double = 0
    Dim sql As String = "select coalesce(sum(grantotal-saldo_pendiente),0) As total from ESTADOS_CUENTA where id_contrato=" & id_Contrato & " and estatus=1"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      saldo_pendiente = Val(dt(0)(0).ToString)
    End If
    Return saldo_pendiente
  End Function
  Private Function getOtrosCobros(ByVal id_Contrato As Integer) As Double
    Dim importe As Double = 0
    Dim sql As String = "SELECT coalesce(sum(importe),0) AS importe FROM dbo.OTROS_CARGOS oc WHERE oc.id_contrato=" & id_Contrato & " AND oc.estatus=1 AND importe>0 and id_otro_cargo not in(select id_otro_cargo from DET_EDOS_CARGOS)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      importe = Val(dt(0)(0).ToString)
    End If
    Return importe
  End Function

  Private Function getDataCargos(ByVal id_contrato As Integer) As DataTable
    Dim sql As String = "SELECT id_otro_cargo FROM dbo.OTROS_CARGOS oc WHERE oc.id_contrato=" & id_contrato & " AND oc.estatus=1 AND importe<>0 and id_otro_cargo not in(select id_otro_cargo from DET_EDOS_CARGOS)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    Return dt
  End Function

  Private Function getDataBillCharges(ByVal idBill As Integer) As DataTable
    Dim sql As String = "select oc.*,t.nombre from OTROS_CARGOS oc inner join TIPOS_CARGOS t on oc.id_tipo_cargo=t.id_tipo_cargo " &
        "inner join DET_EDOS_CARGOS dt on dt.id_otro_cargo = oc.id_otro_cargo " &
        "where dt.id_estado_cuenta=" & idBill & " and t.id_tipo_cargo != 6;"
    Dim dt As DataTable = con.ConsultarDT(sql)
    Return dt
  End Function

  Private Function getDataBillDiscounts(ByVal idBill As Integer) As DataTable
    Dim sql As String = "select oc.*,t.nombre from OTROS_CARGOS oc inner join TIPOS_CARGOS t on oc.id_tipo_cargo=t.id_tipo_cargo " &
        "inner join DET_EDOS_CARGOS dt on dt.id_otro_cargo = oc.id_otro_cargo " &
        "where dt.id_estado_cuenta=" & idBill & " and t.id_tipo_cargo = 6;"
    Dim dt As DataTable = con.ConsultarDT(sql)
    Return dt
  End Function

  Private Function getStpAccount(ByVal idContract As Integer) As DataTable
    Dim sql As String = "select c.id_contrato,coalesce(a.identifier,'S/A')as clave from CONTRATOS c left join CONTRACT_ACCOUNTS ca on c.id_contrato=ca.id_contrato left join ACCOUNTS a on ca.id_account=a.id_account where c.id_contrato=" & idContract & ";"
    Dim dt As DataTable = con.ConsultarDT(sql)
    Return dt
  End Function

  Private Function getDataAllCharges(ByVal idContract As Integer) As DataTable
    Dim sql As String = "SELECT oc.*,c.nombre FROM OTROS_CARGOS oc inner join TIPOS_CARGOS c on oc.id_tipo_cargo=c.id_tipo_cargo WHERE oc.id_contrato=" & idContract & " AND oc.estatus=1 and oc.id_otro_cargo not in(select id_otro_cargo from DET_EDOS_CARGOS)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    Return dt
  End Function

  Private Function getDescuentos(ByVal id_Contrato As Integer) As Double
    Dim importe As Double = 0
    Dim sql As String = "SELECT coalesce(sum(importe),0) * -1 AS importe FROM dbo.OTROS_CARGOS oc WHERE oc.id_contrato=" & id_Contrato & " AND oc.estatus=1 AND importe<0 and id_otro_cargo not in(select id_otro_cargo from DET_EDOS_CARGOS)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      importe = Val(dt(0)(0).ToString)
    End If
    Return importe
  End Function

  Private Function countBills(ByVal id_contract As Integer) As Integer
    Dim sqlStr = $"select count(*) as cont from ESTADOS_CUENTA where id_contrato=" & id_contract & ";"
    Dim dtBills = con.ConsultarDT(sqlStr)
    Dim cont As Integer = 0

    If dtBills IsNot Nothing AndAlso dtBills.Rows.Count > 0 Then
      cont = Val(dtBills(0)("cont").ToString)
    End If

    Return cont
  End Function

  Private Function getPeriodos(ByVal id_contrato As Integer, ByVal estatus As Integer) As Object
    Dim fecha1 As Date
    Dim fecha2 As Date
    Dim sql As String = ""
    Dim cont As Integer = countBills(id_contrato)

    If estatus = 2 Then
      If cont = 0 Then
        sql = "select TOP 1 periodoB FROM PAGOS WHERE id_contrato=" & id_contrato & " ORDER BY id_pago"
      Else
        sql = "select top 1 periodoB from PAGOS where id_contrato=" & id_contrato & " and id_estado_cuenta > 0 ORDER BY id_pago desc"
      End If

      'sql = "select TOP 1 periodoB FROM PAGOS WHERE id_contrato=" & id_contrato & " ORDER BY id_pago desc"
      'sql = "select top 1 periodoB from PAGOS where id_contrato=" & id_contrato & " and id_estado_cuenta > 0 ORDER BY id_pago desc"
      'sql = "select top 1 dt.periodoB  from(
      'select periodoB, id_pago from PAGOS where id_contrato=" & id_contrato & " and id_estado_cuenta > 0 
      'union 
      'select periodoB,id_pago from PAGOS where id_contrato=" & id_contrato & "
      ')dt order by dt.id_pago desc"
    Else
      sql = "Select TOP 1 periodoB FROM estados_cuenta WHERE id_contrato=" & id_contrato & " ORDER BY id_estado_cuenta DESC"
    End If

    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      fecha1 = dt(0)(0).ToString
      fecha2 = DateAdd(DateInterval.Month, 1, fecha1)
    End If
    Return New With {.PeriodoA = fecha1, .PeriodoB = fecha2}
  End Function
  Private Function getServicios(ByVal id_paquete) As String
    Dim res As String = ""
    Dim sql As String = "Select s.nombre FROM dbo.DETALLE_PAQUETES dp INNER JOIN dbo.SERVICIOS s 
ON s.id_servicio = dp.id_servicio WHERE id_paquete=" & id_paquete & " ORDER BY id_categoria desc"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      For i = 0 To dt.Rows.Count - 1
        res = res & IIf(i = 0, "", " + ") & dt(i)("nombre").ToString
      Next
    End If
    Return res
  End Function

  Private Sub Generar_pdf(ByVal id_estado_cuenta As Integer, ByVal id_contrato As Integer, ByVal path As String)
    Dim sqledo As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtedo As DataTable = con.ConsultarDT(sqledo)
    If dtedo IsNot Nothing AndAlso dtedo.Rows.Count > 0 Then
      Dim fecha As Date = dtedo(0)("fecha").ToString
      Dim grantotal As Double = Val(dtedo(0)("grantotal").ToString)
      Dim saldo_pendiente As Double = Val(dtedo(0)("saldo_pendiente").ToString)
      Dim total_edo As Double = grantotal - saldo_pendiente
      Dim periodoA As Date = dtedo(0)("periodoA").ToString
      Dim periodoB As Date = dtedo(0)("periodoB").ToString

      Dim sqlcli As String = $"SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero,t3.id_contrato,id_paquete FROM (" &
        " SELECT t2.*,upper(p.nombre) AS paquete FROM (" &
        " SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete FROM(" &
        " SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre) AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias" &
        " FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e" &
        " ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato &
        " ) AS t1 INNER JOIN CALLES ca ON ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete)" &
        " AS t3 INNER JOIN dbo.EQUIPOS e INNER JOIN EQUIPOS_TELEFONIA et INNER JOIN LINEAS l" &
        " ON l.id_linea=et.id_linea ON et.id_equipo=e.id_equipo ON e.id_contrato=t3.id_contrato  where e.estatus=1 AND et.estatus=1"
      Dim dtcli As DataTable = con.ConsultarDT(sqlcli)
      If dtcli IsNot Nothing AndAlso dtcli.Rows.Count > 0 Then
        Dim nombre As String = dtcli(0)("nombre").ToString
        Dim contrato As String = dtcli(0)("contrato").ToString
        Dim calle As String = dtcli(0)("calle").ToString
        Dim numext As String = dtcli(0)("numext").ToString
        Dim numint As String = dtcli(0)("numint").ToString
        Dim colonia As String = dtcli(0)("colonia").ToString
        Dim cp As String = dtcli(0)("cp").ToString
        Dim municipio As String = dtcli(0)("municipio").ToString
        Dim estado As String = dtcli(0)("estado").ToString
        Dim referencias As String = dtcli(0)("referencias").ToString
        Dim paquete As String = dtcli(0)("paquete").ToString
        Dim numero As String = dtcli(0)("numero").ToString
        Dim id_paquete As Integer = Val(dtcli(0)("id_paquete").ToString)
        Dim servicios As String = getServicios(id_paquete)



        Dim ruta As String = path & "\EstadoCuenta(" & id_estado_cuenta.ToString & ").pdf "
        Dim oDoc As New iTextSharp.text.Document(PageSize.LETTER, 50, 50, 50, 50)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim linea As PdfContentByte
        Dim rectangulo As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Try
          pdfw = PdfWriter.GetInstance(oDoc, New FileStream(ruta,
                    FileMode.Create, FileAccess.Write, FileShare.None))

          Me.PageState = New CustomPageState()
          ''//Wire our event handler and pass in the page state
          pdfw.PageEvent = New MyCustomPdfEvent(Me.PageState)



          'Apertura del documento.
          oDoc.Open()
          cb = pdfw.DirectContent
          linea = pdfw.DirectContent
          rectangulo = pdfw.DirectContent

          'Agregamos una pagina.
          oDoc.NewPage()

          cb.BeginText()
          fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
          cb.SetFontAndSize(fuente, 10) 'fuente definida en la linea anterior y tamaño

          Dim f10 As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f10.SetColor(2, 51, 130)

          Dim f10Bold As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10Bold.SetColor(2, 51, 130)


          Dim f14 As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f14.SetColor(2, 51, 130)


          'HEADER
          Dim tblHeader As New PdfPTable(3)
          tblHeader.HorizontalAlignment = 0
          tblHeader.LockedWidth = True
          tblHeader.TotalWidth = 540.0F
          tblHeader.DefaultCell.Border = PdfPCell.NO_BORDER
          tblHeader.DefaultCell.MinimumHeight = 12
          tblHeader.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeader.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblHeader.SetWidthPercentage({140.0F, 100.0F, 300.0F}, PageSize.LETTER)


          'IMAGEN
          Dim imagen As iTextSharp.text.Image 'declaración de imagen
          imagen = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png") 'nombre y ruta de la imagen a insertar
          imagen.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)

          tblHeader.AddCell(imagen)
          tblHeader.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))

          Dim cellInfoEmpresa As New PdfPTable(1)
          cellInfoEmpresa.DefaultCell.Border = PdfPCell.NO_BORDER
          cellInfoEmpresa.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT

          cellInfoEmpresa.AddCell(New Phrase("Comunícalo de México S.A. de C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("Domicilio Fiscal: CONVENTO DE CHURUBUSCO NO. 4,", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("COL. JARDINES DE SANTA MÓNICA, MPIO. TLALNEPANTLA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("DE BAZ, ESTADO DE MÉXICO, C.P. 54050", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("RFC: CME0806162SA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim nesthousing As New PdfPCell(cellInfoEmpresa)
          nesthousing.Border = PdfPCell.NO_BORDER
          nesthousing.Padding = 0F
          nesthousing.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeader.AddCell(nesthousing)

          oDoc.Add(tblHeader)
          oDoc.Add(New Paragraph(" "))


          'INFO CLIENTE
          Dim tblInfoCliente As New PdfPTable(1)
          tblInfoCliente.HorizontalAlignment = 0
          tblInfoCliente.LockedWidth = True
          tblInfoCliente.TotalWidth = 540.0F
          tblInfoCliente.DefaultCell.Border = PdfPCell.NO_BORDER
          tblInfoCliente.DefaultCell.MinimumHeight = 12
          tblInfoCliente.DefaultCell.HorizontalAlignment = 0
          tblInfoCliente.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblInfoCliente.SetWidthPercentage({540.0F}, PageSize.LETTER)


          tblInfoCliente.AddCell(New Phrase(nombre, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(calle, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(referencias & " " & numext & " " & numint, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(colonia, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(municipio & ", " & estado & ", C.P. " & cp, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          oDoc.Add(tblInfoCliente)
          oDoc.Add(New Paragraph(" "))


          Dim tblPeriodo As New PdfPTable(5)
          tblPeriodo.HorizontalAlignment = 0
          tblPeriodo.LockedWidth = True
          tblPeriodo.TotalWidth = 540.0F
          tblPeriodo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblPeriodo.DefaultCell.MinimumHeight = 12
          tblPeriodo.DefaultCell.HorizontalAlignment = 0
          tblPeriodo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblPeriodo.DefaultCell.PaddingLeft = 12.0F
          tblPeriodo.SetWidthPercentage({150.0F, 80.0F, 40.0F, 125.0F, 145.0F}, PageSize.LETTER)


          Dim cellPeriodo3 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPeriodo3.Border = PdfPCell.BOTTOM_BORDER
          cellPeriodo3.BorderWidthBottom = 2
          cellPeriodo3.PaddingTop = 0
          cellPeriodo3.HorizontalAlignment = 0
          cellPeriodo3.Colspan = 5
          cellPeriodo3.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellPeriodo3)
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          tblPeriodo.AddCell(New Phrase("MES DE FACTURACIÓN", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(MonthName(periodoA.Month).ToUpper, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          Dim cell1periodo2 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cell1periodo2.Border = PdfPCell.RIGHT_BORDER
          cell1periodo2.BorderWidthRight = 2
          cell1periodo2.HorizontalAlignment = 0
          cell1periodo2.BorderColorRight = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))
          tblPeriodo.AddCell(cell1periodo2)

          tblPeriodo.AddCell(New Phrase("TELÉFONO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(numero, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          tblPeriodo.AddCell(New Phrase("FORMA DE PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase("EFECTIVO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("CONTRATO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(contrato, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("TOTAL A PAGAR", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(total_edo, 2), f10))




          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("PAGAR ANTES DE", f10Bold))
          tblPeriodo.AddCell(New Phrase(periodoA.ToString("dd/MM/yyyy"), f10Bold))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)))

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("SALDO VENCIDO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), f10))

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim cellEdocta As New PdfPCell(New Phrase("ESTADO DE CUENTA", f10))
          cellEdocta.Border = PdfPCell.BOTTOM_BORDER
          cellEdocta.BorderWidthBottom = 2
          cellEdocta.PaddingTop = 12.0F
          cellEdocta.PaddingBottom = 5.0F
          cellEdocta.HorizontalAlignment = 1
          cellEdocta.Colspan = 5
          cellEdocta.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellEdocta)

          Dim cellServiciosContratados As New PdfPCell(New Phrase("Servicios contratados", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServiciosContratados.Border = PdfPCell.BOTTOM_BORDER
          cellServiciosContratados.BorderWidthBottom = 2
          cellServiciosContratados.PaddingTop = 5.0F
          cellServiciosContratados.PaddingBottom = 5.0F
          cellServiciosContratados.HorizontalAlignment = 0
          cellServiciosContratados.Colspan = 5
          cellServiciosContratados.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellServiciosContratados)

          Dim cellPaqueteContratado As New PdfPCell(New Phrase(paquete, f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 5.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPaqueteContratado)

          Dim cellServicios As New PdfPCell(New Phrase(servicios, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServicios.Border = PdfPCell.NO_BORDER
          cellServicios.BorderWidth = 0
          cellServicios.PaddingTop = 0
          cellServicios.HorizontalAlignment = 0
          cellServicios.Colspan = 4
          cellServicios.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellServicios)
          tblPeriodo.AddCell(New Phrase(FormatCurrency(total_edo, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))



          Dim celltotal As New PdfPCell(New Phrase("TOTAL A PAGAR " & FormatCurrency(grantotal, 2), f14))
          celltotal.Border = PdfPCell.NO_BORDER
          celltotal.BorderWidth = 0
          celltotal.PaddingTop = 10.0F
          celltotal.PaddingLeft = 12.0F
          celltotal.HorizontalAlignment = 0
          celltotal.Colspan = 2
          celltotal.BorderColor = Color.WHITE

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(celltotal)

          Dim celltotalLetra As New PdfPCell(New Phrase("(" & totalLetra(grantotal) & ")", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          celltotalLetra.Border = PdfPCell.NO_BORDER
          celltotalLetra.BorderWidth = 0
          celltotalLetra.PaddingTop = 0
          celltotalLetra.PaddingLeft = 12.0F
          celltotalLetra.HorizontalAlignment = 0
          celltotalLetra.Colspan = 2
          celltotalLetra.BorderColor = Color.WHITE

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(celltotalLetra)

          tblPeriodo.AddCell(cellPeriodo3)

          Dim cellFormasPago As New PdfPCell(New Phrase("FORMAS DE PAGO", f10))
          cellFormasPago.Border = PdfPCell.NO_BORDER
          cellFormasPago.BorderWidthBottom = 0
          cellFormasPago.PaddingTop = 10.0F
          cellFormasPago.HorizontalAlignment = 0
          cellFormasPago.Colspan = 5
          cellFormasPago.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellFormasPago)

          Dim cellFormasPago2 As New PdfPCell(New Phrase("BANCO: SCOTIABANK INVERLAT", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago2.Border = PdfPCell.NO_BORDER
          cellFormasPago2.BorderWidthBottom = 0
          cellFormasPago2.PaddingTop = 10.0F
          cellFormasPago2.HorizontalAlignment = 0
          cellFormasPago2.Colspan = 5
          cellFormasPago2.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago2)

          Dim cellFormasPago3 As New PdfPCell(New Phrase("BENEFICIARIO: COMUNICALO DE MÉXICO, S.A. DE C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago3.Border = PdfPCell.NO_BORDER
          cellFormasPago3.BorderWidthBottom = 0
          cellFormasPago3.PaddingTop = 0.0F
          cellFormasPago3.HorizontalAlignment = 0
          cellFormasPago3.Colspan = 5
          cellFormasPago3.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago3)

          Dim cellFormasPago4 As New PdfPCell(New Phrase("CUENTA: 25600765365", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago4.Border = PdfPCell.NO_BORDER
          cellFormasPago4.BorderWidthBottom = 0
          cellFormasPago4.PaddingTop = 0.0F
          cellFormasPago4.HorizontalAlignment = 0
          cellFormasPago4.Colspan = 5
          cellFormasPago4.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago4)

          Dim cellFormasPago5 As New PdfPCell(New Phrase("PAGOS EN OXXO: 5579225041383467", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago5.Border = PdfPCell.NO_BORDER
          cellFormasPago5.BorderWidthBottom = 0
          cellFormasPago5.PaddingTop = 0.0F
          cellFormasPago5.HorizontalAlignment = 0
          cellFormasPago5.Colspan = 5
          cellFormasPago5.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago5)

          Dim cellFormasPago6 As New PdfPCell(New Phrase("Evite molestias pague su servicio a tiempo", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.ITALIC, Color.BLACK)))
          cellFormasPago6.Border = PdfPCell.NO_BORDER
          cellFormasPago6.BorderWidthBottom = 0
          cellFormasPago6.PaddingTop = 5.0F
          cellFormasPago6.HorizontalAlignment = 0
          cellFormasPago6.Colspan = 5
          cellFormasPago6.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago6)

          Dim cellNota As New PdfPCell(New Phrase("NOTA: LE PEDIMOS QUE POR UNICA OCASIÓN ENVÍE SU COMPROBANTE DE PAGO POR CORREO ltorres@comunicalo.mx o WhatsApp 5564161055", New Font(iTextSharp.text.Font.HELVETICA, 12.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellNota.Border = PdfPCell.NO_BORDER
          cellNota.BorderWidthBottom = 0
          cellNota.PaddingTop = 15.0F
          cellNota.HorizontalAlignment = 0
          cellNota.Colspan = 5
          cellNota.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellNota)


          Dim cellNota2 As New PdfPCell(New Phrase("ltorres@comunicalo.mx o WhatsApp 5564161055", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellNota2.Border = PdfPCell.NO_BORDER
          cellNota2.BorderWidthBottom = 0
          cellNota2.PaddingTop = 0.0F
          cellNota2.HorizontalAlignment = 0
          cellNota2.Colspan = 5
          cellNota2.BorderColorBottom = Color.WHITE

          'tblPeriodo.AddCell(cellNota2)

          Dim cellGracias As New PdfPCell(New Phrase("¡MUCHAS GRACIAS POR DARNOS LA OPORTUNIDAD DE SERVIRLE!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellGracias.Border = PdfPCell.NO_BORDER
          cellGracias.BorderWidthBottom = 0
          cellGracias.PaddingTop = 25.0F
          cellGracias.PaddingBottom = 60.0F
          cellGracias.HorizontalAlignment = 1
          cellGracias.Colspan = 5
          cellGracias.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellGracias)

          tblPeriodo.AddCell(cellPeriodo3)

          Dim cellPie1 As New PdfPCell(New Phrase("ATENCIÓN A CLIENTES: 5526014010", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie1.Border = PdfPCell.NO_BORDER
          cellPie1.BorderWidthBottom = 0
          cellPie1.PaddingTop = 2.0F
          cellPie1.HorizontalAlignment = 0
          cellPie1.Colspan = 3
          cellPie1.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPie1)

          Dim cellPie2 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie2.Border = PdfPCell.NO_BORDER
          cellPie2.BorderWidthBottom = 0
          cellPie2.PaddingTop = 2.0F
          cellPie2.HorizontalAlignment = 2
          cellPie2.Colspan = 2
          cellPie2.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPie2)

          oDoc.Add(tblPeriodo)


          'Fin del flujo de bytes.
          cb.EndText()
          'Forzamos vaciamiento del buffer.
          pdfw.Flush()
          'Cerramos el documento.
          oDoc.Close()



        Catch ex As Exception
          'Si hubo una excepcion y el archivo existe …
          If File.Exists(ruta) Then
            'Cerramos el documento si esta abierto.
            'Y asi desbloqueamos el archivo para su eliminacion.
            If oDoc.IsOpen Then oDoc.Close()
            '… lo eliminamos de disco.
            File.Delete(ruta)
          End If
          'Throw New Exception("Error al generar archivo PDF (" & ex.Message & ")" & ex.Source)
          MsgBox(ex.Message & "--- " & ex.StackTrace)
          'Dim sqlerror As String = "insert into"
          'Dim sql As String = "insert into netcel..Correos(cliente,mensaje,asunto,estatus,respuesta) values('-1','ERROR AL GENERAR ESTADO DE CUENTA DE COMUNICALO  " & cli_id.ToString & ", MENSAJE:" & ex.Message & "<br/> SOURCE: " & ex.Source & " <br/> STACK TRACE:" & ex.StackTrace & "','ERROR ESTADO DE CUENTA ILOXTELECOM" & cli_id.ToString & "','1','sinfante@mail.ilox.mx')"
          'con.ModRegEli(sql)
          'escribir_log("ERROR AL GENERAR ESTADO DE CUENTA DEL CONTRATO_ID " & idcliente.ToString & ", MENSAJE:" & ex.Message & " SOURCE: " & ex.Source & " STACK TRACE:" & ex.StackTrace)

        Finally
          cb = Nothing
          pdfw = Nothing
          oDoc = Nothing
        End Try


      End If
    End If

  End Sub
  Private Function tiene_telefonia(ByVal id_contrato As Integer) As Boolean
    Dim res As Boolean = True
    Dim sql As String = " SELECT * FROM equipos e INNER JOIN dbo.EQUIPOS_TELEFONIA et" &
 " ON et.id_equipo = e.id_equipo WHERE id_contrato=" & id_contrato & " AND" &
 " e.estatus=1 AND et.estatus=1"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt Is Nothing Or dt.Rows.Count <= 0 Then
      res = False
    End If

    Return res
  End Function

  Private Sub Generar_pdfOXXO2(ByVal id_estado_cuenta As Integer, ByVal id_contrato As Integer, ByVal path As String, ByVal refOxxo As String, ByVal codigoBarraOxxo As String)
    Dim sqledo As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtedo As DataTable = con.ConsultarDT(sqledo)
    If dtedo IsNot Nothing AndAlso dtedo.Rows.Count > 0 Then
      Dim fecha As Date = dtedo(0)("fecha").ToString
      Dim grantotal As Double = Val(dtedo(0)("grantotal").ToString)
      Dim saldo_pendiente As Double = Val(dtedo(0)("saldo_pendiente").ToString)
      Dim total_edo As Double = grantotal - saldo_pendiente
      Dim periodoA As Date = dtedo(0)("periodoA").ToString
      Dim periodoB As Date = dtedo(0)("periodoB").ToString
      Dim totalPlan As Double = Val(dtedo(0)("mensualidad").ToString())
      Dim sqlcli As String = ""

      If tiene_telefonia(id_contrato) Then
        sqlcli = $"SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero,t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM (" &
" SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete FROM(" &
" SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre) AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias" &
" FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e" &
" ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato &
" ) AS t1 INNER JOIN CALLES ca ON ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete)" &
" AS t3 INNER JOIN dbo.EQUIPOS e INNER JOIN EQUIPOS_TELEFONIA et INNER JOIN LINEAS l" &
" ON l.id_linea=et.id_linea ON et.id_equipo=e.id_equipo ON e.id_contrato=t3.id_contrato  where e.estatus=1 AND et.estatus=1"
      Else
        sqlcli = "SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero," &
" t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM" &
" ( SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete,numero" &
" FROM( SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre)" &
 " AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias,telefono AS numero FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c " &
 " INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON" &
 " col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato & " ) AS t1 INNER JOIN CALLES ca ON" &
 " ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete) AS t3"
      End If

      Dim dtcli As DataTable = con.ConsultarDT(sqlcli)
      If dtcli IsNot Nothing AndAlso dtcli.Rows.Count > 0 Then
        Dim nombre As String = dtcli(0)("nombre").ToString
        Dim contrato As String = dtcli(0)("contrato").ToString
        Dim calle As String = dtcli(0)("calle").ToString
        Dim numext As String = dtcli(0)("numext").ToString
        Dim numint As String = dtcli(0)("numint").ToString
        Dim colonia As String = dtcli(0)("colonia").ToString
        Dim cp As String = dtcli(0)("cp").ToString
        Dim municipio As String = dtcli(0)("municipio").ToString
        Dim estado As String = dtcli(0)("estado").ToString
        Dim referencias As String = dtcli(0)("referencias").ToString
        Dim paquete As String = dtcli(0)("paquete").ToString
        Dim numero As String = dtcli(0)("numero").ToString
        Dim id_paquete As Integer = Val(dtcli(0)("id_paquete").ToString)
        Dim servicios As String = getServicios(id_paquete)



        Dim ruta As String = path & "\EstadoCuenta(" & id_estado_cuenta.ToString & ").pdf "
        Dim oDoc As New iTextSharp.text.Document(PageSize.LETTER, 50, 50, 50, 50)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim linea As PdfContentByte
        Dim rectangulo As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Try
          pdfw = PdfWriter.GetInstance(oDoc, New FileStream(ruta,
                    FileMode.Create, FileAccess.Write, FileShare.None))

          Me.PageState = New CustomPageState()
          ''//Wire our event handler and pass in the page state
          pdfw.PageEvent = New MyCustomPdfEvent(Me.PageState)



          'Apertura del documento.
          oDoc.Open()
          cb = pdfw.DirectContent
          linea = pdfw.DirectContent
          rectangulo = pdfw.DirectContent

          'Agregamos una pagina.
          oDoc.NewPage()

          cb.BeginText()
          fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
          cb.SetFontAndSize(fuente, 10) 'fuente definida en la linea anterior y tamaño

          Dim f10 As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f10.SetColor(2, 51, 130)

          Dim f10Bold As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10Bold.SetColor(2, 51, 130)


          Dim f14 As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f14.SetColor(2, 51, 130)


          'HEADER
          Dim tblHeader As New PdfPTable(3)
          tblHeader.HorizontalAlignment = 0
          tblHeader.LockedWidth = True
          tblHeader.TotalWidth = 540.0F
          tblHeader.DefaultCell.Border = PdfPCell.NO_BORDER
          tblHeader.DefaultCell.MinimumHeight = 12
          tblHeader.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeader.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblHeader.SetWidthPercentage({140.0F, 100.0F, 300.0F}, PageSize.LETTER)


          'IMAGEN
          Dim imagen As iTextSharp.text.Image 'declaración de imagen
          imagen = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png") 'nombre y ruta de la imagen a insertar
          imagen.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)

          tblHeader.AddCell(imagen)
          tblHeader.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))

          Dim cellInfoEmpresa As New PdfPTable(1)
          cellInfoEmpresa.DefaultCell.Border = PdfPCell.NO_BORDER
          cellInfoEmpresa.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT

          cellInfoEmpresa.AddCell(New Phrase("Comunícalo de México S.A. de C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("Domicilio Fiscal: CONVENTO DE CHURUBUSCO NO. 4,", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("COL. JARDINES DE SANTA MÓNICA, MPIO. TLALNEPANTLA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("DE BAZ, ESTADO DE MÉXICO, C.P. 54050", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("RFC: CME0806162SA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim nesthousing As New PdfPCell(cellInfoEmpresa)
          nesthousing.Border = PdfPCell.NO_BORDER
          nesthousing.Padding = 0F
          nesthousing.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeader.AddCell(nesthousing)

          oDoc.Add(tblHeader)
          oDoc.Add(New Paragraph(" "))


          'INFO CLIENTE
          Dim tblInfoCliente As New PdfPTable(1)
          tblInfoCliente.HorizontalAlignment = 0
          tblInfoCliente.LockedWidth = True
          tblInfoCliente.TotalWidth = 540.0F
          tblInfoCliente.DefaultCell.Border = PdfPCell.NO_BORDER
          tblInfoCliente.DefaultCell.MinimumHeight = 12
          tblInfoCliente.DefaultCell.HorizontalAlignment = 0
          tblInfoCliente.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblInfoCliente.SetWidthPercentage({540.0F}, PageSize.LETTER)


          tblInfoCliente.AddCell(New Phrase(nombre, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(calle, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(referencias & " " & numext & " " & numint, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(colonia, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(municipio & ", " & estado & ", C.P. " & cp, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          oDoc.Add(tblInfoCliente)
          oDoc.Add(New Paragraph(" "))


          Dim tblPeriodo As New PdfPTable(5)
          tblPeriodo.HorizontalAlignment = 0
          tblPeriodo.LockedWidth = True
          tblPeriodo.TotalWidth = 540.0F
          tblPeriodo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblPeriodo.DefaultCell.MinimumHeight = 12
          tblPeriodo.DefaultCell.HorizontalAlignment = 0
          tblPeriodo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblPeriodo.DefaultCell.PaddingLeft = 12.0F
          tblPeriodo.SetWidthPercentage({150.0F, 80.0F, 40.0F, 125.0F, 145.0F}, PageSize.LETTER)


          Dim cellPeriodo3 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPeriodo3.Border = PdfPCell.BOTTOM_BORDER
          cellPeriodo3.BorderWidthBottom = 2
          cellPeriodo3.PaddingTop = 0
          cellPeriodo3.HorizontalAlignment = 0
          cellPeriodo3.Colspan = 5
          cellPeriodo3.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellPeriodo3)
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          tblPeriodo.AddCell(New Phrase("MES DE FACTURACIÓN", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(MonthName(periodoA.Month).ToUpper, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          Dim cell1periodo2 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cell1periodo2.Border = PdfPCell.RIGHT_BORDER
          cell1periodo2.BorderWidthRight = 2
          cell1periodo2.HorizontalAlignment = 0
          cell1periodo2.BorderColorRight = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))
          tblPeriodo.AddCell(cell1periodo2)

          tblPeriodo.AddCell(New Phrase("TELÉFONO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(numero, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          tblPeriodo.AddCell(New Phrase("FORMA DE PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase("EFECTIVO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("CONTRATO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(contrato, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("TOTAL A PAGAR", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(grantotal, 2), f10))




          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("PAGAR ANTES DE", f10Bold))
          tblPeriodo.AddCell(New Phrase(periodoA.ToString("dd/MM/yyyy"), f10Bold))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)))

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("SALDO VENCIDO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), f10))

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim cellEdocta As New PdfPCell(New Phrase("ESTADO DE CUENTA", f10))
          cellEdocta.Border = PdfPCell.BOTTOM_BORDER
          cellEdocta.BorderWidthBottom = 2
          cellEdocta.PaddingTop = 12.0F
          cellEdocta.PaddingBottom = 5.0F
          cellEdocta.HorizontalAlignment = 1
          cellEdocta.Colspan = 5
          cellEdocta.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellEdocta)

          Dim cellServiciosContratados As New PdfPCell(New Phrase("Servicios contratados", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServiciosContratados.Border = PdfPCell.BOTTOM_BORDER
          cellServiciosContratados.BorderWidthBottom = 2
          cellServiciosContratados.PaddingTop = 5.0F
          cellServiciosContratados.PaddingBottom = 5.0F
          cellServiciosContratados.HorizontalAlignment = 0
          cellServiciosContratados.Colspan = 5
          cellServiciosContratados.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellServiciosContratados)

          Dim cellPaqueteContratado As New PdfPCell(New Phrase(paquete, f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 5.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPaqueteContratado)

          Dim cellServicios As New PdfPCell(New Phrase(servicios, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServicios.Border = PdfPCell.NO_BORDER
          cellServicios.BorderWidth = 0
          cellServicios.PaddingTop = 0
          cellServicios.HorizontalAlignment = 0
          cellServicios.Colspan = 4
          cellServicios.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellServicios)
          'tblPeriodo.AddCell(New Phrase(FormatCurrency(total_edo, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(totalPlan, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim cellPending As New PdfPCell(New Phrase("SALDO PENDIENTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPending.Border = PdfPCell.NO_BORDER
          cellPending.BorderWidth = 0
          cellPending.PaddingTop = 0
          cellPending.HorizontalAlignment = 0
          cellPending.Colspan = 4
          cellPending.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellPending)
          tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim dtCharges As DataTable = getDataBillCharges(id_estado_cuenta)

          If dtCharges IsNot Nothing AndAlso dtCharges.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("OTROS CARGOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtCharges.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtCharges.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtCharges.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim dtDiscount As DataTable = getDataBillDiscounts(id_estado_cuenta)

          If dtDiscount IsNot Nothing AndAlso dtDiscount.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("DESCUENTOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtDiscount.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtDiscount.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtDiscount.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim celltotal As New PdfPCell(New Phrase("TOTAL A PAGAR " & FormatCurrency(grantotal, 2), f14))
          celltotal.Border = PdfPCell.NO_BORDER
          celltotal.BorderWidth = 0
          celltotal.PaddingTop = 10.0F
          celltotal.PaddingLeft = 12.0F
          celltotal.HorizontalAlignment = 0
          celltotal.Colspan = 2
          celltotal.BorderColor = Color.WHITE

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(celltotal)

          Dim celltotalLetra As New PdfPCell(New Phrase("(" & totalLetra(grantotal) & ")", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          celltotalLetra.Border = PdfPCell.NO_BORDER
          celltotalLetra.BorderWidth = 0
          celltotalLetra.PaddingTop = 0
          celltotalLetra.PaddingLeft = 12.0F
          celltotalLetra.HorizontalAlignment = 0
          celltotalLetra.Colspan = 2
          celltotalLetra.BorderColor = Color.WHITE

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(celltotalLetra)

          tblPeriodo.AddCell(cellPeriodo3)

          Dim cellFormasPago As New PdfPCell(New Phrase("FORMAS DE PAGO", f10))
          cellFormasPago.Border = PdfPCell.NO_BORDER
          cellFormasPago.BorderWidthBottom = 0
          cellFormasPago.PaddingTop = 10.0F
          cellFormasPago.HorizontalAlignment = 1
          cellFormasPago.Colspan = 5
          cellFormasPago.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellFormasPago)

          Dim cellDeposito As New PdfPCell(New Phrase("DEPOSITO BANCARIO:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellDeposito.Border = PdfPCell.NO_BORDER
          cellDeposito.BorderWidthBottom = 0
          cellDeposito.PaddingTop = 10.0F
          cellDeposito.HorizontalAlignment = 0
          cellDeposito.Colspan = 3
          cellDeposito.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellDeposito)

          Dim cellTransfer As New PdfPCell(New Phrase("TRANSFERENCIA ELECTRÓNICA:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTransfer.Border = PdfPCell.NO_BORDER
          cellTransfer.BorderWidthBottom = 0
          cellTransfer.PaddingTop = 10.0F
          cellTransfer.HorizontalAlignment = 1
          cellTransfer.Colspan = 2
          cellTransfer.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellTransfer)

          Dim cellFormasPago2 As New PdfPCell(New Phrase("BANCO: SCOTIABANK INVERLAT", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago2.Border = PdfPCell.NO_BORDER
          cellFormasPago2.BorderWidthBottom = 0
          cellFormasPago2.PaddingTop = 2.0F
          cellFormasPago2.HorizontalAlignment = 0
          cellFormasPago2.Colspan = 3
          cellFormasPago2.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago2)


          Dim cellClabe As New PdfPCell(New Phrase("CLABE INTERBANCARIA: 044180256007653656", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellClabe.Border = PdfPCell.NO_BORDER
          cellClabe.BorderWidthBottom = 0
          cellClabe.PaddingTop = 2.0F
          cellClabe.HorizontalAlignment = 2
          cellClabe.Colspan = 2
          cellClabe.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellClabe)

          Dim cellFormasPago3 As New PdfPCell(New Phrase("BENEFICIARIO: COMUNICALO DE MÉXICO, S.A. DE C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago3.Border = PdfPCell.NO_BORDER
          cellFormasPago3.BorderWidthBottom = 0
          cellFormasPago3.PaddingTop = 0.0F
          cellFormasPago3.HorizontalAlignment = 0
          cellFormasPago3.Colspan = 5
          cellFormasPago3.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago3)

          Dim cellFormasPago4 As New PdfPCell(New Phrase("CUENTA: 25600765365", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago4.Border = PdfPCell.NO_BORDER
          cellFormasPago4.BorderWidthBottom = 0
          cellFormasPago4.PaddingTop = 0.0F
          cellFormasPago4.HorizontalAlignment = 0
          cellFormasPago4.Colspan = 5
          cellFormasPago4.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago4)


          Dim cellEspacio As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellEspacio.Border = PdfPCell.NO_BORDER
          cellEspacio.BorderWidthBottom = 0
          cellEspacio.PaddingTop = 5.0F
          cellEspacio.HorizontalAlignment = 1
          cellEspacio.Colspan = 5
          cellEspacio.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          'Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS PAYNET OPENPAY", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellPagoOxxo.Border = PdfPCell.NO_BORDER
          cellPagoOxxo.BorderWidthBottom = 0
          cellPagoOxxo.PaddingTop = 5.0F
          cellPagoOxxo.HorizontalAlignment = 1
          cellPagoOxxo.Colspan = 5
          cellPagoOxxo.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPagoOxxo)



          If refOxxo.Trim <> "" And codigoBarraOxxo <> "" Then
            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3
            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
            Dim imgOxxo As iTextSharp.text.Image 'declaración de imagen
            imgOxxo = iTextSharp.text.Image.GetInstance(codigoBarraOxxo) 'nombre y ruta de la imagen a insertar
            'imagen.ScalePercent(50) 'escala al tamaño de la imagen

            Dim cellimgOxxo As New PdfPCell(imgOxxo)
            cellimgOxxo.Border = PdfPCell.NO_BORDER
            cellimgOxxo.BorderWidthBottom = 0
            cellimgOxxo.PaddingTop = 5.0F
            cellimgOxxo.HorizontalAlignment = 1
            cellimgOxxo.Colspan = 5
            cellimgOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellimgOxxo)


            Dim cellrefOxxo As New PdfPCell(New Phrase(refOxxo))
            cellrefOxxo.Border = PdfPCell.NO_BORDER
            cellrefOxxo.BorderWidthBottom = 0
            cellrefOxxo.PaddingTop = 5.0F
            cellrefOxxo.HorizontalAlignment = 1
            cellrefOxxo.Colspan = 5
            cellrefOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellrefOxxo)


          End If



          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)


          Dim cellTiendas As New PdfPCell(New Phrase("TIENDAS PARA REALIZAR SU PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTiendas.Border = PdfPCell.NO_BORDER
          cellTiendas.BorderWidthBottom = 0
          cellTiendas.PaddingTop = 5.0F
          cellTiendas.HorizontalAlignment = 1
          cellTiendas.Colspan = 5
          cellTiendas.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellTiendas)

          Dim imagenTiendas As iTextSharp.text.Image 'declaración de imagen para las tiendas.

          'imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/tiendasopen.jpg") 'nombre y ruta de la imagen a insertar
          imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/tiendas.jpeg") 'nombre y ruta de la imagen a insertar
          'imagenTiendas.ScalePercent(44) 'escala al tamaño de la imagen openpay
          imagenTiendas.ScalePercent(50)
          Dim cellimgTiendas As New PdfPCell(imagenTiendas)
          cellimgTiendas.Border = PdfPCell.NO_BORDER
          cellimgTiendas.BorderWidthBottom = 0
          cellimgTiendas.PaddingTop = 5.0F
          cellimgTiendas.HorizontalAlignment = 1  ' 0 para open pay
          cellimgTiendas.Colspan = 5
          cellimgTiendas.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellimgTiendas)

          Dim cellInstrucciones As New PdfPCell(New Phrase("INSTRUCCIONES PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInstrucciones.Border = PdfPCell.NO_BORDER
          cellInstrucciones.BorderWidthBottom = 0
          cellInstrucciones.PaddingTop = 20.0F
          cellInstrucciones.PaddingBottom = 10.0F
          cellInstrucciones.HorizontalAlignment = 1
          cellInstrucciones.Colspan = 5
          cellInstrucciones.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellInstrucciones)

          Dim cellPasps As New PdfPCell(New Phrase("1.- DEBES ELEGIR LA TIENDA QUE MÁS TE CONVENGA ENTRE LAS CADENAS INDICADAS (SOLO SE PUEDE PAGAR EN ESAS TIENDAS).", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          Dim boldText As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          'Dim compania As Chunk = New Chunk("PAYNET OPENPAY", boldText)
          Dim compania As Chunk = New Chunk("CONEKTA", boldText)
          compania.SetUnderline(0.4, -0.8)
          Dim instruccion As String = "2.- AL ACERCARSE AL MOSTRADOR, DEBERÁ MENCIONAR QUE VIENE A PAGAR "
          Dim ph As Phrase = New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          ph.Add(compania)
          ph.Add(",Y MOSTRAR AL CAJERO EL CÓDIGO DE BARRAS O DICTAR LOS NÚMEROS QUE APARECEN EN LA REFERENCIA.")

          cellPasps = New PdfPCell(ph)
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          instruccion = "3.- UNA VEZ REALIZADO EL PAGO EN EFECTIVO, ENVIAREMOS UNA NOTIFICACIÓN DE PAGO EN TIEMPO REAL A SU CORREO Y ¡LISTO!"
          cellPasps = New PdfPCell(New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          'tblPeriodo.AddCell(cellNota2)

          Dim cellGracias As New PdfPCell(New Phrase("¡MUCHAS GRACIAS POR DARNOS LA OPORTUNIDAD DE SERVIRLE!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellGracias.Border = PdfPCell.NO_BORDER
          cellGracias.BorderWidthBottom = 0
          cellGracias.PaddingTop = 20.0F
          cellGracias.PaddingBottom = 30.0F
          cellGracias.HorizontalAlignment = 1
          cellGracias.Colspan = 5
          cellGracias.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellGracias)

          tblPeriodo.AddCell(cellPeriodo3)

          Dim cellPie1 As New PdfPCell(New Phrase("ATENCIÓN A CLIENTES: 5526014010", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie1.Border = PdfPCell.NO_BORDER
          cellPie1.BorderWidthBottom = 0
          cellPie1.PaddingTop = 2.0F
          cellPie1.HorizontalAlignment = 0
          cellPie1.Colspan = 3
          cellPie1.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPie1)

          Dim cellPie2 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie2.Border = PdfPCell.NO_BORDER
          cellPie2.BorderWidthBottom = 0
          cellPie2.PaddingTop = 2.0F
          cellPie2.HorizontalAlignment = 2
          cellPie2.Colspan = 2
          cellPie2.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPie2)

          oDoc.Add(tblPeriodo)


          'Fin del flujo de bytes.
          cb.EndText()
          'Forzamos vaciamiento del buffer.
          pdfw.Flush()
          'Cerramos el documento.
          oDoc.Close()



        Catch ex As Exception
          'Si hubo una excepcion y el archivo existe …
          If File.Exists(ruta) Then
            'Cerramos el documento si esta abierto.
            'Y asi desbloqueamos el archivo para su eliminacion.
            If oDoc.IsOpen Then oDoc.Close()
            '… lo eliminamos de disco.
            File.Delete(ruta)
          End If
          'Throw New Exception("Error al generar archivo PDF (" & ex.Message & ")" & ex.Source)
          MsgBox(ex.Message & "--- " & ex.StackTrace)
          'Dim sqlerror As String = "insert into"
          'Dim sql As String = "insert into netcel..Correos(cliente,mensaje,asunto,estatus,respuesta) values('-1','ERROR AL GENERAR ESTADO DE CUENTA DE COMUNICALO  " & cli_id.ToString & ", MENSAJE:" & ex.Message & "<br/> SOURCE: " & ex.Source & " <br/> STACK TRACE:" & ex.StackTrace & "','ERROR ESTADO DE CUENTA ILOXTELECOM" & cli_id.ToString & "','1','sinfante@mail.ilox.mx')"
          'con.ModRegEli(sql)
          'escribir_log("ERROR AL GENERAR ESTADO DE CUENTA DEL CONTRATO_ID " & idcliente.ToString & ", MENSAJE:" & ex.Message & " SOURCE: " & ex.Source & " STACK TRACE:" & ex.StackTrace)

        Finally
          cb = Nothing
          pdfw = Nothing
          oDoc = Nothing
        End Try


      End If
    End If

  End Sub

  Private Sub Generar_pdfOXXO_v2_2(ByVal id_estado_cuenta As Integer,
                                  ByVal id_contrato As Integer,
                                  ByVal path As String,
                                  ByVal refOxxo As String,
                                  ByVal codigoBarraOxxo As String)

    ' Datos dummy. Sustituir después por datos de BD.
    Dim mesFacturacion As String = "OCTUBRE"
    Dim contrato As String = "06155"
    Dim telefono As String = "7791425119"
    Dim totalPagar As Decimal = 750D
    Dim saldoVencido As Decimal = 400D
    Dim cargoPagoTardio As Decimal = 50D
    Dim totalDespuesFecha As Decimal = 800D
    Dim fechaLimite As New Date(2023, 10, 12)
    Dim clabe As String = "646180558500043128"
    Dim banco As String = "STP"
    Dim beneficiario As String = "Comunícalo de México, S.A. de C.V."
    Dim nombreCliente As String = "MARGARITA DEL CARMEN BARROSO GONZALES"
    Dim direccionCliente1 As String = "PRIV. VILAFRANCA 112 4-B, BOSQUES DE IBIZA"
    Dim direccionCliente2 As String = "TIZAYUCA, HIDALGO, C.P. 43815"
    Dim planActual As String = "PLAN COMUNICALO D30"
    Dim conceptoActual As String = "Servicio Telefonía Ilimitada + Servicio Internet 30 Mbps"
    Dim importeActual As Decimal = 350D
    Dim mesAnterior As String = "SEPTIEMBRE"
    Dim conceptoMesAnterior As String = "Mensualidad Septiembre — no pagada"
    Dim importeMesAnterior As Decimal = 350D

    If String.IsNullOrWhiteSpace(refOxxo) Then refOxxo = "1010102677978684"

    Dim ruta As String = System.IO.Path.Combine(path, "EstadoCuentaV2(" & id_estado_cuenta.ToString() & ").pdf")

    Dim azul As New Color(20, 67, 126)
    Dim azulOscuro As New Color(14, 47, 94)
    Dim azulClaro As New Color(239, 244, 250)
    Dim grisClaro As New Color(246, 247, 249)
    Dim grisBorde As New Color(218, 222, 228)
    Dim grisTexto As New Color(90, 96, 105)
    Dim amarilloClaro As New Color(252, 247, 226)
    Dim amarillo As New Color(190, 144, 37)

    Dim f7 As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7.0F, iTextSharp.text.Font.NORMAL, grisTexto)
    Dim f7b As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
    Dim f8 As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, grisTexto)
    Dim f8k As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)
    Dim f8b As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.BOLD, azulOscuro)
    Dim f8bw As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.BOLD, Color.WHITE)
    Dim f9b As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 9.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
    Dim f10b As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, azulOscuro)
    Dim f10bw As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.WHITE)
    Dim f12bw As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 12.0F, iTextSharp.text.Font.BOLD, Color.WHITE)
    Dim f13b As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 13.0F, iTextSharp.text.Font.BOLD, azulOscuro)
    Dim f16bw As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 16.0F, iTextSharp.text.Font.BOLD, Color.WHITE)
    Dim f18b As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 18.0F, iTextSharp.text.Font.BOLD, azulOscuro)
    Dim f20bw As New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 20.0F, iTextSharp.text.Font.BOLD, Color.WHITE)

    Dim oDoc As New Document(PageSize.LETTER, 36, 36, 28, 28)
    Dim pdfw As PdfWriter = Nothing

    Try
      pdfw = PdfWriter.GetInstance(oDoc, New FileStream(ruta, FileMode.Create, FileAccess.Write, FileShare.None))
      Me.PageState = New CustomPageState()
      pdfw.PageEvent = New MyCustomPdfEvent(Me.PageState)
      oDoc.Open()

      '======================== PÁGINA 1 ========================
      Dim header As New PdfPTable(2)
      header.TotalWidth = 540.0F : header.LockedWidth = True
      header.SetWidths(New Single() {220.0F, 320.0F})

      Dim logo As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
      logo.ScaleToFit(170.0F, 70.0F)
      Dim cLogo As New PdfPCell(logo)
      cLogo.Border = PdfPCell.NO_BORDER : cLogo.PaddingBottom = 8.0F
      header.AddCell(cLogo)

      Dim title As New PdfPTable(1)
      title.WidthPercentage = 100
      Dim ct As New PdfPCell(New Phrase("ESTADO DE CUENTA", f18b))
      ct.Border = PdfPCell.NO_BORDER : ct.HorizontalAlignment = Element.ALIGN_RIGHT
      title.AddCell(ct)
      Dim subPh As New Phrase()
      subPh.Add(New Chunk("Mes de facturación: ", f8))
      subPh.Add(New Chunk(mesFacturacion, f8b))
      subPh.Add(New Chunk("  ·  Contrato: ", f8))
      subPh.Add(New Chunk(contrato, f8b))
      Dim cs As New PdfPCell(subPh)
      cs.Border = PdfPCell.NO_BORDER : cs.HorizontalAlignment = Element.ALIGN_RIGHT
      title.AddCell(cs)
      Dim cTitle As New PdfPCell(title)
      cTitle.Border = PdfPCell.NO_BORDER : cTitle.VerticalAlignment = Element.ALIGN_MIDDLE
      header.AddCell(cTitle)

      Dim hLine As New PdfPCell(New Phrase(""))
      hLine.Colspan = 2 : hLine.Border = PdfPCell.BOTTOM_BORDER
      hLine.BorderColorBottom = azulOscuro : hLine.BorderWidthBottom = 2.2F
      header.AddCell(hLine)
      oDoc.Add(header)
      oDoc.Add(New Paragraph(" ", f7))

      Dim summary As New PdfPTable(3)
      summary.TotalWidth = 540.0F : summary.LockedWidth = True
      summary.SetWidths(New Single() {210.0F, 185.0F, 145.0F})

      Dim cTotal As New PdfPCell()
      cTotal.BackgroundColor = azul : cTotal.Border = PdfPCell.NO_BORDER : cTotal.Padding = 11.0F
      cTotal.AddElement(New Phrase("TOTAL A PAGAR", f8bw))
      cTotal.AddElement(New Phrase(FormatCurrency(totalPagar, 2), f20bw))
      cTotal.AddElement(New Phrase("Incluye cargos del mes y saldo anterior acumulado", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7.0F, iTextSharp.text.Font.NORMAL, Color.WHITE)))
      summary.AddCell(cTotal)

      Dim cFecha As New PdfPCell()
      cFecha.BackgroundColor = grisClaro : cFecha.Border = PdfPCell.NO_BORDER : cFecha.Padding = 11.0F
      cFecha.AddElement(New Phrase("FECHA LÍMITE DE PAGO", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.BOLD, amarillo)))
      cFecha.AddElement(New Phrase(fechaLimite.ToString("dd / MMM / yyyy").ToUpper(), f13b))
      cFecha.AddElement(New Phrase("Pague a tiempo y evite recargos", f7))
      summary.AddCell(cFecha)

      Dim cSaldo As New PdfPCell()
      cSaldo.BackgroundColor = azulOscuro : cSaldo.Border = PdfPCell.NO_BORDER : cSaldo.Padding = 11.0F
      cSaldo.AddElement(New Phrase("SALDO VENCIDO", f8bw))
      cSaldo.AddElement(New Phrase(FormatCurrency(saldoVencido, 2), f16bw))
      cSaldo.AddElement(New Phrase("REQUIERE PAGO INMEDIATO", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7.0F, iTextSharp.text.Font.BOLD, New Color(255, 218, 95))))
      summary.AddCell(cSaldo)
      oDoc.Add(summary)
      oDoc.Add(New Paragraph(" ", f7))

      Dim warning As New PdfPTable(2)
      warning.TotalWidth = 540.0F : warning.LockedWidth = True
      warning.SetWidths(New Single() {430.0F, 110.0F})
      Dim warningText As New PdfPCell(New Phrase("Si no paga antes del " & fechaLimite.ToString("dd/MM/yyyy") & ", se sumará un cargo por pago tardío de " & FormatCurrency(cargoPagoTardio, 2) & " y su servicio podrá ser suspendido.", f8k))
      warningText.BackgroundColor = amarilloClaro : warningText.BorderColor = New Color(232, 220, 170) : warningText.Padding = 12.0F
      warning.AddCell(warningText)
      Dim warningAmount As New PdfPCell()
      warningAmount.BackgroundColor = azul : warningAmount.Border = PdfPCell.NO_BORDER : warningAmount.Padding = 8.0F
      warningAmount.AddElement(New Paragraph(FormatCurrency(totalDespuesFecha, 2), f16bw) With {.Alignment = Element.ALIGN_CENTER})
      warningAmount.AddElement(New Paragraph("TOTAL DESPUÉS" & Environment.NewLine & "DEL " & fechaLimite.ToString("dd/MM/yyyy"), New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 6.8F, iTextSharp.text.Font.BOLD, Color.WHITE)) With {.Alignment = Element.ALIGN_CENTER})
      warning.AddCell(warningAmount)
      oDoc.Add(warning)
      oDoc.Add(New Paragraph(" ", f7))

      Dim clabeTable As New PdfPTable(3)
      clabeTable.TotalWidth = 540.0F : clabeTable.LockedWidth = True
      clabeTable.SetWidths(New Single() {105.0F, 280.0F, 155.0F})
      Dim clabeLabel As New PdfPCell(New Phrase("CLABE" & Environment.NewLine & "INTERBANCARIA" & Environment.NewLine & "PERSONALIZADA", f8bw))
      clabeLabel.BackgroundColor = azul : clabeLabel.BorderColor = azulOscuro : clabeLabel.Padding = 9.0F : clabeLabel.HorizontalAlignment = Element.ALIGN_CENTER
      clabeTable.AddCell(clabeLabel)
      Dim clabeValue As New PdfPCell()
      clabeValue.BackgroundColor = azulClaro : clabeValue.BorderColor = azulOscuro : clabeValue.Padding = 9.0F
      clabeValue.AddElement(New Phrase("6461 8055 8500 04 3128", New iTextSharp.text.Font(iTextSharp.text.Font.COURIER, 18.0F, iTextSharp.text.Font.BOLD, azulOscuro)))
      clabeValue.AddElement(New Phrase("Esta CLABE es exclusiva de su contrato: su pago se aplica automáticamente, sin referencia adicional.", f7))
      clabeTable.AddCell(clabeValue)
      Dim bankValue As New PdfPCell()
      bankValue.BackgroundColor = azulClaro : bankValue.BorderColor = azulOscuro : bankValue.Padding = 8.0F
      bankValue.AddElement(New Paragraph("Banco: " & banco, f7b) With {.Alignment = Element.ALIGN_RIGHT})
      bankValue.AddElement(New Paragraph("Beneficiario: " & beneficiario, f7) With {.Alignment = Element.ALIGN_RIGHT})
      clabeTable.AddCell(bankValue)
      oDoc.Add(clabeTable)
      oDoc.Add(New Paragraph(" ", f7))

      Dim info As New PdfPTable(2)
      info.TotalWidth = 540.0F : info.LockedWidth = True
      info.SetWidths(New Single() {270.0F, 270.0F})
      Dim emisor As New PdfPCell()
      emisor.BorderColor = grisBorde : emisor.Padding = 9.0F
      emisor.AddElement(New Phrase("EMISOR", f8b))
      emisor.AddElement(New Phrase("Comunícalo de México S.A. de C.V.", f9b))
      emisor.AddElement(New Phrase("Convento de Churubusco No. 4," & Environment.NewLine & "Col. Jardines de Santa Mónica" & Environment.NewLine & "Mpio. Tlalnepantla de Baz, Estado de México, C.P. 54050" & Environment.NewLine & "RFC: CME0806162SA", f8))
      info.AddCell(emisor)
      Dim cliente As New PdfPCell()
      cliente.BorderColor = grisBorde : cliente.Padding = 9.0F
      cliente.AddElement(New Phrase("CLIENTE", f8b))
      cliente.AddElement(New Phrase(nombreCliente, f9b))
      cliente.AddElement(New Phrase(direccionCliente1 & Environment.NewLine & direccionCliente2 & Environment.NewLine & Environment.NewLine & "Contrato   " & contrato & Environment.NewLine & "Teléfono   " & telefono, f8))
      info.AddCell(cliente)
      oDoc.Add(info)
      oDoc.Add(New Paragraph(" ", f7))

      Dim sec As New PdfPTable(1)
      sec.TotalWidth = 540.0F : sec.LockedWidth = True
      Dim secCell As New PdfPCell(New Phrase("SERVICIOS CONTRATADOS · DESGLOSE DE CARGOS", f10b))
      secCell.Border = PdfPCell.BOTTOM_BORDER : secCell.BorderColorBottom = grisBorde : secCell.PaddingBottom = 5.0F
      sec.AddCell(secCell)
      oDoc.Add(sec)

      Dim charges As New PdfPTable(3)
      charges.TotalWidth = 540.0F : charges.LockedWidth = True
      charges.SetWidths(New Single() {130.0F, 325.0F, 85.0F})
      Dim hp As New PdfPCell(New Phrase("PLAN", f7b)) : hp.BackgroundColor = grisClaro : hp.BorderColor = grisBorde : hp.Padding = 6.0F : charges.AddCell(hp)
      Dim hc As New PdfPCell(New Phrase("CONCEPTO", f7b)) : hc.BackgroundColor = grisClaro : hc.BorderColor = grisBorde : hc.Padding = 6.0F : charges.AddCell(hc)
      Dim hi As New PdfPCell(New Phrase("IMPORTE", f7b)) : hi.BackgroundColor = grisClaro : hi.BorderColor = grisBorde : hi.Padding = 6.0F : hi.HorizontalAlignment = Element.ALIGN_RIGHT : charges.AddCell(hi)

      Dim mh As New PdfPCell(New Phrase("CARGOS DEL MES · " & mesFacturacion, f8b))
      mh.Colspan = 3 : mh.BackgroundColor = azulClaro : mh.BorderColor = grisBorde : mh.Padding = 6.0F : charges.AddCell(mh)
      Dim p1 As New PdfPCell(New Phrase(planActual, f8k)) : p1.BorderColor = grisBorde : p1.Padding = 7.0F : charges.AddCell(p1)
      Dim c1 As New PdfPCell(New Phrase(conceptoActual, f8k)) : c1.BorderColor = grisBorde : c1.Padding = 7.0F : charges.AddCell(c1)
      Dim i1 As New PdfPCell(New Phrase(FormatCurrency(importeActual, 2), f8k)) : i1.BorderColor = grisBorde : i1.Padding = 7.0F : i1.HorizontalAlignment = Element.ALIGN_RIGHT : charges.AddCell(i1)
      Dim s1 As New PdfPCell(New Phrase("Subtotal cargos del mes", f7)) : s1.Colspan = 2 : s1.BackgroundColor = grisClaro : s1.BorderColor = grisBorde : s1.Padding = 5.0F : charges.AddCell(s1)
      Dim s1a As New PdfPCell(New Phrase(FormatCurrency(importeActual, 2), f7)) : s1a.BackgroundColor = grisClaro : s1a.BorderColor = grisBorde : s1a.Padding = 5.0F : s1a.HorizontalAlignment = Element.ALIGN_RIGHT : charges.AddCell(s1a)

      Dim prev As New PdfPCell(New Phrase("CARGOS DEL MES ANTERIOR · " & mesAnterior & " (SALDO VENCIDO)", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.BOLD, New Color(125, 92, 18))))
      prev.Colspan = 3 : prev.BackgroundColor = amarilloClaro : prev.BorderColor = grisBorde : prev.Padding = 6.0F : charges.AddCell(prev)
      Dim p2 As New PdfPCell(New Phrase(planActual, f8k)) : p2.BorderColor = grisBorde : p2.Padding = 7.0F : charges.AddCell(p2)
      Dim c2 As New PdfPCell(New Phrase(conceptoMesAnterior, f8k)) : c2.BorderColor = grisBorde : c2.Padding = 7.0F : charges.AddCell(c2)
      Dim i2 As New PdfPCell(New Phrase(FormatCurrency(importeMesAnterior, 2), f8k)) : i2.BorderColor = grisBorde : i2.Padding = 7.0F : i2.HorizontalAlignment = Element.ALIGN_RIGHT : charges.AddCell(i2)
      Dim pc As New PdfPCell(New Phrase("COMISIÓN", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.BOLD, New Color(125, 92, 18)))) : pc.BorderColor = grisBorde : pc.Padding = 7.0F : charges.AddCell(pc)
      Dim cc As New PdfPCell(New Phrase("Cargo administrativo por pago tardío", f8k)) : cc.BorderColor = grisBorde : cc.Padding = 7.0F : charges.AddCell(cc)
      Dim ci As New PdfPCell(New Phrase(FormatCurrency(cargoPagoTardio, 2), f8k)) : ci.BorderColor = grisBorde : ci.Padding = 7.0F : ci.HorizontalAlignment = Element.ALIGN_RIGHT : charges.AddCell(ci)
      Dim sp As New PdfPCell(New Phrase("Subtotal mes anterior", f7)) : sp.Colspan = 2 : sp.BackgroundColor = grisClaro : sp.BorderColor = grisBorde : sp.Padding = 5.0F : charges.AddCell(sp)
      Dim spa As New PdfPCell(New Phrase(FormatCurrency(saldoVencido, 2), f7)) : spa.BackgroundColor = grisClaro : spa.BorderColor = grisBorde : spa.Padding = 5.0F : spa.HorizontalAlignment = Element.ALIGN_RIGHT : charges.AddCell(spa)

      Dim totalLbl As New PdfPCell()
      totalLbl.Colspan = 2 : totalLbl.BackgroundColor = azul : totalLbl.Border = PdfPCell.NO_BORDER : totalLbl.Padding = 8.0F
      totalLbl.AddElement(New Phrase("TOTAL A PAGAR", f10bw))
      totalLbl.AddElement(New Phrase("Acumulado: cargos del mes + saldo anterior", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 6.8F, iTextSharp.text.Font.NORMAL, Color.WHITE)))
      charges.AddCell(totalLbl)
      Dim totalAmt As New PdfPCell(New Phrase(FormatCurrency(totalPagar, 2), f12bw))
      totalAmt.BackgroundColor = azul : totalAmt.Border = PdfPCell.NO_BORDER : totalAmt.Padding = 9.0F : totalAmt.HorizontalAlignment = Element.ALIGN_RIGHT
      charges.AddCell(totalAmt)
      oDoc.Add(charges)

      Dim terms As New PdfPTable(1)
      terms.TotalWidth = 540.0F : terms.LockedWidth = True
      Dim termsCell As New PdfPCell(New Phrase("* En el caso de haber realizado un cambio o actualización en su paquete, al realizar el pago de este Estado de Cuenta, usted acepta los nuevos Términos y Condiciones aplicables.", f7))
      termsCell.BackgroundColor = grisClaro : termsCell.Border = PdfPCell.NO_BORDER : termsCell.Padding = 8.0F
      terms.AddCell(termsCell)
      oDoc.Add(terms)

      oDoc.Add(New Paragraph(" ", f7))
      Dim footer1 As New PdfPTable(2)
      footer1.TotalWidth = 540.0F : footer1.LockedWidth = True
      footer1.SetWidths(New Single() {390.0F, 150.0F})
      Dim fc1 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx  ·  Atención a clientes: 55 2601 4010" & Environment.NewLine & "Horario de atención de 9 a 18 hrs", f7))
      fc1.Border = PdfPCell.TOP_BORDER : fc1.BorderColorTop = grisBorde : fc1.PaddingTop = 8.0F : footer1.AddCell(fc1)
      Dim fl1 As New PdfPCell()
      fl1.Border = PdfPCell.TOP_BORDER : fl1.BorderColorTop = grisBorde : fl1.PaddingTop = 6.0F
      Dim logoF1 As Image = Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
      logoF1.ScaleToFit(95.0F, 35.0F) : fl1.AddElement(logoF1) : footer1.AddCell(fl1)
      oDoc.Add(footer1)
      Dim pg1 As New Paragraph("PÁGINA 1 DE 2", f7) : pg1.Alignment = Element.ALIGN_CENTER : oDoc.Add(pg1)

      '======================== PÁGINA 2 ========================
      oDoc.NewPage()
      Dim header2 As New PdfPTable(2)
      header2.TotalWidth = 540.0F : header2.LockedWidth = True
      header2.SetWidths(New Single() {220.0F, 320.0F})
      Dim logo2 As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
      logo2.ScaleToFit(130.0F, 50.0F)
      Dim cl2 As New PdfPCell(logo2) : cl2.Border = PdfPCell.NO_BORDER : header2.AddCell(cl2)
      Dim h2r As New PdfPTable(1)
      Dim h2t As New PdfPCell(New Phrase("FORMAS DE PAGO", f13b)) : h2t.Border = PdfPCell.NO_BORDER : h2t.HorizontalAlignment = Element.ALIGN_RIGHT : h2r.AddCell(h2t)
      Dim h2s As New PdfPCell(New Phrase("Contrato: " & contrato & " · " & mesFacturacion, f8)) : h2s.Border = PdfPCell.NO_BORDER : h2s.HorizontalAlignment = Element.ALIGN_RIGHT : h2r.AddCell(h2s)
      Dim ch2 As New PdfPCell(h2r) : ch2.Border = PdfPCell.NO_BORDER : header2.AddCell(ch2)
      Dim hl2 As New PdfPCell(New Phrase("")) : hl2.Colspan = 2 : hl2.Border = PdfPCell.BOTTOM_BORDER : hl2.BorderColorBottom = azulOscuro : hl2.BorderWidthBottom = 2.2F : header2.AddCell(hl2)
      oDoc.Add(header2)
      oDoc.Add(New Paragraph(" ", f7))

      Dim lateTitle As New PdfPTable(1)
      lateTitle.TotalWidth = 540.0F : lateTitle.LockedWidth = True
      Dim ltc As New PdfPCell(New Phrase("IMPORTANTE — PAGO TARDÍO", f10bw))
      ltc.BackgroundColor = azul : ltc.Border = PdfPCell.NO_BORDER : ltc.Padding = 9.0F : lateTitle.AddCell(ltc)
      oDoc.Add(lateTitle)

      Dim lateBody As New PdfPTable(2)
      lateBody.TotalWidth = 540.0F : lateBody.LockedWidth = True
      lateBody.SetWidths(New Single() {435.0F, 105.0F})
      Dim lbt As New PdfPCell(New Phrase("A partir del mes de AGOSTO DE 2026, los pagos realizados después de la fecha límite establecida generarán un cargo administrativo por pago tardío de " & FormatCurrency(cargoPagoTardio, 2) & " (Cincuenta pesos 00/100 M.N.), el cual se aplicará de forma inmediata." & Environment.NewLine & Environment.NewLine & "En caso de suspensión del servicio por falta de pago, este cargo deberá liquidarse junto con la mensualidad vencida para la reactivación del servicio. Le invitamos a pagar dentro de la fecha establecida para evitar cargos adicionales.", f8k))
      lbt.BackgroundColor = amarilloClaro : lbt.BorderColor = New Color(232, 220, 170) : lbt.Padding = 11.0F : lateBody.AddCell(lbt)
      Dim lba As New PdfPCell()
      lba.BackgroundColor = amarilloClaro : lba.BorderColor = New Color(232, 220, 170) : lba.Padding = 12.0F
      lba.AddElement(New Paragraph(FormatCurrency(cargoPagoTardio, 2), New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 18.0F, iTextSharp.text.Font.BOLD, New Color(125, 92, 18))) With {.Alignment = Element.ALIGN_CENTER})
      lba.AddElement(New Paragraph("CARGO POR" & Environment.NewLine & "PAGO TARDÍO", New iTextSharp.text.Font(iTextSharp.text.Font.HELVETICA, 7.0F, iTextSharp.text.Font.BOLD, New Color(125, 92, 18))) With {.Alignment = Element.ALIGN_CENTER})
      lateBody.AddCell(lba)
      oDoc.Add(lateBody)
      oDoc.Add(New Paragraph(" ", f7))

      Dim storeBox As New PdfPTable(1)
      storeBox.TotalWidth = 540.0F : storeBox.LockedWidth = True
      Dim st As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS", f10b))
      st.BorderColor = grisBorde : st.HorizontalAlignment = Element.ALIGN_CENTER : st.Padding = 8.0F : storeBox.AddCell(st)
      Dim si As New PdfPCell()
      si.BorderColor = grisBorde : si.HorizontalAlignment = Element.ALIGN_CENTER : si.Padding = 8.0F
      Dim stores As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/stores_3.jpg")
      stores.ScaleToFit(370.0F, 90.0F) : stores.Alignment = Element.ALIGN_CENTER : si.AddElement(stores)
      storeBox.AddCell(si)
      Dim bc As New PdfPCell()
      bc.BorderColor = grisBorde : bc.HorizontalAlignment = Element.ALIGN_CENTER : bc.Padding = 10.0F
      If Not String.IsNullOrWhiteSpace(codigoBarraOxxo) Then
        ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
        Dim barcode As Image = Image.GetInstance(codigoBarraOxxo)
        barcode.ScaleToFit(185.0F, 80.0F) : barcode.Alignment = Element.ALIGN_CENTER : bc.AddElement(barcode)
      End If
      bc.AddElement(New Paragraph(refOxxo, New iTextSharp.text.Font(iTextSharp.text.Font.COURIER, 12.0F, iTextSharp.text.Font.BOLD, Color.BLACK)) With {.Alignment = Element.ALIGN_CENTER})
      bc.AddElement(New Paragraph("Muestre este código de barras en caja o dicte los dígitos de la referencia.", f7) With {.Alignment = Element.ALIGN_CENTER})
      storeBox.AddCell(bc)
      oDoc.Add(storeBox)
      oDoc.Add(New Paragraph(" ", f7))

      Dim instTitle As New Paragraph("INSTRUCCIONES PARA PAGO EN TIENDAS", f10b)
      oDoc.Add(instTitle)
      oDoc.Add(New Paragraph(" ", f7))
      Dim steps As New PdfPTable(3)
      steps.TotalWidth = 540.0F : steps.LockedWidth = True
      steps.SetWidths(New Single() {180.0F, 180.0F, 180.0F})
      Dim sA As New PdfPCell(New Phrase("1" & Environment.NewLine & Environment.NewLine & "Elija la tienda que más le convenga entre las cadenas indicadas (solo se puede pagar en esas tiendas).", f8k))
      sA.BorderColor = grisBorde : sA.Padding = 10.0F : sA.MinimumHeight = 92.0F : steps.AddCell(sA)
      Dim sB As New PdfPCell(New Phrase("2" & Environment.NewLine & Environment.NewLine & "Al acercarse al mostrador, mencione que viene a pagar CONEKTA y muestre el código de barras o dicte los números de la referencia.", f8k))
      sB.BorderColor = grisBorde : sB.Padding = 10.0F : sB.MinimumHeight = 92.0F : steps.AddCell(sB)
      Dim sC As New PdfPCell(New Phrase("3" & Environment.NewLine & Environment.NewLine & "Una vez realizado el pago en efectivo, enviaremos una notificación de pago en tiempo real a su correo y ¡listo!", f8k))
      sC.BorderColor = grisBorde : sC.Padding = 10.0F : sC.MinimumHeight = 92.0F : steps.AddCell(sC)
      oDoc.Add(steps)

      oDoc.Add(New Paragraph(Environment.NewLine & Environment.NewLine & Environment.NewLine, f7))
      Dim footer2 As New PdfPTable(2)
      footer2.TotalWidth = 540.0F : footer2.LockedWidth = True
      footer2.SetWidths(New Single() {390.0F, 150.0F})
      Dim fc2 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx  ·  Atención a clientes: 55 2601 4010" & Environment.NewLine & "Horario de atención de 9 a 18 hrs", f7))
      fc2.Border = PdfPCell.TOP_BORDER : fc2.BorderColorTop = grisBorde : fc2.PaddingTop = 8.0F : footer2.AddCell(fc2)
      Dim fl2 As New PdfPCell()
      fl2.Border = PdfPCell.TOP_BORDER : fl2.BorderColorTop = grisBorde : fl2.PaddingTop = 6.0F
      Dim logoF2 As Image = Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
      logoF2.ScaleToFit(95.0F, 35.0F) : fl2.AddElement(logoF2) : footer2.AddCell(fl2)
      oDoc.Add(footer2)
      Dim pg2 As New Paragraph("PÁGINA 2 DE 2", f7) : pg2.Alignment = Element.ALIGN_CENTER : oDoc.Add(pg2)

      If oDoc IsNot Nothing AndAlso oDoc.IsOpen Then
        oDoc.Close()
      End If

    Catch ex As Exception

      Try
        If oDoc IsNot Nothing AndAlso oDoc.IsOpen Then
          oDoc.Close()
        End If
      Catch
      End Try

      If File.Exists(ruta) Then
        Try
          File.Delete(ruta)
        Catch
        End Try
      End If

      MsgBox(ex.Message & Environment.NewLine & ex.StackTrace)

    Finally

      pdfw = Nothing
      oDoc = Nothing

    End Try
  End Sub

  Private Function ObtenerEstadoCuenta(ByVal id_estado_cuenta As Integer) As DataTable
    Dim sql As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtEdo As DataTable = con.ConsultarDT(sql)

    Return dtEdo
  End Function

  Private Sub Generar_pdfOXXO_Rediseno(ByVal id_estado_cuenta As Integer,
                                           ByVal id_contrato As Integer,
                                           ByVal path As String,
                                           ByVal refOxxo As String,
                                           ByVal codigoBarraOxxo As String)

    '==========================================================================
    ' DATOS DINÁMICOS
    ' Se conservan las mismas consultas/orígenes usados por Generar_pdfOXXO.
    '==========================================================================
    Dim sqledo As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtedo As DataTable = con.ConsultarDT(sqledo)

    If dtedo Is Nothing OrElse dtedo.Rows.Count = 0 Then
      Exit Sub
    End If

    Dim fecha As Date = dtedo(0)("fecha").ToString
    Dim granTotal As Double = Val(dtedo(0)("grantotal").ToString)
    Dim saldoPendiente As Double = Val(dtedo(0)("saldo_pendiente").ToString)
    Dim periodoA As Date = dtedo(0)("periodoA").ToString
    Dim periodoB As Date = dtedo(0)("periodoB").ToString
    Dim totalPlan As Double = Val(dtedo(0)("mensualidad").ToString())

    Dim sqlBalance As String = "select * from CONTRACTS_BALANCES where id_contrato=" & id_contrato & ";"
    Dim dtBalance As DataTable = con.ConsultarDT(sqlBalance)
    Dim balance As Double = 0

    If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
      balance = Val(dtBalance(0)("balance").ToString)
    End If

    Dim sqlcli As String = ""

    If tiene_telefonia(id_contrato) Then
      sqlcli = $"SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero,t3.id_contrato,id_paquete FROM (" &
      " SELECT t2.*,upper(p.nombre) AS paquete FROM (" &
      " SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete FROM(" &
      " SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre) AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias" &
      " FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e" &
      " ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato &
      " ) AS t1 INNER JOIN CALLES ca ON ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete)" &
      " AS t3 INNER JOIN dbo.EQUIPOS e INNER JOIN EQUIPOS_TELEFONIA et INNER JOIN LINEAS l" &
      " ON l.id_linea=et.id_linea ON et.id_equipo=e.id_equipo ON e.id_contrato=t3.id_contrato where e.estatus=1 AND et.estatus=1"
    Else
      sqlcli = "SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero," &
      " t3.id_contrato,id_paquete FROM (" &
      " SELECT t2.*,upper(p.nombre) AS paquete FROM" &
      " ( SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete,numero" &
      " FROM( SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre)" &
      " AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias,telefono AS numero FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c " &
      " INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON" &
      " col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato & " ) AS t1 INNER JOIN CALLES ca ON" &
      " ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete) AS t3"
    End If

    Dim dtcli As DataTable = con.ConsultarDT(sqlcli)
    If dtcli Is Nothing OrElse dtcli.Rows.Count = 0 Then
      Exit Sub
    End If

    Dim nombreCliente As String = dtcli(0)("nombre").ToString
    Dim contrato As String = dtcli(0)("contrato").ToString
    Dim telefono As String = dtcli(0)("numero").ToString
    Dim calle As String = dtcli(0)("calle").ToString
    Dim numext As String = dtcli(0)("numext").ToString
    Dim numint As String = dtcli(0)("numint").ToString
    Dim colonia As String = dtcli(0)("colonia").ToString
    Dim cp As String = dtcli(0)("cp").ToString
    Dim municipio As String = dtcli(0)("municipio").ToString
    Dim estadoCliente As String = dtcli(0)("estado").ToString
    Dim referencias As String = dtcli(0)("referencias").ToString
    Dim plan As String = dtcli(0)("paquete").ToString
    Dim id_paquete As Integer = Val(dtcli(0)("id_paquete").ToString)
    Dim servicios As String = getServicios(id_paquete)

    Dim mesFacturacion As String = MonthName(periodoA.Month).ToUpper()
    Dim fechaLimite As Date = periodoA
    Dim cargoMesActual As Decimal = CDec(totalPlan)
    Dim saldoVencido As Decimal = CDec(saldoPendiente)
    Dim totalPagar As Decimal = CDec(granTotal)

    ' La versión original no consulta el recargo desde BD; usa $50 en el aviso.
    Dim cargoPagoTardio As Decimal = 50D
    Dim totalDespuesFecha As Decimal = totalPagar + cargoPagoTardio

    Dim banco As String = "STP"
    Dim beneficiario As String = "Comunícalo de México, S.A. de C.V."

    Dim direccionCliente1 As String = calle
    If Not String.IsNullOrWhiteSpace(referencias) Then
      direccionCliente1 &= " " & referencias
    End If
    If Not String.IsNullOrWhiteSpace(numext) Then
      direccionCliente1 &= " " & numext
    End If
    If Not String.IsNullOrWhiteSpace(numint) Then
      direccionCliente1 &= " " & numint
    End If
    If Not String.IsNullOrWhiteSpace(colonia) Then
      direccionCliente1 &= ", " & colonia
    End If

    Dim direccionCliente2 As String = municipio & ", " & estadoCliente & ", C.P. " & cp
    Dim conceptoActual As String = servicios

    Dim dtCharges As DataTable = getDataBillCharges(id_estado_cuenta)
    Dim dtDiscount As DataTable = getDataBillDiscounts(id_estado_cuenta)

    Dim ruta As String = System.IO.Path.Combine(
      path,
      "EstadoCuentaRediseno(" & id_estado_cuenta.ToString() & ").pdf"
    )

    '==========================================================================
    ' COLORES
    '==========================================================================
    Dim azul As New iTextSharp.text.Color(22, 67, 126)
    Dim azulOscuro As New iTextSharp.text.Color(12, 45, 88)
    Dim azulClaro As New iTextSharp.text.Color(239, 244, 250)
    Dim grisClaro As New iTextSharp.text.Color(246, 247, 249)
    Dim grisBorde As New iTextSharp.text.Color(218, 222, 228)
    Dim grisTexto As New iTextSharp.text.Color(88, 94, 103)
    Dim amarilloClaro As New iTextSharp.text.Color(252, 247, 225)
    Dim amarillo As New iTextSharp.text.Color(177, 128, 27)

    '==========================================================================
    ' FUENTES
    '==========================================================================
    Dim f7 As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 7.0F,
    iTextSharp.text.Font.NORMAL, grisTexto)

    Dim f7Bold As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 7.0F,
    iTextSharp.text.Font.BOLD, iTextSharp.text.Color.BLACK)

    Dim f8 As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 8.0F,
    iTextSharp.text.Font.NORMAL, grisTexto)

    Dim f8Black As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 8.0F,
    iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.BLACK)

    Dim f8BoldBlue As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 8.0F,
    iTextSharp.text.Font.BOLD, azulOscuro)

    Dim f8BoldWhite As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 8.0F,
    iTextSharp.text.Font.BOLD, iTextSharp.text.Color.WHITE)

    Dim f9Bold As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 9.0F,
    iTextSharp.text.Font.BOLD, iTextSharp.text.Color.BLACK)

    Dim f10BoldBlue As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 10.0F,
    iTextSharp.text.Font.BOLD, azulOscuro)

    Dim f10BoldWhite As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 10.0F,
    iTextSharp.text.Font.BOLD, iTextSharp.text.Color.WHITE)

    Dim f13BoldBlue As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 13.0F,
    iTextSharp.text.Font.BOLD, azulOscuro)

    Dim f16BoldWhite As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 16.0F,
    iTextSharp.text.Font.BOLD, iTextSharp.text.Color.WHITE)

    Dim f18BoldBlue As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 18.0F,
    iTextSharp.text.Font.BOLD, azulOscuro)

    Dim f20BoldWhite As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA, 20.0F,
    iTextSharp.text.Font.BOLD, iTextSharp.text.Color.WHITE)

    Dim documento As New iTextSharp.text.Document(
    iTextSharp.text.PageSize.LETTER, 36, 36, 28, 28)

    Dim writer As iTextSharp.text.pdf.PdfWriter = Nothing

    Try
      writer = iTextSharp.text.pdf.PdfWriter.GetInstance(
      documento,
      New System.IO.FileStream(
        ruta,
        System.IO.FileMode.Create,
        System.IO.FileAccess.Write,
        System.IO.FileShare.None
      )
    )

      documento.Open()

      '========================================================================
      ' PÁGINA 1 - ENCABEZADO
      '========================================================================
      Dim tblHeader As New iTextSharp.text.pdf.PdfPTable(2)
      tblHeader.TotalWidth = 540.0F
      tblHeader.LockedWidth = True
      tblHeader.SetWidths(New Single() {220.0F, 320.0F})

      Dim logo As iTextSharp.text.Image =
      iTextSharp.text.Image.GetInstance(
        Application.StartupPath & "/imgs/LOGOCOMUNICALO.png"
      )
      logo.ScaleToFit(165.0F, 65.0F)

      Dim cellLogo As New iTextSharp.text.pdf.PdfPCell(logo)
      cellLogo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellLogo.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      cellLogo.PaddingBottom = 8.0F
      tblHeader.AddCell(cellLogo)

      Dim tblTitulo As New iTextSharp.text.pdf.PdfPTable(1)
      tblTitulo.WidthPercentage = 100

      Dim cellTitulo As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("ESTADO DE CUENTA", f18BoldBlue)
    )
      cellTitulo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTitulo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      tblTitulo.AddCell(cellTitulo)

      Dim subtitulo As New iTextSharp.text.Phrase()
      subtitulo.Add(New iTextSharp.text.Chunk("Mes de facturación: ", f8))
      subtitulo.Add(New iTextSharp.text.Chunk(mesFacturacion, f8BoldBlue))
      subtitulo.Add(New iTextSharp.text.Chunk("  ·  Contrato: ", f8))
      subtitulo.Add(New iTextSharp.text.Chunk(contrato, f8BoldBlue))

      Dim cellSubtitulo As New iTextSharp.text.pdf.PdfPCell(subtitulo)
      cellSubtitulo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellSubtitulo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      cellSubtitulo.PaddingTop = 3.0F
      tblTitulo.AddCell(cellSubtitulo)

      Dim cellTituloContenedor As New iTextSharp.text.pdf.PdfPCell(tblTitulo)
      cellTituloContenedor.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTituloContenedor.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      cellTituloContenedor.PaddingBottom = 8.0F
      tblHeader.AddCell(cellTituloContenedor)

      Dim cellLineaHeader As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("")
    )
      cellLineaHeader.Colspan = 2
      cellLineaHeader.Border = iTextSharp.text.pdf.PdfPCell.BOTTOM_BORDER
      cellLineaHeader.BorderColorBottom = azulOscuro
      cellLineaHeader.BorderWidthBottom = 2.2F
      cellLineaHeader.FixedHeight = 4.0F
      tblHeader.AddCell(cellLineaHeader)

      documento.Add(tblHeader)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      '========================================================================
      '========================================================================
      ' RESUMEN PRINCIPAL
      ' Se usan tablas anidadas con una fila por texto para evitar duplicados.
      '========================================================================
      Dim tblResumen As New iTextSharp.text.pdf.PdfPTable(3)
      tblResumen.TotalWidth = 540.0F
      tblResumen.LockedWidth = True
      tblResumen.SetWidths(New Single() {235.0F, 195.0F, 110.0F})

      ' TOTAL A PAGAR
      Dim tblTotal As New iTextSharp.text.pdf.PdfPTable(1)
      tblTotal.WidthPercentage = 100

      Dim totalTitulo As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("TOTAL A PAGAR", f8BoldWhite))
      totalTitulo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      totalTitulo.BackgroundColor = azul
      totalTitulo.PaddingLeft = 10.0F
      totalTitulo.PaddingTop = 8.0F
      totalTitulo.PaddingBottom = 0.0F
      tblTotal.AddCell(totalTitulo)

      Dim totalImporte As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(FormatCurrency(granTotal, 2), f20BoldWhite))
      totalImporte.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      totalImporte.BackgroundColor = azul
      totalImporte.PaddingLeft = 10.0F
      totalImporte.PaddingTop = 0.0F
      totalImporte.PaddingBottom = 0.0F
      tblTotal.AddCell(totalImporte)

      Dim totalNota As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "Incluye cargos del mes y saldo anterior acumulado",
        New iTextSharp.text.Font(
          iTextSharp.text.Font.HELVETICA, 7.0F,
          iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.WHITE)))
      totalNota.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      totalNota.BackgroundColor = azul
      totalNota.PaddingLeft = 10.0F
      totalNota.PaddingTop = 1.0F
      totalNota.PaddingBottom = 8.0F
      tblTotal.AddCell(totalNota)

      Dim contTotal As New iTextSharp.text.pdf.PdfPCell(tblTotal)
      contTotal.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      contTotal.BackgroundColor = azul
      contTotal.Padding = 0.0F
      tblResumen.AddCell(contTotal)

      ' FECHA LÍMITE
      Dim tblFecha As New iTextSharp.text.pdf.PdfPTable(1)
      tblFecha.WidthPercentage = 100

      Dim fechaTitulo As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "FECHA LÍMITE DE PAGO",
        New iTextSharp.text.Font(
          iTextSharp.text.Font.HELVETICA, 8.0F,
          iTextSharp.text.Font.BOLD, amarillo)))
      fechaTitulo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      fechaTitulo.BackgroundColor = grisClaro
      fechaTitulo.PaddingLeft = 12.0F
      fechaTitulo.PaddingTop = 8.0F
      fechaTitulo.PaddingBottom = 0.0F
      tblFecha.AddCell(fechaTitulo)

      Dim fechaValor As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        periodoA.ToString("dd / MMM / yyyy").ToUpper(),
        f13BoldBlue))
      fechaValor.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      fechaValor.BackgroundColor = grisClaro
      fechaValor.PaddingLeft = 12.0F
      fechaValor.PaddingTop = 0.0F
      fechaValor.PaddingBottom = 0.0F
      tblFecha.AddCell(fechaValor)

      Dim fechaNota As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("Pague a tiempo y evite recargos", f7))
      fechaNota.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      fechaNota.BackgroundColor = grisClaro
      fechaNota.PaddingLeft = 12.0F
      fechaNota.PaddingTop = 1.0F
      fechaNota.PaddingBottom = 8.0F
      tblFecha.AddCell(fechaNota)

      Dim contFecha As New iTextSharp.text.pdf.PdfPCell(tblFecha)
      contFecha.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      contFecha.BackgroundColor = grisClaro
      contFecha.Padding = 0.0F
      tblResumen.AddCell(contFecha)

      ' SALDO VENCIDO
      Dim tblSaldo As New iTextSharp.text.pdf.PdfPTable(1)
      tblSaldo.WidthPercentage = 100

      Dim saldoTitulo As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("SALDO VENCIDO", f8BoldWhite))
      saldoTitulo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      saldoTitulo.BackgroundColor = azulOscuro
      saldoTitulo.PaddingLeft = 10.0F
      saldoTitulo.PaddingTop = 8.0F
      saldoTitulo.PaddingBottom = 0.0F
      tblSaldo.AddCell(saldoTitulo)

      Dim saldoImporte As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        FormatCurrency(saldoVencido, 2),
        New iTextSharp.text.Font(
          iTextSharp.text.Font.HELVETICA, 16.0F,
          iTextSharp.text.Font.BOLD,
          New iTextSharp.text.Color(255, 213, 90))))
      saldoImporte.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      saldoImporte.BackgroundColor = azulOscuro
      saldoImporte.PaddingLeft = 10.0F
      saldoImporte.PaddingTop = 0.0F
      saldoImporte.PaddingBottom = 1.0F
      tblSaldo.AddCell(saldoImporte)

      Dim saldoAviso As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "REQUIERE PAGO" & Environment.NewLine & "INMEDIATO",
        New iTextSharp.text.Font(
          iTextSharp.text.Font.HELVETICA, 7.0F,
          iTextSharp.text.Font.BOLD,
          New iTextSharp.text.Color(255, 218, 95))))
      saldoAviso.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      saldoAviso.BackgroundColor = azulOscuro
      saldoAviso.PaddingLeft = 10.0F
      saldoAviso.PaddingTop = 1.0F
      saldoAviso.PaddingBottom = 8.0F
      tblSaldo.AddCell(saldoAviso)

      Dim contSaldo As New iTextSharp.text.pdf.PdfPCell(tblSaldo)
      contSaldo.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      contSaldo.BackgroundColor = azulOscuro
      contSaldo.Padding = 0.0F
      tblResumen.AddCell(contSaldo)

      documento.Add(tblResumen)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      '========================================================================
      ' AVISO DE RECARGO
      '========================================================================
      Dim tblAviso As New iTextSharp.text.pdf.PdfPTable(2)
      tblAviso.TotalWidth = 540.0F
      tblAviso.LockedWidth = True
      tblAviso.SetWidths(New Single() {430.0F, 110.0F})

      Dim tblAvisoIzq As New iTextSharp.text.pdf.PdfPTable(2)
      tblAvisoIzq.WidthPercentage = 100
      tblAvisoIzq.SetWidths(New Single() {38.0F, 392.0F})

      Dim iconoAviso As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "!",
        New iTextSharp.text.Font(
          iTextSharp.text.Font.HELVETICA, 16.0F,
          iTextSharp.text.Font.BOLD, iTextSharp.text.Color.WHITE)))
      iconoAviso.BackgroundColor = amarillo
      iconoAviso.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      iconoAviso.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
      iconoAviso.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      iconoAviso.PaddingTop = 10.0F
      iconoAviso.PaddingBottom = 10.0F
      tblAvisoIzq.AddCell(iconoAviso)

      Dim textoAviso As New iTextSharp.text.Phrase()
      textoAviso.Add(New iTextSharp.text.Chunk(
      "Si no paga antes del " & fechaLimite.ToString("dd/MM/yyyy") & ", ",
      New iTextSharp.text.Font(
        iTextSharp.text.Font.HELVETICA, 8.0F,
        iTextSharp.text.Font.BOLD,
        New iTextSharp.text.Color(125, 92, 18))))
      textoAviso.Add(New iTextSharp.text.Chunk(
      "se sumará un cargo por pago tardío de ",
      f8Black))
      textoAviso.Add(New iTextSharp.text.Chunk(
      FormatCurrency(cargoPagoTardio, 2),
      New iTextSharp.text.Font(
        iTextSharp.text.Font.HELVETICA, 8.0F,
        iTextSharp.text.Font.BOLD,
        New iTextSharp.text.Color(125, 92, 18))))
      textoAviso.Add(New iTextSharp.text.Chunk(
      " y su servicio podrá ser suspendido.",
      f8Black))

      Dim avisoTextoCell As New iTextSharp.text.pdf.PdfPCell(textoAviso)
      avisoTextoCell.BackgroundColor = amarilloClaro
      avisoTextoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      avisoTextoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      avisoTextoCell.PaddingLeft = 12.0F
      avisoTextoCell.PaddingRight = 8.0F
      avisoTextoCell.PaddingTop = 10.0F
      avisoTextoCell.PaddingBottom = 10.0F
      tblAvisoIzq.AddCell(avisoTextoCell)

      Dim contAvisoIzq As New iTextSharp.text.pdf.PdfPCell(tblAvisoIzq)
      contAvisoIzq.BackgroundColor = amarilloClaro
      contAvisoIzq.BorderColor = New iTextSharp.text.Color(232, 220, 170)
      contAvisoIzq.BorderWidth = 0.7F
      contAvisoIzq.Padding = 0.0F
      tblAviso.AddCell(contAvisoIzq)

      Dim tblTotalDespues As New iTextSharp.text.pdf.PdfPTable(1)
      tblTotalDespues.WidthPercentage = 100

      Dim despuesImporte As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        FormatCurrency(totalDespuesFecha, 2),
        f16BoldWhite))
      despuesImporte.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      despuesImporte.BackgroundColor = azul
      despuesImporte.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
      despuesImporte.PaddingTop = 8.0F
      despuesImporte.PaddingBottom = 0.0F
      tblTotalDespues.AddCell(despuesImporte)

      Dim despuesLeyenda As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "TOTAL DESPUÉS" & Environment.NewLine &
        "DEL " & fechaLimite.ToString("dd/MM/yyyy"),
        New iTextSharp.text.Font(
          iTextSharp.text.Font.HELVETICA, 6.8F,
          iTextSharp.text.Font.BOLD,
          iTextSharp.text.Color.WHITE)))
      despuesLeyenda.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      despuesLeyenda.BackgroundColor = azul
      despuesLeyenda.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
      despuesLeyenda.PaddingTop = 0.0F
      despuesLeyenda.PaddingBottom = 8.0F
      tblTotalDespues.AddCell(despuesLeyenda)

      Dim contTotalDespues As New iTextSharp.text.pdf.PdfPCell(tblTotalDespues)
      contTotalDespues.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      contTotalDespues.BackgroundColor = azul
      contTotalDespues.Padding = 0.0F
      tblAviso.AddCell(contTotalDespues)

      documento.Add(tblAviso)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      '========================================================================
      ' CLABE PERSONALIZADA
      '========================================================================
      Dim tblClabe As New iTextSharp.text.pdf.PdfPTable(3)
      tblClabe.TotalWidth = 540.0F
      tblClabe.LockedWidth = True
      tblClabe.SetWidths(New Single() {105.0F, 300.0F, 135.0F})

      Dim cellClabeTitulo As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "CLABE" & Environment.NewLine &
        "INTERBANCARIA" & Environment.NewLine &
        "PERSONALIZADA",
        f8BoldWhite))
      cellClabeTitulo.BackgroundColor = azul
      cellClabeTitulo.BorderColor = azulOscuro
      cellClabeTitulo.BorderWidth = 0.8F
      cellClabeTitulo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
      cellClabeTitulo.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      cellClabeTitulo.Padding = 9.0F
      tblClabe.AddCell(cellClabeTitulo)

      Dim tblClabeCentro As New iTextSharp.text.pdf.PdfPTable(1)
      tblClabeCentro.WidthPercentage = 100
      Dim dtAccount As DataTable = getStpAccount(id_contrato)
      Dim account As String = "s/a"

      If dtAccount IsNot Nothing AndAlso dtAccount.Rows.Count > 0 Then
        account = dtAccount.Rows(0)("clave").ToString().Trim()
      End If

      Dim clabeFormateada As String = account
      If account.Length = 18 Then
        clabeFormateada = String.Format("{0} {1} {2} {3} {4}",
          account.Substring(0, 4),
          account.Substring(4, 4),
          account.Substring(8, 4),
          account.Substring(12, 2),
          account.Substring(14, 4))
      End If

      Dim clabeNumeroCell As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        clabeFormateada,
        New iTextSharp.text.Font(
          iTextSharp.text.Font.COURIER, 18.0F,
          iTextSharp.text.Font.BOLD, azulOscuro)))
      clabeNumeroCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      clabeNumeroCell.BackgroundColor = azulClaro
      clabeNumeroCell.PaddingLeft = 12.0F
      clabeNumeroCell.PaddingTop = 8.0F
      clabeNumeroCell.PaddingBottom = 0.0F
      tblClabeCentro.AddCell(clabeNumeroCell)

      Dim clabeNotaCell As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "Esta CLABE es exclusiva de su contrato: su pago se aplica automáticamente, sin referencia adicional.",
        f7))
      clabeNotaCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      clabeNotaCell.BackgroundColor = azulClaro
      clabeNotaCell.PaddingLeft = 12.0F
      clabeNotaCell.PaddingTop = 1.0F
      clabeNotaCell.PaddingBottom = 8.0F
      tblClabeCentro.AddCell(clabeNotaCell)

      Dim contClabeCentro As New iTextSharp.text.pdf.PdfPCell(tblClabeCentro)
      contClabeCentro.BackgroundColor = azulClaro
      contClabeCentro.BorderColor = azulOscuro
      contClabeCentro.BorderWidth = 0.8F
      contClabeCentro.Padding = 0.0F
      tblClabe.AddCell(contClabeCentro)

      Dim tblBanco As New iTextSharp.text.pdf.PdfPTable(1)
      tblBanco.WidthPercentage = 100

      Dim bancoCell As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("Banco: " & banco, f7Bold))
      bancoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      bancoCell.BackgroundColor = azulClaro
      bancoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      bancoCell.PaddingRight = 8.0F
      bancoCell.PaddingTop = 12.0F
      bancoCell.PaddingBottom = 1.0F
      tblBanco.AddCell(bancoCell)

      Dim beneficiarioCell As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "Beneficiario: " & beneficiario,
        f7))
      beneficiarioCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      beneficiarioCell.BackgroundColor = azulClaro
      beneficiarioCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      beneficiarioCell.PaddingRight = 8.0F
      beneficiarioCell.PaddingTop = 0.0F
      beneficiarioCell.PaddingBottom = 10.0F
      tblBanco.AddCell(beneficiarioCell)

      Dim contBanco As New iTextSharp.text.pdf.PdfPCell(tblBanco)
      contBanco.BackgroundColor = azulClaro
      contBanco.BorderColor = azulOscuro
      contBanco.BorderWidth = 0.8F
      contBanco.Padding = 0.0F
      tblClabe.AddCell(contBanco)

      documento.Add(tblClabe)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      ' EMISOR / CLIENTE
      '========================================================================
      '========================================================================
      ' EMISOR / CLIENTE
      '========================================================================
      Dim tblInformacion As New iTextSharp.text.pdf.PdfPTable(2)

      tblInformacion.TotalWidth = 540.0F
      tblInformacion.LockedWidth = True
      tblInformacion.SetWidths(New Single() {270.0F, 270.0F})


      '========================================================================
      ' EMISOR
      '========================================================================
      Dim fraseEmisor As New iTextSharp.text.Phrase()

      ' Título
      fraseEmisor.Add(
    New iTextSharp.text.Chunk(
        "EMISOR" &
        Environment.NewLine &
        Environment.NewLine,
        f8BoldBlue
    )
)

      ' Razón social
      fraseEmisor.Add(
    New iTextSharp.text.Chunk(
        "Comunícalo de México S.A. de C.V." &
        Environment.NewLine,
        f9Bold
    )
)

      ' Dirección
      fraseEmisor.Add(
    New iTextSharp.text.Chunk(
        "Convento de Churubusco No. 4," &
        Environment.NewLine &
        "Col. Jardines de Santa Mónica" &
        Environment.NewLine &
        "Mpio. Tlalnepantla de Baz, Estado de México, C.P. 54050" &
        Environment.NewLine &
        "RFC: CME0806162SA",
        f8
    )
)

      ' Convertimos todo a un solo Paragraph para controlar el interlineado.
      Dim parrafoEmisor As New iTextSharp.text.Paragraph(fraseEmisor)

      ' Espacio vertical entre líneas.
      parrafoEmisor.Leading = 11.0F
      parrafoEmisor.SpacingBefore = 0.0F
      parrafoEmisor.SpacingAfter = 0.0F


      Dim cellEmisor As New iTextSharp.text.pdf.PdfPCell(parrafoEmisor)

      cellEmisor.BorderColor = grisBorde
      cellEmisor.BorderWidth = 0.8F

      cellEmisor.PaddingLeft = 10.0F
      cellEmisor.PaddingRight = 10.0F
      cellEmisor.PaddingTop = 10.0F
      cellEmisor.PaddingBottom = 10.0F

      cellEmisor.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblInformacion.AddCell(cellEmisor)


      '========================================================================
      ' CLIENTE
      '========================================================================
      Dim fraseClienteInfo As New iTextSharp.text.Phrase()

      ' Título
      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        "CLIENTE" &
        Environment.NewLine &
        Environment.NewLine,
        f8BoldBlue
    )
)

      ' Nombre
      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        nombreCliente &
        Environment.NewLine,
        f9Bold
    )
)

      ' Dirección
      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        direccionCliente1 &
        Environment.NewLine &
        direccionCliente2 &
        Environment.NewLine &
        Environment.NewLine,
        f8
    )
)

      ' Contrato
      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        "Contrato      ",
        f8
    )
)

      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        contrato &
        Environment.NewLine,
        f8BoldBlue
    )
)

      ' Teléfono
      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        "Teléfono      ",
        f8
    )
)

      fraseClienteInfo.Add(
    New iTextSharp.text.Chunk(
        telefono,
        f8BoldBlue
    )
)


      ' Convertimos todo a un solo Paragraph.
      Dim parrafoCliente As New iTextSharp.text.Paragraph(fraseClienteInfo)

      ' Mismo interlineado que EMISOR.
      parrafoCliente.Leading = 11.0F
      parrafoCliente.SpacingBefore = 0.0F
      parrafoCliente.SpacingAfter = 0.0F


      Dim cellCliente As New iTextSharp.text.pdf.PdfPCell(parrafoCliente)

      cellCliente.BorderColor = grisBorde
      cellCliente.BorderWidth = 0.8F

      cellCliente.PaddingLeft = 10.0F
      cellCliente.PaddingRight = 10.0F
      cellCliente.PaddingTop = 10.0F
      cellCliente.PaddingBottom = 10.0F

      cellCliente.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblInformacion.AddCell(cellCliente)


      '========================================================================
      ' AGREGAR TABLA AL DOCUMENTO
      '========================================================================
      documento.Add(tblInformacion)

      documento.Add(
    New iTextSharp.text.Paragraph(
        " ",
        f7
    )
)

      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      '========================================================================
      ' DESGLOSE DE CARGOS
      '========================================================================
      Dim tblTituloCargos As New iTextSharp.text.pdf.PdfPTable(1)
      tblTituloCargos.TotalWidth = 540.0F
      tblTituloCargos.LockedWidth = True

      Dim cellTituloCargos As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "SERVICIOS CONTRATADOS · DESGLOSE DE CARGOS",
        f10BoldBlue
      )
    )
      cellTituloCargos.Border = iTextSharp.text.pdf.PdfPCell.BOTTOM_BORDER
      cellTituloCargos.BorderColorBottom = grisBorde
      cellTituloCargos.BorderWidthBottom = 0.8F
      cellTituloCargos.PaddingBottom = 5.0F
      tblTituloCargos.AddCell(cellTituloCargos)
      documento.Add(tblTituloCargos)

      Dim tblCargos As New iTextSharp.text.pdf.PdfPTable(3)
      tblCargos.TotalWidth = 540.0F
      tblCargos.LockedWidth = True
      tblCargos.SetWidths(New Single() {130.0F, 325.0F, 85.0F})

      Dim encabezados() As String = {"PLAN", "CONCEPTO", "IMPORTE"}
      For i As Integer = 0 To encabezados.Length - 1
        Dim c As New iTextSharp.text.pdf.PdfPCell(
        New iTextSharp.text.Phrase(encabezados(i), f7Bold)
      )
        c.BackgroundColor = grisClaro
        c.BorderColor = grisBorde
        c.Padding = 6.0F
        If i = 2 Then
          c.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
        End If
        tblCargos.AddCell(c)
      Next

      Dim cellMesActual As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "CARGOS DEL MES · " & mesFacturacion, f8BoldBlue)
    )
      cellMesActual.Colspan = 3
      cellMesActual.BackgroundColor = azulClaro
      cellMesActual.BorderColor = grisBorde
      cellMesActual.Padding = 6.0F
      tblCargos.AddCell(cellMesActual)

      AgregarFilaCargo(
        tblCargos, plan, conceptoActual, cargoMesActual,
        grisBorde, f8Black, f8Black
      )

      If dtCharges IsNot Nothing AndAlso dtCharges.Rows.Count > 0 Then
        For i As Integer = 0 To dtCharges.Rows.Count - 1
          AgregarFilaCargo(
            tblCargos,
            "OTRO CARGO",
            dtCharges.Rows(i)("nombre").ToString(),
            CDec(Val(dtCharges.Rows(i)("importe").ToString())),
            grisBorde,
            f8Black,
            f8Black
          )
        Next
      End If

      If dtDiscount IsNot Nothing AndAlso dtDiscount.Rows.Count > 0 Then
        For i As Integer = 0 To dtDiscount.Rows.Count - 1
          Dim importeDescuento As Decimal = CDec(Val(dtDiscount.Rows(i)("importe").ToString()))
          If importeDescuento > 0 Then
            importeDescuento *= -1D
          End If

          AgregarFilaCargo(
            tblCargos,
            "DESCUENTO",
            dtDiscount.Rows(i)("nombre").ToString(),
            importeDescuento,
            grisBorde,
            f8Black,
            f8Black
          )
        Next
      End If

      AgregarSubtotal(
        tblCargos, "Mensualidad del plan", cargoMesActual,
        grisClaro, grisBorde, f7
      )

      If saldoVencido > 0D Then
        Dim cellMesAnterior As New iTextSharp.text.pdf.PdfPCell(
          New iTextSharp.text.Phrase(
            "SALDO ANTERIOR / VENCIDO",
            New iTextSharp.text.Font(
              iTextSharp.text.Font.HELVETICA, 8.0F,
              iTextSharp.text.Font.BOLD,
              New iTextSharp.text.Color(125, 92, 18)
            )
          )
        )
        cellMesAnterior.Colspan = 3
        cellMesAnterior.BackgroundColor = amarilloClaro
        cellMesAnterior.BorderColor = grisBorde
        cellMesAnterior.Padding = 6.0F
        tblCargos.AddCell(cellMesAnterior)

        AgregarFilaCargo(
          tblCargos,
          "SALDO VENCIDO",
          "Saldo pendiente de periodos anteriores",
          saldoVencido,
          grisBorde,
          f8Black,
          New iTextSharp.text.Font(
            iTextSharp.text.Font.HELVETICA, 8.0F,
            iTextSharp.text.Font.BOLD,
            New iTextSharp.text.Color(125, 92, 18)
          )
        )
      End If

      Dim fraseTotal As New iTextSharp.text.Phrase()

      fraseTotal.Add(
    New iTextSharp.text.Chunk(
        "TOTAL A PAGAR",
        f10BoldWhite
    )
)

      fraseTotal.Add(
    New iTextSharp.text.Chunk(
        Environment.NewLine &
        "Importe total registrado en el estado de cuenta",
        New iTextSharp.text.Font(
            iTextSharp.text.Font.HELVETICA,
            6.8F,
            iTextSharp.text.Font.NORMAL,
            iTextSharp.text.Color.WHITE
        )
    )
)

      Dim cellTotalLabel As New iTextSharp.text.pdf.PdfPCell(fraseTotal)

      cellTotalLabel.Colspan = 2
      cellTotalLabel.BackgroundColor = azul
      cellTotalLabel.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTotalLabel.PaddingLeft = 8.0F
      cellTotalLabel.PaddingRight = 8.0F
      cellTotalLabel.PaddingTop = 8.0F
      cellTotalLabel.PaddingBottom = 8.0F
      cellTotalLabel.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE

      tblCargos.AddCell(cellTotalLabel)


      Dim cellTotalImporte As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
        FormatCurrency(totalPagar, 2),
        f10BoldWhite
    )
)

      cellTotalImporte.BackgroundColor = azul
      cellTotalImporte.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTotalImporte.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      cellTotalImporte.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      cellTotalImporte.Padding = 9.0F

      tblCargos.AddCell(cellTotalImporte)

      documento.Add(tblCargos)

      Dim tblTerminos As New iTextSharp.text.pdf.PdfPTable(1)
      tblTerminos.TotalWidth = 540.0F
      tblTerminos.LockedWidth = True

      Dim cellTerminos As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "* En el caso de haber realizado un cambio o actualización en su paquete, al realizar el pago de este Estado de Cuenta, usted acepta los nuevos Términos y Condiciones aplicables.",
        f7
      )
    )
      cellTerminos.BackgroundColor = grisClaro
      cellTerminos.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTerminos.Padding = 8.0F
      tblTerminos.AddCell(cellTerminos)
      documento.Add(tblTerminos)

      AgregarFooter(
      documento, logo, grisBorde, f7, "PÁGINA 1 DE 2"
    )

      '========================================================================
      ' PÁGINA 2
      '========================================================================
      documento.NewPage()

      Dim tblHeader2 As New iTextSharp.text.pdf.PdfPTable(2)
      tblHeader2.TotalWidth = 540.0F
      tblHeader2.LockedWidth = True
      tblHeader2.SetWidths(New Single() {220.0F, 320.0F})

      Dim logo2 As iTextSharp.text.Image =
      iTextSharp.text.Image.GetInstance(
        Application.StartupPath & "/imgs/LOGOCOMUNICALO.png"
      )
      logo2.ScaleToFit(130.0F, 50.0F)

      Dim cellLogo2 As New iTextSharp.text.pdf.PdfPCell(logo2)
      cellLogo2.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellLogo2.PaddingBottom = 7.0F
      tblHeader2.AddCell(cellLogo2)

      Dim tblTitulo2 As New iTextSharp.text.pdf.PdfPTable(1)

      Dim cellTitulo2 As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase("FORMAS DE PAGO", f13BoldBlue)
    )
      cellTitulo2.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTitulo2.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      tblTitulo2.AddCell(cellTitulo2)

      Dim cellSubtitulo2 As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "Contrato: " & contrato & " · " & mesFacturacion, f8)
    )
      cellSubtitulo2.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellSubtitulo2.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
      tblTitulo2.AddCell(cellSubtitulo2)

      Dim cellTituloContenedor2 As New iTextSharp.text.pdf.PdfPCell(tblTitulo2)
      cellTituloContenedor2.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTituloContenedor2.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE
      tblHeader2.AddCell(cellTituloContenedor2)

      Dim cellLineaHeader2 As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(""))
      cellLineaHeader2.Colspan = 2
      cellLineaHeader2.Border = iTextSharp.text.pdf.PdfPCell.BOTTOM_BORDER
      cellLineaHeader2.BorderColorBottom = azulOscuro
      cellLineaHeader2.BorderWidthBottom = 2.2F
      cellLineaHeader2.FixedHeight = 4.0F
      tblHeader2.AddCell(cellLineaHeader2)

      documento.Add(tblHeader2)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      ' Aviso pago tardío.
      Dim tblLateTitle As New iTextSharp.text.pdf.PdfPTable(1)
      tblLateTitle.TotalWidth = 540.0F
      tblLateTitle.LockedWidth = True

      Dim cellLateTitle As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "IMPORTANTE — PAGO TARDÍO", f10BoldWhite)
    )
      cellLateTitle.BackgroundColor = azul
      cellLateTitle.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellLateTitle.Padding = 9.0F
      tblLateTitle.AddCell(cellLateTitle)
      documento.Add(tblLateTitle)

      Dim tblLateBody As New iTextSharp.text.pdf.PdfPTable(2)
      tblLateBody.TotalWidth = 540.0F
      tblLateBody.LockedWidth = True
      tblLateBody.SetWidths(New Single() {435.0F, 105.0F})

      Dim cellLateText As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "A partir del mes de AGOSTO DE 2026, los pagos realizados después de la fecha límite establecida generarán un cargo administrativo por pago tardío de " &
        FormatCurrency(cargoPagoTardio, 2) &
        " (Cincuenta pesos 00/100 M.N.), el cual se aplicará de forma inmediata." &
        Environment.NewLine & Environment.NewLine &
        "En caso de suspensión del servicio por falta de pago, este cargo deberá liquidarse junto con la mensualidad vencida para la reactivación del servicio. Le invitamos a pagar dentro de la fecha establecida para evitar cargos adicionales.",
        f8Black
      )
    )
      cellLateText.BackgroundColor = amarilloClaro
      cellLateText.BorderColor = New iTextSharp.text.Color(232, 220, 170)
      cellLateText.BorderWidth = 0.7F
      cellLateText.Padding = 11.0F
      tblLateBody.AddCell(cellLateText)

      Dim cellLateAmount As New iTextSharp.text.pdf.PdfPCell()
      cellLateAmount.BackgroundColor = amarilloClaro
      cellLateAmount.BorderColor = New iTextSharp.text.Color(232, 220, 170)
      cellLateAmount.BorderWidth = 0.7F
      cellLateAmount.Padding = 10.0F

      Dim pLateAmount As New iTextSharp.text.Paragraph(
      FormatCurrency(cargoPagoTardio, 2),
      New iTextSharp.text.Font(
        iTextSharp.text.Font.HELVETICA, 18.0F,
        iTextSharp.text.Font.BOLD,
        New iTextSharp.text.Color(125, 92, 18)
      )
    )
      pLateAmount.Alignment = iTextSharp.text.Element.ALIGN_CENTER
      cellLateAmount.AddElement(pLateAmount)

      Dim pLateText As New iTextSharp.text.Paragraph(
      "CARGO POR" & Environment.NewLine & "PAGO TARDÍO",
      New iTextSharp.text.Font(
        iTextSharp.text.Font.HELVETICA, 7.0F,
        iTextSharp.text.Font.BOLD,
        New iTextSharp.text.Color(125, 92, 18)
      )
    )
      pLateText.Alignment = iTextSharp.text.Element.ALIGN_CENTER
      cellLateAmount.AddElement(pLateText)

      tblLateBody.AddCell(cellLateAmount)
      documento.Add(tblLateBody)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      ' Código para pago en tiendas.
      Dim tblTiendas As New iTextSharp.text.pdf.PdfPTable(1)
      tblTiendas.TotalWidth = 540.0F
      tblTiendas.LockedWidth = True

      Dim cellTiendasTitulo As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
        "CÓDIGO PARA PAGO EN TIENDAS", f10BoldBlue)
    )
      cellTiendasTitulo.Border = PdfCell.NO_BORDER
      cellTiendasTitulo.BorderWidthBottom = 0
      cellTiendasTitulo.PaddingTop = 10.0F
      cellTiendasTitulo.HorizontalAlignment = 1
      cellTiendasTitulo.Colspan = 5
      cellTiendasTitulo.BorderColorBottom = Color.WHITE
      tblTiendas.AddCell(cellTiendasTitulo)

      Dim imagenTiendas As iTextSharp.text.Image =
      iTextSharp.text.Image.GetInstance(
          Application.StartupPath & "/imgs/stores_3.jpg"
      )

      imagenTiendas.ScalePercent(50)

      Dim cellImagenTiendas As New iTextSharp.text.pdf.PdfPCell(imagenTiendas)

      cellImagenTiendas.Border = PdfCell.NO_BORDER
      cellImagenTiendas.BorderWidth = 0
      cellImagenTiendas.BorderWidthTop = 0
      cellImagenTiendas.BorderWidthBottom = 0
      cellImagenTiendas.HorizontalAlignment =
    iTextSharp.text.Element.ALIGN_CENTER


      cellImagenTiendas.Padding = 6.0F

      tblTiendas.AddCell(cellImagenTiendas)


      '========================================================================
      ' CÓDIGO DE BARRAS Y REFERENCIA
      '========================================================================
      Dim tblCodigoInterior As New iTextSharp.text.pdf.PdfPTable(1)
      tblCodigoInterior.WidthPercentage = 100

      '-----------------------------------------------------------------------
      ' IMAGEN DEL CÓDIGO DE BARRAS
      ' Se replica el comportamiento de la versión original:
      ' la imagen se pasa directamente al PdfPCell.
      '-----------------------------------------------------------------------
      If Not String.IsNullOrWhiteSpace(codigoBarraOxxo) Then

        System.Net.ServicePointManager.SecurityProtocol =
        DirectCast(3072, System.Net.SecurityProtocolType)

        Dim imagenCodigo As iTextSharp.text.Image =
        iTextSharp.text.Image.GetInstance(codigoBarraOxxo)

        ' IMPORTANTE:
        ' No usamos AddElement(imagenCodigo).
        ' Tampoco hacemos ScaleToFit.
        ' La versión original coloca directamente la imagen en el PdfPCell.
        Dim cellImagenCodigo As New iTextSharp.text.pdf.PdfPCell(imagenCodigo)

        cellImagenCodigo.Border =
        iTextSharp.text.pdf.PdfPCell.NO_BORDER

        cellImagenCodigo.HorizontalAlignment =
        iTextSharp.text.Element.ALIGN_CENTER

        cellImagenCodigo.VerticalAlignment =
        iTextSharp.text.Element.ALIGN_MIDDLE

        cellImagenCodigo.PaddingTop = 5.0F
        cellImagenCodigo.PaddingBottom = 3.0F

        tblCodigoInterior.AddCell(cellImagenCodigo)

      End If

      '-----------------------------------------------------------------------
      ' REFERENCIA
      '-----------------------------------------------------------------------
      Dim cellReferencia As New iTextSharp.text.pdf.PdfPCell(
      New iTextSharp.text.Phrase(
          refOxxo,
          New iTextSharp.text.Font(
              iTextSharp.text.Font.COURIER,
              12.0F,
              iTextSharp.text.Font.BOLD,
              iTextSharp.text.Color.BLACK
          )
        )
      )

      cellReferencia.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellReferencia.HorizontalAlignment =
    iTextSharp.text.Element.ALIGN_CENTER

      cellReferencia.PaddingTop = 3.0F
      cellReferencia.PaddingBottom = 2.0F

      tblCodigoInterior.AddCell(cellReferencia)

      '-----------------------------------------------------------------------
      ' LEYENDA
      '-----------------------------------------------------------------------
      Dim cellAyuda As New iTextSharp.text.pdf.PdfPCell(
        New iTextSharp.text.Phrase(
            "Muestre este código de barras en caja o dicte los dígitos de la referencia.",
            f7
        )
    )

      cellAyuda.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellAyuda.HorizontalAlignment =
    iTextSharp.text.Element.ALIGN_CENTER

      cellAyuda.PaddingTop = 1.0F
      cellAyuda.PaddingBottom = 8.0F

      tblCodigoInterior.AddCell(cellAyuda)

      '-----------------------------------------------------------------------
      ' CONTENEDOR PRINCIPAL
      '-----------------------------------------------------------------------
      Dim cellCodigo As New iTextSharp.text.pdf.PdfPCell(tblCodigoInterior)

      cellCodigo.Border = PdfCell.NO_BORDER
      cellCodigo.BorderWidth = 0
      cellCodigo.BorderWidthTop = 0

      cellCodigo.HorizontalAlignment =
    iTextSharp.text.Element.ALIGN_CENTER

      cellCodigo.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_MIDDLE

      cellCodigo.Padding = 0.0F

      tblTiendas.AddCell(cellCodigo)

      documento.Add(tblTiendas)
      documento.Add(New iTextSharp.text.Paragraph(" ", f7))

      ' Instrucciones.


      'AGREGAR PASOS'
      '========================================================================
      ' INSTRUCCIONES PARA PAGO EN TIENDAS
      '========================================================================

      Dim tblTituloPasos As New iTextSharp.text.pdf.PdfPTable(1)
      tblTituloPasos.TotalWidth = 540.0F
      tblTituloPasos.LockedWidth = True

      Dim cellTituloPasos As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
        "INSTRUCCIONES PARA PAGO EN TIENDAS",
        f10BoldBlue
    )
)

      cellTituloPasos.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER
      cellTituloPasos.PaddingBottom = 7.0F

      tblTituloPasos.AddCell(cellTituloPasos)

      documento.Add(tblTituloPasos)


      ' Fuente específica para el texto de los pasos
      Dim fPasoTexto As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA,
    8.5F,
    iTextSharp.text.Font.NORMAL,
    iTextSharp.text.Color.BLACK
)

      Dim fPasoNumero As New iTextSharp.text.Font(
    iTextSharp.text.Font.HELVETICA,
    10.0F,
    iTextSharp.text.Font.BOLD,
    azul
)


      '========================================================================
      ' TABLA PRINCIPAL DE LOS 3 PASOS
      '========================================================================
      Dim tblPasos As New iTextSharp.text.pdf.PdfPTable(3)

      tblPasos.TotalWidth = 540.0F
      tblPasos.LockedWidth = True

      tblPasos.SetWidths(
    New Single() {
        180.0F,
        180.0F,
        180.0F
    }
)


      '========================================================================
      ' PASO 1
      '========================================================================
      Dim tblPaso1 As New iTextSharp.text.pdf.PdfPTable(1)
      tblPaso1.WidthPercentage = 100


      ' Número
      Dim cellNumero1 As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
        "1",
        fPasoNumero
    )
)

      cellNumero1.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellNumero1.PaddingLeft = 10.0F
      cellNumero1.PaddingRight = 10.0F
      cellNumero1.PaddingTop = 10.0F
      cellNumero1.PaddingBottom = 9.0F

      tblPaso1.AddCell(cellNumero1)


      ' Texto
      Dim cellTexto1 As New iTextSharp.text.pdf.PdfPCell(
          New iTextSharp.text.Phrase(
              "Elija la tienda que más le convenga entre las cadenas indicadas " &
              "(solo se puede pagar en esas tiendas).",
              fPasoTexto
          )
      )

      cellTexto1.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellTexto1.PaddingLeft = 10.0F
      cellTexto1.PaddingRight = 10.0F
      cellTexto1.PaddingTop = 0.0F
      cellTexto1.PaddingBottom = 10.0F

      cellTexto1.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblPaso1.AddCell(cellTexto1)


      ' Contenedor
      Dim contPaso1 As New iTextSharp.text.pdf.PdfPCell(tblPaso1)

      contPaso1.BorderColor = grisBorde
      contPaso1.BorderWidth = 0.8F
      contPaso1.Padding = 0.0F
      contPaso1.MinimumHeight = 62.0F
      contPaso1.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblPasos.AddCell(contPaso1)


      '========================================================================
      ' PASO 2
      '========================================================================
      Dim tblPaso2 As New iTextSharp.text.pdf.PdfPTable(1)
      tblPaso2.WidthPercentage = 100


      ' Número
      Dim cellNumero2 As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
        "2",
        fPasoNumero
    )
)

      cellNumero2.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellNumero2.PaddingLeft = 10.0F
      cellNumero2.PaddingRight = 10.0F
      cellNumero2.PaddingTop = 10.0F
      cellNumero2.PaddingBottom = 9.0F

      tblPaso2.AddCell(cellNumero2)


      ' Texto
      Dim cellTexto2 As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
        "Al acercarse al mostrador, mencione que viene a pagar CONEKTA " &
        "y muestre el código de barras o dicte los números de la referencia.",
        fPasoTexto
    )
)

      cellTexto2.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellTexto2.PaddingLeft = 10.0F
      cellTexto2.PaddingRight = 10.0F
      cellTexto2.PaddingTop = 0.0F
      cellTexto2.PaddingBottom = 10.0F

      cellTexto2.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblPaso2.AddCell(cellTexto2)


      ' Contenedor
      Dim contPaso2 As New iTextSharp.text.pdf.PdfPCell(tblPaso2)

      contPaso2.BorderColor = grisBorde
      contPaso2.BorderWidth = 0.8F
      contPaso2.Padding = 0.0F
      contPaso2.MinimumHeight = 62.0F
      contPaso2.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblPasos.AddCell(contPaso2)


      '========================================================================
      ' PASO 3
      '========================================================================
      Dim tblPaso3 As New iTextSharp.text.pdf.PdfPTable(1)
      tblPaso3.WidthPercentage = 100


      ' Número
      Dim cellNumero3 As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
        "3",
        fPasoNumero
    )
)

      cellNumero3.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellNumero3.PaddingLeft = 10.0F
      cellNumero3.PaddingRight = 10.0F
      cellNumero3.PaddingTop = 10.0F
      cellNumero3.PaddingBottom = 9.0F

      tblPaso3.AddCell(cellNumero3)


      ' Texto
      Dim cellTexto3 As New iTextSharp.text.pdf.PdfPCell(
        New iTextSharp.text.Phrase(
            "Una vez realizado el pago en efectivo, recibirá una notificación " &
            "de pago en tiempo real a su correo y ¡listo!",
            fPasoTexto
        )
    )

      cellTexto3.Border =
    iTextSharp.text.pdf.PdfPCell.NO_BORDER

      cellTexto3.PaddingLeft = 10.0F
      cellTexto3.PaddingRight = 10.0F
      cellTexto3.PaddingTop = 0.0F
      cellTexto3.PaddingBottom = 10.0F

      cellTexto3.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblPaso3.AddCell(cellTexto3)


      ' Contenedor
      Dim contPaso3 As New iTextSharp.text.pdf.PdfPCell(tblPaso3)

      contPaso3.BorderColor = grisBorde
      contPaso3.BorderWidth = 0.8F
      contPaso3.Padding = 0.0F
      contPaso3.MinimumHeight = 62.0F
      contPaso3.VerticalAlignment =
    iTextSharp.text.Element.ALIGN_TOP

      tblPasos.AddCell(contPaso3)


      '========================================================================
      ' AGREGAR AL DOCUMENTO
      '========================================================================
      documento.Add(tblPasos)

      AgregarFooter(
      documento, logo2, grisBorde, f7, "PÁGINA 2 DE 2"
    )

      ' Close() termina de escribir y cierra writer/stream.
      documento.Close()

    Catch ex As Exception

      Try
        If documento IsNot Nothing AndAlso documento.IsOpen Then
          documento.Close()
        End If
      Catch
      End Try

      If System.IO.File.Exists(ruta) Then
        Try
          System.IO.File.Delete(ruta)
        Catch
        End Try
      End If

      MsgBox(ex.Message & Environment.NewLine & ex.StackTrace)

    Finally
      writer = Nothing
      documento = Nothing
    End Try
  End Sub


  Private Sub AgregarFilaCargo(ByVal tabla As iTextSharp.text.pdf.PdfPTable,
                                  ByVal plan As String,
                                  ByVal concepto As String,
                                  ByVal importe As Decimal,
                                  ByVal colorBorde As iTextSharp.text.Color,
                                  ByVal fuentePlan As iTextSharp.text.Font,
                                  ByVal fuenteImporte As iTextSharp.text.Font)

    Dim cellPlan As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(plan, fuentePlan)
  )
    cellPlan.BorderColor = colorBorde
    cellPlan.Padding = 7.0F
    tabla.AddCell(cellPlan)

    Dim cellConcepto As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
      concepto,
      New iTextSharp.text.Font(
        iTextSharp.text.Font.HELVETICA, 8.0F,
        iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.BLACK)
    )
  )
    cellConcepto.BorderColor = colorBorde
    cellConcepto.Padding = 7.0F
    tabla.AddCell(cellConcepto)

    Dim cellImporte As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(FormatCurrency(importe, 2), fuenteImporte)
  )
    cellImporte.BorderColor = colorBorde
    cellImporte.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
    cellImporte.Padding = 7.0F
    tabla.AddCell(cellImporte)
  End Sub


  Private Sub AgregarSubtotal(ByVal tabla As iTextSharp.text.pdf.PdfPTable,
                                 ByVal texto As String,
                                 ByVal importe As Decimal,
                                 ByVal colorFondo As iTextSharp.text.Color,
                                 ByVal colorBorde As iTextSharp.text.Color,
                                 ByVal fuente As iTextSharp.text.Font)

    Dim cellTexto As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(texto, fuente)
  )
    cellTexto.Colspan = 2
    cellTexto.BackgroundColor = colorFondo
    cellTexto.BorderColor = colorBorde
    cellTexto.Padding = 5.0F
    tabla.AddCell(cellTexto)

    Dim cellImporte As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(FormatCurrency(importe, 2), fuente)
  )
    cellImporte.BackgroundColor = colorFondo
    cellImporte.BorderColor = colorBorde
    cellImporte.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
    cellImporte.Padding = 5.0F
    tabla.AddCell(cellImporte)
  End Sub





  Private Sub AgregarFooter(ByVal documento As iTextSharp.text.Document,
                               ByVal logoOriginal As iTextSharp.text.Image,
                               ByVal colorBorde As iTextSharp.text.Color,
                               ByVal fuente As iTextSharp.text.Font,
                               ByVal numeroPagina As String)

    documento.Add(New iTextSharp.text.Paragraph(
    Environment.NewLine & Environment.NewLine, fuente))

    Dim tblFooter As New iTextSharp.text.pdf.PdfPTable(2)
    tblFooter.TotalWidth = 540.0F
    tblFooter.LockedWidth = True
    tblFooter.SetWidths(New Single() {390.0F, 150.0F})

    Dim cellContacto As New iTextSharp.text.pdf.PdfPCell(
    New iTextSharp.text.Phrase(
      "soporte_residencial@comunicalo.mx  ·  Atención a clientes: 55 2601 4010" &
      Environment.NewLine & "Horario de atención de 9 a 18 hrs",
      fuente
    )
  )
    cellContacto.Border = iTextSharp.text.pdf.PdfPCell.TOP_BORDER
    cellContacto.BorderColorTop = colorBorde
    cellContacto.PaddingTop = 8.0F
    tblFooter.AddCell(cellContacto)

    Dim logoFooter As iTextSharp.text.Image =
    iTextSharp.text.Image.GetInstance(
      Application.StartupPath & "/imgs/LOGOCOMUNICALO.png"
    )
    logoFooter.ScaleToFit(95.0F, 35.0F)

    Dim cellLogo As New iTextSharp.text.pdf.PdfPCell(logoFooter)
    cellLogo.Border = iTextSharp.text.pdf.PdfPCell.TOP_BORDER
    cellLogo.BorderColorTop = colorBorde
    cellLogo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT
    cellLogo.PaddingTop = 6.0F
    tblFooter.AddCell(cellLogo)

    documento.Add(tblFooter)

    Dim pPagina As New iTextSharp.text.Paragraph(numeroPagina, fuente)
    pPagina.Alignment = iTextSharp.text.Element.ALIGN_CENTER
    documento.Add(pPagina)
  End Sub




  Private Sub Generar_pdfOXXO_v2(ByVal id_estado_cuenta As Integer, ByVal id_contrato As Integer, ByVal path As String, ByVal refOxxo As String, ByVal codigoBarraOxxo As String)
    Dim sqledo As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtedo As DataTable = con.ConsultarDT(sqledo)
    If dtedo IsNot Nothing AndAlso dtedo.Rows.Count > 0 Then
      Dim fecha As Date = dtedo(0)("fecha").ToString
      Dim grantotal As Double = Val(dtedo(0)("grantotal").ToString)
      Dim saldo_pendiente As Double = Val(dtedo(0)("saldo_pendiente").ToString)
      Dim total_edo As Double = grantotal - saldo_pendiente
      Dim periodoA As Date = dtedo(0)("periodoA").ToString
      Dim periodoB As Date = dtedo(0)("periodoB").ToString
      Dim totalPlan As Double = Val(dtedo(0)("mensualidad").ToString())

      Dim lateFeeAmount As Double = 50.0
      Dim showLateFeeHistory As Boolean = False
      Dim lateFeeHistoryConcept As String = "CARGO ADMINISTRATIVO POR PAGO TARDÍO - PERIODO ANTERIOR"
      Dim sqlcli As String = ""
      Dim sqlBalance = "select * from CONTRACTS_BALANCES where id_contrato=" & id_contrato & ";"
      Dim dtBalance = con.ConsultarDT(sqlBalance)
      Dim balance As Double = 0

      If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
        balance = Val(dtBalance(0)("balance").ToString)
      End If

      If tiene_telefonia(id_contrato) Then
        sqlcli = $"SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero,t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM (" &
" SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete FROM(" &
" SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre) AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias" &
" FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e" &
" ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato &
" ) AS t1 INNER JOIN CALLES ca ON ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete)" &
" AS t3 INNER JOIN dbo.EQUIPOS e INNER JOIN EQUIPOS_TELEFONIA et INNER JOIN LINEAS l" &
" ON l.id_linea=et.id_linea ON et.id_equipo=e.id_equipo ON e.id_contrato=t3.id_contrato  where e.estatus=1 AND et.estatus=1"
      Else
        sqlcli = "SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero," &
" t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM" &
" ( SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete,numero" &
" FROM( SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre)" &
 " AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias,telefono AS numero FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c " &
 " INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON" &
 " col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato & " ) AS t1 INNER JOIN CALLES ca ON" &
 " ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete) AS t3"
      End If

      Dim dtcli As DataTable = con.ConsultarDT(sqlcli)
      If dtcli IsNot Nothing AndAlso dtcli.Rows.Count > 0 Then
        Dim nombre As String = dtcli(0)("nombre").ToString
        Dim contrato As String = dtcli(0)("contrato").ToString
        Dim contract As String = dtcli(0)("contrato").ToString
        Dim calle As String = dtcli(0)("calle").ToString
        Dim numext As String = dtcli(0)("numext").ToString
        Dim numint As String = dtcli(0)("numint").ToString
        Dim colonia As String = dtcli(0)("colonia").ToString
        Dim cp As String = dtcli(0)("cp").ToString
        Dim municipio As String = dtcli(0)("municipio").ToString
        Dim estado As String = dtcli(0)("estado").ToString
        Dim referencias As String = dtcli(0)("referencias").ToString
        Dim paquete As String = dtcli(0)("paquete").ToString
        Dim numero As String = dtcli(0)("numero").ToString
        Dim phone As String = dtcli(0)("numero").ToString
        Dim id_paquete As Integer = Val(dtcli(0)("id_paquete").ToString)
        Dim servicios As String = getServicios(id_paquete)
        Dim dtAccount As DataTable = getStpAccount(id_contrato)
        Dim account As String = "s/a"

        If dtAccount IsNot Nothing AndAlso dtAccount.Rows.Count > 0 Then
          account = dtAccount.Rows(0)("clave").ToString()
        End If

        Dim ruta As String = path & "\EstadoCuenta(" & id_estado_cuenta.ToString & ").pdf "
        Dim oDoc As New iTextSharp.text.Document(PageSize.LETTER, 50, 50, 50, 50)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim linea As PdfContentByte
        Dim rectangulo As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Try
          pdfw = PdfWriter.GetInstance(oDoc, New FileStream(ruta,
                    FileMode.Create, FileAccess.Write, FileShare.None))

          Me.PageState = New CustomPageState()
          ''//Wire our event handler and pass in the page state
          pdfw.PageEvent = New MyCustomPdfEvent(Me.PageState)
          'Apertura del documento.
          oDoc.Open()
          cb = pdfw.DirectContent
          linea = pdfw.DirectContent
          rectangulo = pdfw.DirectContent

          'Agregamos una pagina. // check later
          'oDoc.NewPage()

          cb.BeginText()
          fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
          cb.SetFontAndSize(fuente, 10) 'fuente definida en la linea anterior y tamaño

          Dim f10 As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f10.SetColor(2, 51, 130)

          Dim f10Bold As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10Bold.SetColor(2, 51, 130)

          Dim f10BoldMain As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10BoldMain.SetColor(2, 51, 130)


          Dim f14 As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f14.SetColor(2, 51, 130)

          Dim tblBanner As New PdfPTable(1)
          tblBanner.HorizontalAlignment = 0
          tblBanner.LockedWidth = True
          tblBanner.TotalWidth = 550.0F
          tblBanner.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBanner.DefaultCell.MinimumHeight = 12
          tblBanner.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER
          tblBanner.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          'tblBanner.SetWidthPercentage({140.0F, 100.0F, 300.0F}, PageSize.LETTER)

          'Dim banner As iTextSharp.text.Image
          'banner = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/banner.jpg") 'nombre y ruta de la imagen a insertar
          'imagen.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)
          'tblBanner.AddCell(banner)
          'oDoc.Add(tblBanner)
          oDoc.Add(New Paragraph(" "))

          Dim tblHeaderInfo As New PdfPTable(3)
          tblHeaderInfo.HorizontalAlignment = 0
          tblHeaderInfo.LockedWidth = True
          tblHeaderInfo.TotalWidth = 540.0F
          tblHeaderInfo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblHeaderInfo.DefaultCell.MinimumHeight = 12
          tblHeaderInfo.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeaderInfo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblHeaderInfo.SetWidthPercentage({220.0F, 50.0F, 270.0F}, PageSize.LETTER)

          'IMAGEN
          Dim imagenInfo As iTextSharp.text.Image
          imagenInfo = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
          imagenInfo.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)
          'tblHeaderInfo.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))

          Dim cellInfoCompany As New PdfPTable(1)
          cellInfoCompany.DefaultCell.Border = PdfPCell.NO_BORDER
          cellInfoCompany.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT
          ' Comunicalo info.
          cellInfoCompany.AddCell(imagenInfo)
          cellInfoCompany.AddCell(New Phrase("Comunícalo de México S.A. de C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("CONVENTO DE CHURUBUSCO NO. 4,", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("COL. JARDINES DE SANTA MÓNICA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("MPIO. TLALNEPANTLA DE BAZ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("ESTADO DE MÉXICO", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("C.P. 54050", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("RFC: CME0806162SA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(nombre, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          ' Contract info.
          cellInfoCompany.AddCell(New Phrase(calle, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(referencias & " " & numext & " " & numint, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(colonia, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(municipio & ", " & estado & ", C.P. " & cp, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'cellInfoCompany.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("SERVICIOS CONTRATADOS", f10Bold))

          Dim nesthousingInfo As New PdfPCell(cellInfoCompany)
          nesthousingInfo.Border = PdfPCell.NO_BORDER
          nesthousingInfo.Padding = 0F
          nesthousingInfo.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeaderInfo.AddCell(nesthousingInfo)

          ' Header right column , information about bill.
          Dim cellBillTitle As New PdfPCell(New Phrase("ESTADO DE CUENTA", f10BoldMain))
          cellBillTitle.Border = PdfPCell.BOTTOM_BORDER
          cellBillTitle.BorderWidthBottom = 4
          cellBillTitle.PaddingTop = 12.0F
          cellBillTitle.PaddingBottom = 15.0F
          cellBillTitle.HorizontalAlignment = 1
          cellBillTitle.Colspan = 1
          cellBillTitle.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          Dim tblBillTitle As New PdfPTable(1)
          tblBillTitle.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBillTitle.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblBillTitle.AddCell(cellBillTitle)
          tblBillTitle.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim tblBillItems As New PdfPTable(2)
          tblBillItems.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBillItems.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT
          tblBillItems.DefaultCell.PaddingLeft = 7.0F
          tblBillItems.DefaultCell.PaddingRight = 7.0F
          tblBillItems.DefaultCell.PaddingBottom = 3.0F

          tblBillItems.AddCell(New Phrase("MES DE FACTURACIÓN", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(MonthName(periodoA.Month).ToUpper, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase("CONTRATO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(contract, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblBillItems.AddCell(New Phrase("TELÉFONO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(phone, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase("TOTAL A PAGAR", f10Bold))
          tblBillItems.AddCell(New Phrase(FormatCurrency(grantotal, 2), f10Bold))
          tblBillItems.AddCell(New Phrase("DESPUÉS DEL " & periodoA.ToString("dd/MM/yyyy"), f10Bold))
          tblBillItems.AddCell(New Phrase(FormatCurrency(grantotal + 50, 2), f10Bold))
          tblBillItems.AddCell(New Phrase("FECHA LÍMITE", f10Bold))
          tblBillItems.AddCell(New Phrase(periodoA.ToString("dd/MM/yyyy"), f10Bold))
          tblBillItems.AddCell(New Phrase("SALDO VENCIDO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          ' Se quita el banner morado de términos/consulta del paquete según diseño solicitado.
          tblBillTitle.AddCell(tblBillItems)
          tblHeaderInfo.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))
          tblHeaderInfo.AddCell(tblBillTitle)

          Dim cellSeparator As New PdfPCell(New Phrase("", f10))
          cellSeparator.Border = PdfPCell.BOTTOM_BORDER
          cellSeparator.BorderWidthBottom = 2
          cellSeparator.PaddingTop = 1.0F
          cellSeparator.PaddingBottom = 1.0F
          cellSeparator.HorizontalAlignment = 1
          cellSeparator.Colspan = 3
          cellSeparator.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))
          tblHeaderInfo.AddCell(cellSeparator)

          oDoc.Add(tblHeaderInfo)
          'oDoc.Add(New Paragraph(" "))

          Dim cellEspacio As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellEspacio.Border = PdfPCell.NO_BORDER
          cellEspacio.BorderWidthBottom = 0
          cellEspacio.PaddingTop = 5.0F
          cellEspacio.HorizontalAlignment = 1
          cellEspacio.Colspan = 5
          cellEspacio.BorderColorBottom = Color.WHITE

          Dim tblPeriodo As New PdfPTable(5)
          tblPeriodo.HorizontalAlignment = 0
          tblPeriodo.LockedWidth = True
          tblPeriodo.TotalWidth = 540.0F
          tblPeriodo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblPeriodo.DefaultCell.MinimumHeight = 12
          tblPeriodo.DefaultCell.HorizontalAlignment = 0
          tblPeriodo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblPeriodo.DefaultCell.PaddingLeft = 12.0F
          tblPeriodo.SetWidthPercentage({150.0F, 80.0F, 40.0F, 125.0F, 145.0F}, PageSize.LETTER)

          Dim cellPeriodo3 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPeriodo3.Border = PdfPCell.BOTTOM_BORDER
          cellPeriodo3.BorderWidthBottom = 2
          cellPeriodo3.PaddingTop = 0
          cellPeriodo3.HorizontalAlignment = 0
          cellPeriodo3.Colspan = 5
          cellPeriodo3.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          Dim cellPaqueteContratado As New PdfPCell(New Phrase("Cargos del mes", f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 5.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPaqueteContratado)
          cellPaqueteContratado = New PdfPCell(New Phrase(paquete, f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 1.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPaqueteContratado)

          Dim cellServicios As New PdfPCell(New Phrase(servicios, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServicios.Border = PdfPCell.NO_BORDER
          cellServicios.BorderWidth = 0
          cellServicios.PaddingTop = 0
          cellServicios.HorizontalAlignment = 0
          cellServicios.Colspan = 4
          cellServicios.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellServicios)
          tblPeriodo.AddCell(New Phrase(FormatCurrency(totalPlan, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          If showLateFeeHistory Then
            Dim cellLateFeeHistory As New PdfPCell(New Phrase(lateFeeHistoryConcept, New Font(iTextSharp.text.Font.HELVETICA, 9.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellLateFeeHistory.Border = PdfPCell.NO_BORDER
            cellLateFeeHistory.BorderWidth = 0
            cellLateFeeHistory.PaddingTop = 0
            cellLateFeeHistory.HorizontalAlignment = 0
            cellLateFeeHistory.Colspan = 4
            cellLateFeeHistory.BorderColor = Color.WHITE

            tblPeriodo.AddCell(cellLateFeeHistory)
            tblPeriodo.AddCell(New Phrase(FormatCurrency(lateFeeAmount, 2), New Font(iTextSharp.text.Font.HELVETICA, 9.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          End If

          ' Before "saldo pendiente"
          Dim cellPending As New PdfPCell(New Phrase("SALDO PENDIENTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPending.Border = PdfPCell.NO_BORDER
          cellPending.BorderWidth = 0
          cellPending.PaddingTop = 0
          cellPending.HorizontalAlignment = 0
          cellPending.Colspan = 4
          cellPending.BorderColor = Color.WHITE

          'tblPeriodo.AddCell(cellPending)
          'tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          'Show balance'
          Dim auxBalance As Double = If(balance < 0, balance * -1, balance)
          Dim showBalance As Boolean = If(balance <= 0, True, False)

          'If balance < 0 Then
          '  auxBalance = balance * -1
          'End If

          If showBalance Then
            Dim auxShowBalance As Double = auxBalance + grantotal
            Dim cellShowBalance As New PdfPCell(New Phrase("SALDO A FAVOR", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellShowBalance.Border = PdfPCell.NO_BORDER
            cellShowBalance.BorderWidth = 0
            cellShowBalance.PaddingTop = 0
            cellShowBalance.HorizontalAlignment = 0
            cellShowBalance.Colspan = 4
            cellShowBalance.BorderColor = Color.WHITE

            'tblPeriodo.AddCell(cellShowBalance)
            'tblPeriodo.AddCell(New Phrase(FormatCurrency(auxShowBalance, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          End If

          Dim dtCharges As DataTable = getDataBillCharges(id_estado_cuenta)

          If dtCharges IsNot Nothing AndAlso dtCharges.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("OTROS CARGOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtCharges.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtCharges.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtCharges.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim dtDiscount As DataTable = getDataBillDiscounts(id_estado_cuenta)

          If dtDiscount IsNot Nothing AndAlso dtDiscount.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("DESCUENTOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtDiscount.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtDiscount.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtDiscount.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim celltotal As New PdfPCell(New Phrase("TOTAL A PAGAR " & FormatCurrency(grantotal, 2), f14))
          celltotal.Border = PdfPCell.NO_BORDER
          celltotal.BorderWidth = 0
          celltotal.PaddingTop = 10.0F
          celltotal.PaddingLeft = 12.0F
          celltotal.HorizontalAlignment = 0
          celltotal.Colspan = 2
          celltotal.BorderColor = Color.WHITE

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(celltotal)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          Dim celltotalLetra As New PdfPCell(New Phrase("(" & totalLetra(grantotal) & ")", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          celltotalLetra.Border = PdfPCell.NO_BORDER
          celltotalLetra.BorderWidth = 0
          celltotalLetra.PaddingTop = 0
          celltotalLetra.PaddingLeft = 12.0F
          celltotalLetra.HorizontalAlignment = 0
          celltotalLetra.Colspan = 2
          celltotalLetra.BorderColor = Color.WHITE

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(celltotalLetra)

          Dim advisoryPlan As String = $"*En el caso de haber realizado un cambio o actualización en su paquete, al realizar el pago de este Estado de Cuenta, usted acepta los nuevos términos y Condiciones aplicables."
          Dim cellAdvisoryPlan As New PdfPCell(New Phrase(advisoryPlan, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellAdvisoryPlan.Border = PdfPCell.NO_BORDER
          cellAdvisoryPlan.BorderWidth = 0
          cellAdvisoryPlan.PaddingTop = 0
          cellAdvisoryPlan.PaddingLeft = 0
          cellAdvisoryPlan.HorizontalAlignment = 0
          cellAdvisoryPlan.Colspan = 5
          cellAdvisoryPlan.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellAdvisoryPlan)
          tblPeriodo.AddCell(cellPeriodo3)
          tblPeriodo.AddCell(cellEspacio)

          'Balance'
          If showBalance Then
            Dim cellBalance As New PdfPCell(New Phrase("SALDO A FAVOR RESTANTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellBalance.Border = PdfPCell.NO_BORDER
            cellBalance.BorderWidth = 0
            cellBalance.PaddingTop = 0
            cellBalance.HorizontalAlignment = 0
            cellBalance.Colspan = 4
            cellBalance.BorderColor = Color.WHITE

            'tblPeriodo.AddCell(cellBalance)
            'tblPeriodo.AddCell(New Phrase(FormatCurrency(auxBalance, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            'tblPeriodo.AddCell(cellPeriodo3)
          End If

          ' Se quita advisory_2.jpg / warning ATENCIÓN; los clientes ya tienen clara esa indicación.
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)

          ' Warning payments
          Dim phWarning As Phrase = New Phrase("IMPORTANTE SOBRE SU PAGO.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.RED))
          Dim cellPaymentWarning As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellPaymentWarning = New PdfPCell(phWarning)
          cellPaymentWarning.Border = PdfPCell.NO_BORDER
          cellPaymentWarning.BorderWidthBottom = 0
          cellPaymentWarning.PaddingTop = 1.0F
          cellPaymentWarning.HorizontalAlignment = 0
          cellPaymentWarning.Colspan = 5
          cellPaymentWarning.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellPaymentWarning)

          Dim warningContent As String = "Cada contrato tiene su propia CLABE interbancaria. No utilice la CLABE de un contrato para pagar otro diferente.Los pagos no pueden transferirse entre contratos y quedarán como saldo a favor del contrato asociado a la CLABE utilizada."
          Dim phWarningContent As Phrase = New Phrase(warningContent)

          Dim cellPaymentWarningContent As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.RED)))
          cellPaymentWarningContent = New PdfPCell(phWarningContent)
          cellPaymentWarningContent.Border = PdfPCell.NO_BORDER
          cellPaymentWarningContent.BorderWidthBottom = 0
          cellPaymentWarningContent.PaddingTop = 1.0F
          cellPaymentWarningContent.HorizontalAlignment = 0
          cellPaymentWarningContent.Colspan = 5
          cellPaymentWarningContent.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellPaymentWarningContent)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)

          Dim cellFormasPago As New PdfPCell(New Phrase("FORMAS DE PAGO", f10Bold))
          cellFormasPago.Border = PdfPCell.NO_BORDER
          cellFormasPago.BorderWidthBottom = 0
          cellFormasPago.PaddingTop = 2.0F
          cellFormasPago.HorizontalAlignment = 0
          cellFormasPago.Colspan = 5
          cellFormasPago.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellFormasPago)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)

          Dim cellStp As New PdfPCell(New Phrase("ATENCIÓN", New Font(iTextSharp.text.Font.HELVETICA, 11.0F, iTextSharp.text.Font.BOLD, Color.RED)))
          cellStp.Border = PdfPCell.NO_BORDER
          cellStp.BorderWidthBottom = 0
          cellStp.PaddingTop = 10.0F
          cellStp.HorizontalAlignment = 0
          cellStp.Colspan = 5
          'tblPeriodo.AddCell(cellStp)

          Dim instructions As String = "A PARTIR DE AHORA CADA CLIENTE TENDRÁ UNA CLABE INTERBANCARIA ÚNICA Y PERSONALIZADA POR CONTRATO."
          Dim cellStpInstructions As New PdfPCell(New Phrase(instructions, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellStpInstructions.Border = PdfPCell.NO_BORDER
          cellStpInstructions.BorderWidthBottom = 0
          cellStpInstructions.PaddingTop = 10.0F
          cellStpInstructions.HorizontalAlignment = 0
          cellStpInstructions.Colspan = 5
          cellStpInstructions.BorderColorBottom = Color.BLACK
          'tblPeriodo.AddCell(cellStpInstructions)

          Dim cellDeposito As New PdfPCell(New Phrase("DATOS PARA TRANSFERENCIA:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellDeposito.Border = PdfPCell.NO_BORDER
          cellDeposito.BorderWidthBottom = 0
          cellDeposito.PaddingTop = 1.0F
          cellDeposito.HorizontalAlignment = 0
          cellDeposito.Colspan = 5
          cellDeposito.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellDeposito)

          Dim cellTransfer As New PdfPCell(New Phrase("TRANSFERENCIA ELECTRÓNICA:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTransfer.Border = PdfPCell.NO_BORDER
          cellTransfer.BorderWidthBottom = 0
          cellTransfer.PaddingTop = 10.0F
          cellTransfer.HorizontalAlignment = 1
          cellTransfer.Colspan = 2
          cellTransfer.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellTransfer)

          Dim cellFormasPago2 As New PdfPCell(New Phrase("BANCO: STP", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago2.Border = PdfPCell.NO_BORDER
          cellFormasPago2.BorderWidthBottom = 0
          cellFormasPago2.PaddingTop = 2.0F
          cellFormasPago2.HorizontalAlignment = 0
          cellFormasPago2.Colspan = 5
          cellFormasPago2.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago2)

          Dim cellClabe As New PdfPCell(New Phrase("CLABE INTERBANCARIA: 044180256007653656", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellClabe.Border = PdfPCell.NO_BORDER
          cellClabe.BorderWidthBottom = 0
          cellClabe.PaddingTop = 2.0F
          cellClabe.HorizontalAlignment = 2
          cellClabe.Colspan = 2
          cellClabe.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellClabe)

          Dim cellFormasPago3 As New PdfPCell(New Phrase("BENEFICIARIO: COMUNICALO DE MÉXICO, S.A. DE C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago3.Border = PdfPCell.NO_BORDER
          cellFormasPago3.BorderWidthBottom = 0
          cellFormasPago3.PaddingTop = 0.0F
          cellFormasPago3.HorizontalAlignment = 0
          cellFormasPago3.Colspan = 5
          cellFormasPago3.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago3)
          Dim boldTextAccount As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          Dim clabeText As Chunk = New Chunk(account, boldTextAccount)
          Dim indClabe As String = "CLABE INTERBANCARIA PERSONALIZADA: "
          Dim phClabe As Phrase = New Phrase(indClabe, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          phClabe.Add(clabeText)

          Dim cellFormasPago4 As New PdfPCell(New Phrase("CLABE INTERBANCARIA PERSONALIZADA: " + Environment.NewLine + account, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago4 = New PdfPCell(phClabe)
          cellFormasPago4.Border = PdfPCell.NO_BORDER
          cellFormasPago4.BorderWidthBottom = 0
          cellFormasPago4.PaddingTop = 0.0F
          cellFormasPago4.HorizontalAlignment = 0
          cellFormasPago4.Colspan = 5
          cellFormasPago4.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago4)
          'tblPeriodo.AddCell(phClabe)
          tblPeriodo.AddCell(cellEspacio)

          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          ' Late fee.
          Dim lateFeeTitle As Font = New Font(iTextSharp.text.Font.HELVETICA, 11.0F, iTextSharp.text.Font.BOLD, New Color(255, 204, 0))
          Dim lateFeeText As Font = New Font(iTextSharp.text.Font.HELVETICA, 9.0F, iTextSharp.text.Font.BOLD, Color.WHITE)

          Dim phLateFee As New Phrase()
          phLateFee.Add(New Chunk("——  IMPORTANTE - PAGO TARDÍO  ——" & Environment.NewLine & Environment.NewLine, lateFeeTitle))

          phLateFee.Add(New Chunk(
"A partir del mes de AGOSTO DE 2026, los pagos realizados después de la fecha límite establecida generarán un cargo administrativo por pago tardío de $50.00 (Cincuenta pesos 00/100 M.N.), el cual se aplicará de forma inmediata." &
Environment.NewLine & Environment.NewLine &
"En caso de suspensión del servicio por falta de pago, este cargo deberá liquidarse junto con la mensualidad vencida para la reactivación del servicio." &
Environment.NewLine &
"Le invitamos a pagar dentro de la fecha establecida para evitar cargos adicionales.",
lateFeeText))

          Dim cellLateFee As New PdfPCell(phLateFee)
          With cellLateFee
            .BackgroundColor = New Color(System.Drawing.ColorTranslator.FromHtml("#08154D")) 'azul oscuro
            .Border = PdfPCell.NO_BORDER
            .Colspan = 5
            .HorizontalAlignment = Element.ALIGN_CENTER
            .VerticalAlignment = Element.ALIGN_MIDDLE
            .PaddingTop = 12
            .PaddingBottom = 12
            .PaddingLeft = 18
            .PaddingRight = 18
          End With

          tblPeriodo.AddCell(cellLateFee)
          tblPeriodo.AddCell(cellEspacio)


          ' Primera hoja: información del estado de cuenta, cargos, formas de pago y aviso de pago tardío.
          oDoc.Add(tblPeriodo)

          ' Segunda hoja: tiendas, código de barras e instrucciones de pago.
          'oDoc.NewPage()

          Dim tblPagoTiendas As New PdfPTable(5)
          tblPagoTiendas.HorizontalAlignment = 0
          tblPagoTiendas.LockedWidth = True
          tblPagoTiendas.TotalWidth = 540.0F
          tblPagoTiendas.DefaultCell.Border = PdfPCell.NO_BORDER
          tblPagoTiendas.DefaultCell.MinimumHeight = 12
          tblPagoTiendas.DefaultCell.HorizontalAlignment = 0
          tblPagoTiendas.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblPagoTiendas.DefaultCell.PaddingLeft = 12.0F
          tblPagoTiendas.SetWidthPercentage({150.0F, 80.0F, 40.0F, 125.0F, 145.0F}, PageSize.LETTER)

          If refOxxo.Trim <> "" And codigoBarraOxxo <> "" Then
            Dim imagenTiendas As iTextSharp.text.Image
            'imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/tiendasopen.jpg") 'nombre y ruta de la imagen a insertar
            imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/stores_3.jpg") 'nombre y ruta de la imagen a insertar
            'imagenTiendas.ScalePercent(44) 'escala al tamaño de la imagen openpay
            imagenTiendas.ScalePercent(50)
            Dim cellimgTiendas As New PdfPCell(imagenTiendas)
            cellimgTiendas.Border = PdfPCell.NO_BORDER
            cellimgTiendas.BorderWidthBottom = 0
            cellimgTiendas.PaddingTop = 1.0F
            cellimgTiendas.HorizontalAlignment = 1  ' 0 para open pay
            cellimgTiendas.Colspan = 5
            cellimgTiendas.BorderColorBottom = Color.WHITE
            tblPagoTiendas.AddCell(cellimgTiendas)

            'Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS PAYNET OPENPAY", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
            Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
            cellPagoOxxo.Border = PdfPCell.NO_BORDER
            cellPagoOxxo.BorderWidthBottom = 0
            cellPagoOxxo.PaddingTop = 10.0F
            cellPagoOxxo.HorizontalAlignment = 1
            cellPagoOxxo.Colspan = 5
            cellPagoOxxo.BorderColorBottom = Color.WHITE
            tblPagoTiendas.AddCell(cellPagoOxxo)

            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3
            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
            Dim imgOxxo As iTextSharp.text.Image 'declaración de imagen
            imgOxxo = iTextSharp.text.Image.GetInstance(codigoBarraOxxo) 'nombre y ruta de la imagen a insertar
            'imagen.ScalePercent(50) 'escala al tamaño de la imagen

            Dim cellimgOxxo As New PdfPCell(imgOxxo)
            cellimgOxxo.Border = PdfPCell.NO_BORDER
            cellimgOxxo.BorderWidthBottom = 0
            cellimgOxxo.PaddingTop = 5.0F
            cellimgOxxo.HorizontalAlignment = 1
            cellimgOxxo.Colspan = 5
            cellimgOxxo.BorderColorBottom = Color.WHITE
            tblPagoTiendas.AddCell(cellimgOxxo)

            Dim cellrefOxxo As New PdfPCell(New Phrase(refOxxo))
            cellrefOxxo.Border = PdfPCell.NO_BORDER
            cellrefOxxo.BorderWidthBottom = 0
            cellrefOxxo.PaddingTop = 5.0F
            cellrefOxxo.HorizontalAlignment = 1
            cellrefOxxo.Colspan = 5
            cellrefOxxo.BorderColorBottom = Color.WHITE
            tblPagoTiendas.AddCell(cellrefOxxo)
          End If

          tblPagoTiendas.AddCell(cellEspacio)
          tblPagoTiendas.AddCell(cellEspacio)
          tblPagoTiendas.AddCell(cellEspacio)

          Dim cellTiendas As New PdfPCell(New Phrase("TIENDAS PARA REALIZAR SU PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTiendas.Border = PdfPCell.NO_BORDER
          cellTiendas.BorderWidthBottom = 0
          cellTiendas.PaddingTop = 5.0F
          cellTiendas.HorizontalAlignment = 1
          cellTiendas.Colspan = 5
          cellTiendas.BorderColorBottom = Color.WHITE
          'tblPagoTiendas.AddCell(cellTiendas)

          Dim cellInstrucciones As New PdfPCell(New Phrase("INSTRUCCIONES PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInstrucciones.Border = PdfPCell.NO_BORDER
          cellInstrucciones.BorderWidthBottom = 0
          cellInstrucciones.PaddingTop = 2.0F
          cellInstrucciones.PaddingBottom = 5.0F
          cellInstrucciones.HorizontalAlignment = 1
          cellInstrucciones.Colspan = 5
          cellInstrucciones.BorderColorBottom = Color.WHITE

          tblPagoTiendas.AddCell(cellInstrucciones)

          Dim cellPasps As New PdfPCell(New Phrase("1.- DEBES ELEGIR LA TIENDA QUE MÁS TE CONVENGA ENTRE LAS CADENAS INDICADAS (SOLO SE PUEDE PAGAR EN ESAS TIENDAS).", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPagoTiendas.AddCell(cellPasps)

          Dim boldText As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          'Dim compania As Chunk = New Chunk("PAYNET OPENPAY", boldText)
          Dim compania As Chunk = New Chunk("CONEKTA", boldText)
          compania.SetUnderline(0.4, -0.8)
          Dim instruccion As String = "2.- AL ACERCARSE AL MOSTRADOR, DEBERÁ MENCIONAR QUE VIENE A PAGAR "
          Dim ph As Phrase = New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          ph.Add(compania)
          ph.Add(",Y MOSTRAR AL CAJERO EL CÓDIGO DE BARRAS O DICTAR LOS NÚMEROS QUE APARECEN EN LA REFERENCIA.")

          cellPasps = New PdfPCell(ph)
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPagoTiendas.AddCell(cellPasps)

          instruccion = "3.- UNA VEZ REALIZADO EL PAGO EN EFECTIVO, ENVIAREMOS UNA NOTIFICACIÓN DE PAGO EN TIEMPO REAL A SU CORREO Y ¡LISTO!"
          cellPasps = New PdfPCell(New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPagoTiendas.AddCell(cellPasps)



          'tblPagoTiendas.AddCell(cellNota2)
          Dim cellGracias As New PdfPCell(New Phrase("¡MUCHAS GRACIAS POR DARNOS LA OPORTUNIDAD DE SERVIRLE!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellGracias.Border = PdfPCell.NO_BORDER
          cellGracias.BorderWidthBottom = 0
          cellGracias.PaddingTop = 20.0F
          cellGracias.PaddingBottom = 30.0F
          cellGracias.HorizontalAlignment = 1
          cellGracias.Colspan = 5
          cellGracias.BorderColorBottom = Color.WHITE

          'tblPagoTiendas.AddCell(cellGracias)
          tblPagoTiendas.AddCell(cellEspacio)
          tblPagoTiendas.AddCell(cellEspacio)
          ' tblPagoTiendas.AddCell(cellPeriodo3)

          Dim cellPie1 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie1.Border = PdfPCell.NO_BORDER
          cellPie1.BorderWidthBottom = 0
          cellPie1.PaddingTop = 2.0F
          cellPie1.HorizontalAlignment = 0
          cellPie1.Colspan = 4
          cellPie1.BorderColorBottom = Color.WHITE

          'tblPagoTiendas.AddCell(cellPie1)
          imagenInfo.ScalePercent(40)
          Dim cellPie2 As New PdfPCell(imagenInfo)
          cellPie2.Border = PdfPCell.NO_BORDER
          cellPie2.BorderWidthBottom = 0
          cellPie2.PaddingTop = 2.0F
          cellPie2.HorizontalAlignment = 2
          cellPie2.Colspan = 1
          cellPie2.BorderColorBottom = Color.WHITE

          'tblPagoTiendas.AddCell(cellPie2)
          oDoc.Add(tblPagoTiendas)

          'Fin del flujo de bytes.
          cb.EndText()
          'Forzamos vaciamiento del buffer.
          pdfw.Flush()
          'Cerramos el documento.
          oDoc.Close()
        Catch ex As Exception
          'Si hubo una excepcion y el archivo existe …
          If File.Exists(ruta) Then
            'Cerramos el documento si esta abierto.
            'Y asi desbloqueamos el archivo para su eliminacion.
            If oDoc.IsOpen Then oDoc.Close()
            '… lo eliminamos de disco.
            File.Delete(ruta)
          End If
          'Throw New Exception("Error al generar archivo PDF (" & ex.Message & ")" & ex.Source)
          MsgBox(ex.Message & "--- " & ex.StackTrace)
          'Dim sqlerror As String = "insert into"
          'Dim sql As String = "insert into netcel..Correos(cliente,mensaje,asunto,estatus,respuesta) values('-1','ERROR AL GENERAR ESTADO DE CUENTA DE COMUNICALO  " & cli_id.ToString & ", MENSAJE:" & ex.Message & "<br/> SOURCE: " & ex.Source & " <br/> STACK TRACE:" & ex.StackTrace & "','ERROR ESTADO DE CUENTA ILOXTELECOM" & cli_id.ToString & "','1','sinfante@mail.ilox.mx')"
          'con.ModRegEli(sql)
          'escribir_log("ERROR AL GENERAR ESTADO DE CUENTA DEL CONTRATO_ID " & idcliente.ToString & ", MENSAJE:" & ex.Message & " SOURCE: " & ex.Source & " STACK TRACE:" & ex.StackTrace)
        Finally
          cb = Nothing
          pdfw = Nothing
          oDoc = Nothing
        End Try
      End If
    End If
  End Sub



  Private Sub Generar_pdfOXXO(ByVal id_estado_cuenta As Integer, ByVal id_contrato As Integer, ByVal path As String, ByVal refOxxo As String, ByVal codigoBarraOxxo As String)
    Dim sqledo As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtedo As DataTable = con.ConsultarDT(sqledo)
    If dtedo IsNot Nothing AndAlso dtedo.Rows.Count > 0 Then
      Dim fecha As Date = dtedo(0)("fecha").ToString
      Dim grantotal As Double = Val(dtedo(0)("grantotal").ToString)
      Dim saldo_pendiente As Double = Val(dtedo(0)("saldo_pendiente").ToString)
      Dim total_edo As Double = grantotal - saldo_pendiente
      Dim periodoA As Date = dtedo(0)("periodoA").ToString
      Dim periodoB As Date = dtedo(0)("periodoB").ToString
      Dim totalPlan As Double = Val(dtedo(0)("mensualidad").ToString())
      Dim sqlcli As String = ""
      Dim sqlBalance = "select * from CONTRACTS_BALANCES where id_contrato=" & id_contrato & ";"
      Dim dtBalance = con.ConsultarDT(sqlBalance)
      Dim balance As Double = 0

      If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
        balance = Val(dtBalance(0)("balance").ToString)
      End If

      If tiene_telefonia(id_contrato) Then
        sqlcli = $"SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero,t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM (" &
" SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete FROM(" &
" SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre) AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias" &
" FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e" &
" ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato &
" ) AS t1 INNER JOIN CALLES ca ON ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete)" &
" AS t3 INNER JOIN dbo.EQUIPOS e INNER JOIN EQUIPOS_TELEFONIA et INNER JOIN LINEAS l" &
" ON l.id_linea=et.id_linea ON et.id_equipo=e.id_equipo ON e.id_contrato=t3.id_contrato  where e.estatus=1 AND et.estatus=1"
      Else
        sqlcli = "SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero," &
" t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM" &
" ( SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete,numero" &
" FROM( SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre)" &
 " AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias,telefono AS numero FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c " &
 " INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON" &
 " col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato & " ) AS t1 INNER JOIN CALLES ca ON" &
 " ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete) AS t3"
      End If

      Dim dtcli As DataTable = con.ConsultarDT(sqlcli)
      If dtcli IsNot Nothing AndAlso dtcli.Rows.Count > 0 Then
        Dim nombre As String = dtcli(0)("nombre").ToString
        Dim contrato As String = dtcli(0)("contrato").ToString
        Dim contract As String = dtcli(0)("contrato").ToString
        Dim calle As String = dtcli(0)("calle").ToString
        Dim numext As String = dtcli(0)("numext").ToString
        Dim numint As String = dtcli(0)("numint").ToString
        Dim colonia As String = dtcli(0)("colonia").ToString
        Dim cp As String = dtcli(0)("cp").ToString
        Dim municipio As String = dtcli(0)("municipio").ToString
        Dim estado As String = dtcli(0)("estado").ToString
        Dim referencias As String = dtcli(0)("referencias").ToString
        Dim paquete As String = dtcli(0)("paquete").ToString
        Dim numero As String = dtcli(0)("numero").ToString
        Dim phone As String = dtcli(0)("numero").ToString
        Dim id_paquete As Integer = Val(dtcli(0)("id_paquete").ToString)
        Dim servicios As String = getServicios(id_paquete)
        Dim dtAccount As DataTable = getStpAccount(id_contrato)
        Dim account As String = "s/a"

        If dtAccount IsNot Nothing AndAlso dtAccount.Rows.Count > 0 Then
          account = dtAccount.Rows(0)("clave").ToString()
        End If

        Dim ruta As String = path & "\EstadoCuenta(" & id_estado_cuenta.ToString & ").pdf "
        Dim oDoc As New iTextSharp.text.Document(PageSize.LETTER, 50, 50, 50, 50)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim linea As PdfContentByte
        Dim rectangulo As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Try
          pdfw = PdfWriter.GetInstance(oDoc, New FileStream(ruta,
                    FileMode.Create, FileAccess.Write, FileShare.None))

          Me.PageState = New CustomPageState()
          ''//Wire our event handler and pass in the page state
          pdfw.PageEvent = New MyCustomPdfEvent(Me.PageState)
          'Apertura del documento.
          oDoc.Open()
          cb = pdfw.DirectContent
          linea = pdfw.DirectContent
          rectangulo = pdfw.DirectContent

          'Agregamos una pagina. // check later
          'oDoc.NewPage()

          cb.BeginText()
          fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
          cb.SetFontAndSize(fuente, 10) 'fuente definida en la linea anterior y tamaño

          Dim f10 As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f10.SetColor(2, 51, 130)

          Dim f10Bold As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10Bold.SetColor(2, 51, 130)

          Dim f10BoldMain As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10BoldMain.SetColor(2, 51, 130)


          Dim f14 As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f14.SetColor(2, 51, 130)

          Dim tblBanner As New PdfPTable(1)
          tblBanner.HorizontalAlignment = 0
          tblBanner.LockedWidth = True
          tblBanner.TotalWidth = 550.0F
          tblBanner.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBanner.DefaultCell.MinimumHeight = 12
          tblBanner.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER
          tblBanner.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          'tblBanner.SetWidthPercentage({140.0F, 100.0F, 300.0F}, PageSize.LETTER)

          'Dim banner As iTextSharp.text.Image
          'banner = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/banner.jpg") 'nombre y ruta de la imagen a insertar
          'imagen.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)
          'tblBanner.AddCell(banner)
          'oDoc.Add(tblBanner)
          oDoc.Add(New Paragraph(" "))

          Dim tblHeaderInfo As New PdfPTable(3)
          tblHeaderInfo.HorizontalAlignment = 0
          tblHeaderInfo.LockedWidth = True
          tblHeaderInfo.TotalWidth = 540.0F
          tblHeaderInfo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblHeaderInfo.DefaultCell.MinimumHeight = 12
          tblHeaderInfo.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeaderInfo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblHeaderInfo.SetWidthPercentage({220.0F, 50.0F, 270.0F}, PageSize.LETTER)

          'IMAGEN
          Dim imagenInfo As iTextSharp.text.Image
          imagenInfo = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
          imagenInfo.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)
          'tblHeaderInfo.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))

          Dim cellInfoCompany As New PdfPTable(1)
          cellInfoCompany.DefaultCell.Border = PdfPCell.NO_BORDER
          cellInfoCompany.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT
          ' Comunicalo info.
          cellInfoCompany.AddCell(imagenInfo)
          cellInfoCompany.AddCell(New Phrase("Comunícalo de México S.A. de C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("CONVENTO DE CHURUBUSCO NO. 4,", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("COL. JARDINES DE SANTA MÓNICA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("MPIO. TLALNEPANTLA DE BAZ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("ESTADO DE MÉXICO", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("C.P. 54050", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("RFC: CME0806162SA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(nombre, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          ' Contract info.
          cellInfoCompany.AddCell(New Phrase(calle, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(referencias & " " & numext & " " & numint, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(colonia, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(municipio & ", " & estado & ", C.P. " & cp, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'cellInfoCompany.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoCompany.AddCell(New Phrase("SERVICIOS CONTRATADOS", f10Bold))

          Dim nesthousingInfo As New PdfPCell(cellInfoCompany)
          nesthousingInfo.Border = PdfPCell.NO_BORDER
          nesthousingInfo.Padding = 0F
          nesthousingInfo.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeaderInfo.AddCell(nesthousingInfo)

          ' Header right column , information about bill.
          Dim cellBillTitle As New PdfPCell(New Phrase("ESTADO DE CUENTA", f10BoldMain))
          cellBillTitle.Border = PdfPCell.BOTTOM_BORDER
          cellBillTitle.BorderWidthBottom = 4
          cellBillTitle.PaddingTop = 12.0F
          cellBillTitle.PaddingBottom = 15.0F
          cellBillTitle.HorizontalAlignment = 1
          cellBillTitle.Colspan = 1
          cellBillTitle.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          Dim tblBillTitle As New PdfPTable(1)
          tblBillTitle.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBillTitle.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblBillTitle.AddCell(cellBillTitle)
          tblBillTitle.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim tblBillItems As New PdfPTable(2)
          tblBillItems.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBillItems.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT
          tblBillItems.DefaultCell.PaddingLeft = 7.0F
          tblBillItems.DefaultCell.PaddingRight = 7.0F
          tblBillItems.DefaultCell.PaddingBottom = 3.0F

          tblBillItems.AddCell(New Phrase("MES DE FACTURACIÓN", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(MonthName(periodoA.Month).ToUpper, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase("CONTRATO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(contract, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblBillItems.AddCell(New Phrase("TELÉFONO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(phone, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase("TOTAL A PAGAR", f10Bold))
          tblBillItems.AddCell(New Phrase(FormatCurrency(grantotal, 2), f10Bold))
          tblBillItems.AddCell(New Phrase("FECHA LIMITE DE PAGO", f10Bold))
          tblBillItems.AddCell(New Phrase(periodoA.ToString("dd/MM/yyyy"), f10Bold))
          tblBillItems.AddCell(New Phrase("SALDO VENCIDO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblBillItems.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim cellAdvisory As New PdfPCell()
          cellAdvisory.PaddingTop = 12.0F
          cellAdvisory.PaddingBottom = 15.0F
          cellAdvisory.HorizontalAlignment = 1
          cellAdvisory.Colspan = 2
          cellAdvisory.BorderWidth = 0

          Dim imagenAdvisory As iTextSharp.text.Image
          imagenAdvisory = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/img_advisory.jpg")
          'imagenAdvisory.ScalePercent(90)
          cellAdvisory.AddElement(imagenAdvisory)
          tblBillItems.AddCell(cellAdvisory)
          tblBillTitle.AddCell(tblBillItems)
          tblHeaderInfo.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))
          tblHeaderInfo.AddCell(tblBillTitle)

          Dim cellSeparator As New PdfPCell(New Phrase("", f10))
          cellSeparator.Border = PdfPCell.BOTTOM_BORDER
          cellSeparator.BorderWidthBottom = 2
          cellSeparator.PaddingTop = 1.0F
          cellSeparator.PaddingBottom = 1.0F
          cellSeparator.HorizontalAlignment = 1
          cellSeparator.Colspan = 3
          cellSeparator.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))
          tblHeaderInfo.AddCell(cellSeparator)

          oDoc.Add(tblHeaderInfo)
          'oDoc.Add(New Paragraph(" "))

          Dim cellEspacio As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellEspacio.Border = PdfPCell.NO_BORDER
          cellEspacio.BorderWidthBottom = 0
          cellEspacio.PaddingTop = 5.0F
          cellEspacio.HorizontalAlignment = 1
          cellEspacio.Colspan = 5
          cellEspacio.BorderColorBottom = Color.WHITE

          Dim tblPeriodo As New PdfPTable(5)
          tblPeriodo.HorizontalAlignment = 0
          tblPeriodo.LockedWidth = True
          tblPeriodo.TotalWidth = 540.0F
          tblPeriodo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblPeriodo.DefaultCell.MinimumHeight = 12
          tblPeriodo.DefaultCell.HorizontalAlignment = 0
          tblPeriodo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblPeriodo.DefaultCell.PaddingLeft = 12.0F
          tblPeriodo.SetWidthPercentage({150.0F, 80.0F, 40.0F, 125.0F, 145.0F}, PageSize.LETTER)

          Dim cellPeriodo3 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPeriodo3.Border = PdfPCell.BOTTOM_BORDER
          cellPeriodo3.BorderWidthBottom = 2
          cellPeriodo3.PaddingTop = 0
          cellPeriodo3.HorizontalAlignment = 0
          cellPeriodo3.Colspan = 5
          cellPeriodo3.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          Dim cellPaqueteContratado As New PdfPCell(New Phrase("Cargos del mes", f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 5.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPaqueteContratado)
          cellPaqueteContratado = New PdfPCell(New Phrase(paquete, f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 1.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPaqueteContratado)

          Dim cellServicios As New PdfPCell(New Phrase(servicios, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServicios.Border = PdfPCell.NO_BORDER
          cellServicios.BorderWidth = 0
          cellServicios.PaddingTop = 0
          cellServicios.HorizontalAlignment = 0
          cellServicios.Colspan = 4
          cellServicios.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellServicios)
          tblPeriodo.AddCell(New Phrase(FormatCurrency(totalPlan, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          ' Before "saldo pendiente"
          Dim cellPending As New PdfPCell(New Phrase("SALDO PENDIENTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPending.Border = PdfPCell.NO_BORDER
          cellPending.BorderWidth = 0
          cellPending.PaddingTop = 0
          cellPending.HorizontalAlignment = 0
          cellPending.Colspan = 4
          cellPending.BorderColor = Color.WHITE

          'tblPeriodo.AddCell(cellPending)
          'tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          'Show balance'
          Dim auxBalance As Double = If(balance < 0, balance * -1, balance)
          Dim showBalance As Boolean = If(balance <= 0, True, False)

          'If balance < 0 Then
          '  auxBalance = balance * -1
          'End If

          If showBalance Then
            Dim auxShowBalance As Double = auxBalance + grantotal
            Dim cellShowBalance As New PdfPCell(New Phrase("SALDO A FAVOR", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellShowBalance.Border = PdfPCell.NO_BORDER
            cellShowBalance.BorderWidth = 0
            cellShowBalance.PaddingTop = 0
            cellShowBalance.HorizontalAlignment = 0
            cellShowBalance.Colspan = 4
            cellShowBalance.BorderColor = Color.WHITE

            'tblPeriodo.AddCell(cellShowBalance)
            'tblPeriodo.AddCell(New Phrase(FormatCurrency(auxShowBalance, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          End If

          Dim dtCharges As DataTable = getDataBillCharges(id_estado_cuenta)

          If dtCharges IsNot Nothing AndAlso dtCharges.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("OTROS CARGOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtCharges.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtCharges.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtCharges.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim dtDiscount As DataTable = getDataBillDiscounts(id_estado_cuenta)

          If dtDiscount IsNot Nothing AndAlso dtDiscount.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("DESCUENTOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtDiscount.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtDiscount.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtDiscount.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim celltotal As New PdfPCell(New Phrase("TOTAL A PAGAR " & FormatCurrency(grantotal, 2), f14))
          celltotal.Border = PdfPCell.NO_BORDER
          celltotal.BorderWidth = 0
          celltotal.PaddingTop = 10.0F
          celltotal.PaddingLeft = 12.0F
          celltotal.HorizontalAlignment = 0
          celltotal.Colspan = 2
          celltotal.BorderColor = Color.WHITE

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(celltotal)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          Dim celltotalLetra As New PdfPCell(New Phrase("(" & totalLetra(grantotal) & ")", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          celltotalLetra.Border = PdfPCell.NO_BORDER
          celltotalLetra.BorderWidth = 0
          celltotalLetra.PaddingTop = 0
          celltotalLetra.PaddingLeft = 12.0F
          celltotalLetra.HorizontalAlignment = 0
          celltotalLetra.Colspan = 2
          celltotalLetra.BorderColor = Color.WHITE

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(celltotalLetra)

          Dim advisoryPlan As String = $"*En el caso de haber realizado un cambio o actualización en su paquete, al realizar el pago de este Estado de Cuenta, usted acepta los nuevos términos y Condiciones aplicables."
          Dim cellAdvisoryPlan As New PdfPCell(New Phrase(advisoryPlan, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellAdvisoryPlan.Border = PdfPCell.NO_BORDER
          cellAdvisoryPlan.BorderWidth = 0
          cellAdvisoryPlan.PaddingTop = 0
          cellAdvisoryPlan.PaddingLeft = 0
          cellAdvisoryPlan.HorizontalAlignment = 0
          cellAdvisoryPlan.Colspan = 5
          cellAdvisoryPlan.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellAdvisoryPlan)
          tblPeriodo.AddCell(cellPeriodo3)
          tblPeriodo.AddCell(cellEspacio)

          'Balance'
          If showBalance Then
            Dim cellBalance As New PdfPCell(New Phrase("SALDO A FAVOR RESTANTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellBalance.Border = PdfPCell.NO_BORDER
            cellBalance.BorderWidth = 0
            cellBalance.PaddingTop = 0
            cellBalance.HorizontalAlignment = 0
            cellBalance.Colspan = 4
            cellBalance.BorderColor = Color.WHITE

            'tblPeriodo.AddCell(cellBalance)
            'tblPeriodo.AddCell(New Phrase(FormatCurrency(auxBalance, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            'tblPeriodo.AddCell(cellPeriodo3)
          End If

          ' Late fee.
          Dim lateFeeTitle As Font = New Font(iTextSharp.text.Font.HELVETICA, 11.0F, iTextSharp.text.Font.BOLD, New Color(255, 204, 0))
          Dim lateFeeText As Font = New Font(iTextSharp.text.Font.HELVETICA, 9.0F, iTextSharp.text.Font.BOLD, Color.WHITE)

          Dim phLateFee As New Phrase()
          phLateFee.Add(New Chunk("——  IMPORTANTE - PAGO TARDÍO  ——" & Environment.NewLine & Environment.NewLine, lateFeeTitle))

          phLateFee.Add(New Chunk(
"A partir del mes de AGOSTO DE 2026, los pagos realizados después de la fecha límite establecida generarán un cargo administrativo por pago tardío de $50.00 (Cincuenta pesos 00/100 M.N.), el cual se aplicará de forma inmediata." &
Environment.NewLine & Environment.NewLine &
"En caso de suspensión del servicio por falta de pago, este cargo deberá liquidarse junto con la mensualidad vencida para la reactivación del servicio." &
Environment.NewLine &
"Le invitamos a pagar dentro de la fecha establecida para evitar cargos adicionales.",
lateFeeText))

          Dim cellLateFee As New PdfPCell(phLateFee)
          With cellLateFee
            .BackgroundColor = New Color(System.Drawing.ColorTranslator.FromHtml("#08154D")) 'azul oscuro
            .Border = PdfPCell.NO_BORDER
            .Colspan = 5
            .HorizontalAlignment = Element.ALIGN_CENTER
            .VerticalAlignment = Element.ALIGN_MIDDLE
            .PaddingTop = 12
            .PaddingBottom = 12
            .PaddingLeft = 18
            .PaddingRight = 18
          End With

          tblPeriodo.AddCell(cellLateFee)
          tblPeriodo.AddCell(cellEspacio)

          Dim imgWarning As iTextSharp.text.Image
          imgWarning = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/advisory_2.jpg")
          imgWarning.ScalePercent(50.0F)
          cellAdvisoryPlan = New PdfPCell(New Phrase(advisoryPlan, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellAdvisoryPlan.Border = PdfPCell.NO_BORDER
          cellAdvisoryPlan.BorderWidth = 0
          cellAdvisoryPlan.PaddingTop = 2.0F
          cellAdvisoryPlan.PaddingLeft = 0
          cellAdvisoryPlan.HorizontalAlignment = 0
          cellAdvisoryPlan.Colspan = 5
          cellAdvisoryPlan.BorderColor = Color.WHITE
          cellAdvisoryPlan.AddElement(imgWarning)
          tblPeriodo.AddCell(cellAdvisoryPlan)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)

          ' Warning payments
          Dim phWarning As Phrase = New Phrase("IMPORTANTE SOBRE SU PAGO.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.RED))
          Dim cellPaymentWarning As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellPaymentWarning = New PdfPCell(phWarning)
          cellPaymentWarning.Border = PdfPCell.NO_BORDER
          cellPaymentWarning.BorderWidthBottom = 0
          cellPaymentWarning.PaddingTop = 1.0F
          cellPaymentWarning.HorizontalAlignment = 0
          cellPaymentWarning.Colspan = 5
          cellPaymentWarning.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellPaymentWarning)

          Dim warningContent As String = "Cada contrato tiene su propia CLABE interbancaria. No utilice la CLABE de un contrato para pagar otro diferente.Los pagos no pueden transferirse entre contratos y quedarán como saldo a favor del contrato asociado a la CLABE utilizada."
          Dim phWarningContent As Phrase = New Phrase(warningContent)

          Dim cellPaymentWarningContent As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.RED)))
          cellPaymentWarningContent = New PdfPCell(phWarningContent)
          cellPaymentWarningContent.Border = PdfPCell.NO_BORDER
          cellPaymentWarningContent.BorderWidthBottom = 0
          cellPaymentWarningContent.PaddingTop = 1.0F
          cellPaymentWarningContent.HorizontalAlignment = 0
          cellPaymentWarningContent.Colspan = 5
          cellPaymentWarningContent.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellPaymentWarningContent)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)

          Dim cellFormasPago As New PdfPCell(New Phrase("FORMAS DE PAGO", f10Bold))
          cellFormasPago.Border = PdfPCell.NO_BORDER
          cellFormasPago.BorderWidthBottom = 0
          cellFormasPago.PaddingTop = 2.0F
          cellFormasPago.HorizontalAlignment = 0
          cellFormasPago.Colspan = 5
          cellFormasPago.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellFormasPago)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)

          Dim cellStp As New PdfPCell(New Phrase("ATENCIÓN", New Font(iTextSharp.text.Font.HELVETICA, 11.0F, iTextSharp.text.Font.BOLD, Color.RED)))
          cellStp.Border = PdfPCell.NO_BORDER
          cellStp.BorderWidthBottom = 0
          cellStp.PaddingTop = 10.0F
          cellStp.HorizontalAlignment = 0
          cellStp.Colspan = 5
          'tblPeriodo.AddCell(cellStp)

          Dim instructions As String = "A PARTIR DE AHORA CADA CLIENTE TENDRÁ UNA CLABE INTERBANCARIA ÚNICA Y PERSONALIZADA POR CONTRATO."
          Dim cellStpInstructions As New PdfPCell(New Phrase(instructions, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellStpInstructions.Border = PdfPCell.NO_BORDER
          cellStpInstructions.BorderWidthBottom = 0
          cellStpInstructions.PaddingTop = 10.0F
          cellStpInstructions.HorizontalAlignment = 0
          cellStpInstructions.Colspan = 5
          cellStpInstructions.BorderColorBottom = Color.BLACK
          'tblPeriodo.AddCell(cellStpInstructions)

          Dim cellDeposito As New PdfPCell(New Phrase("DATOS PARA TRANSFERENCIA:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellDeposito.Border = PdfPCell.NO_BORDER
          cellDeposito.BorderWidthBottom = 0
          cellDeposito.PaddingTop = 1.0F
          cellDeposito.HorizontalAlignment = 0
          cellDeposito.Colspan = 5
          cellDeposito.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellDeposito)

          Dim cellTransfer As New PdfPCell(New Phrase("TRANSFERENCIA ELECTRÓNICA:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTransfer.Border = PdfPCell.NO_BORDER
          cellTransfer.BorderWidthBottom = 0
          cellTransfer.PaddingTop = 10.0F
          cellTransfer.HorizontalAlignment = 1
          cellTransfer.Colspan = 2
          cellTransfer.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellTransfer)

          Dim cellFormasPago2 As New PdfPCell(New Phrase("BANCO: STP", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago2.Border = PdfPCell.NO_BORDER
          cellFormasPago2.BorderWidthBottom = 0
          cellFormasPago2.PaddingTop = 2.0F
          cellFormasPago2.HorizontalAlignment = 0
          cellFormasPago2.Colspan = 5
          cellFormasPago2.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago2)

          Dim cellClabe As New PdfPCell(New Phrase("CLABE INTERBANCARIA: 044180256007653656", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellClabe.Border = PdfPCell.NO_BORDER
          cellClabe.BorderWidthBottom = 0
          cellClabe.PaddingTop = 2.0F
          cellClabe.HorizontalAlignment = 2
          cellClabe.Colspan = 2
          cellClabe.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellClabe)

          Dim cellFormasPago3 As New PdfPCell(New Phrase("BENEFICIARIO: COMUNICALO DE MÉXICO, S.A. DE C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago3.Border = PdfPCell.NO_BORDER
          cellFormasPago3.BorderWidthBottom = 0
          cellFormasPago3.PaddingTop = 0.0F
          cellFormasPago3.HorizontalAlignment = 0
          cellFormasPago3.Colspan = 5
          cellFormasPago3.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago3)
          Dim boldTextAccount As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          Dim clabeText As Chunk = New Chunk(account, boldTextAccount)
          Dim indClabe As String = "CLABE INTERBANCARIA PERSONALIZADA: "
          Dim phClabe As Phrase = New Phrase(indClabe, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          phClabe.Add(clabeText)

          Dim cellFormasPago4 As New PdfPCell(New Phrase("CLABE INTERBANCARIA PERSONALIZADA: " + Environment.NewLine + account, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago4 = New PdfPCell(phClabe)
          cellFormasPago4.Border = PdfPCell.NO_BORDER
          cellFormasPago4.BorderWidthBottom = 0
          cellFormasPago4.PaddingTop = 0.0F
          cellFormasPago4.HorizontalAlignment = 0
          cellFormasPago4.Colspan = 5
          cellFormasPago4.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago4)
          'tblPeriodo.AddCell(phClabe)
          tblPeriodo.AddCell(cellEspacio)

          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          If refOxxo.Trim <> "" And codigoBarraOxxo <> "" Then
            Dim imagenTiendas As iTextSharp.text.Image
            'imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/tiendasopen.jpg") 'nombre y ruta de la imagen a insertar
            imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/stores_3.jpg") 'nombre y ruta de la imagen a insertar
            'imagenTiendas.ScalePercent(44) 'escala al tamaño de la imagen openpay
            imagenTiendas.ScalePercent(50)
            Dim cellimgTiendas As New PdfPCell(imagenTiendas)
            cellimgTiendas.Border = PdfPCell.NO_BORDER
            cellimgTiendas.BorderWidthBottom = 0
            cellimgTiendas.PaddingTop = 1.0F
            cellimgTiendas.HorizontalAlignment = 1  ' 0 para open pay
            cellimgTiendas.Colspan = 5
            cellimgTiendas.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellimgTiendas)

            'Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS PAYNET OPENPAY", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
            Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
            cellPagoOxxo.Border = PdfPCell.NO_BORDER
            cellPagoOxxo.BorderWidthBottom = 0
            cellPagoOxxo.PaddingTop = 10.0F
            cellPagoOxxo.HorizontalAlignment = 1
            cellPagoOxxo.Colspan = 5
            cellPagoOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellPagoOxxo)

            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3
            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
            Dim imgOxxo As iTextSharp.text.Image 'declaración de imagen
            imgOxxo = iTextSharp.text.Image.GetInstance(codigoBarraOxxo) 'nombre y ruta de la imagen a insertar
            'imagen.ScalePercent(50) 'escala al tamaño de la imagen

            Dim cellimgOxxo As New PdfPCell(imgOxxo)
            cellimgOxxo.Border = PdfPCell.NO_BORDER
            cellimgOxxo.BorderWidthBottom = 0
            cellimgOxxo.PaddingTop = 5.0F
            cellimgOxxo.HorizontalAlignment = 1
            cellimgOxxo.Colspan = 5
            cellimgOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellimgOxxo)

            Dim cellrefOxxo As New PdfPCell(New Phrase(refOxxo))
            cellrefOxxo.Border = PdfPCell.NO_BORDER
            cellrefOxxo.BorderWidthBottom = 0
            cellrefOxxo.PaddingTop = 5.0F
            cellrefOxxo.HorizontalAlignment = 1
            cellrefOxxo.Colspan = 5
            cellrefOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellrefOxxo)
          End If

          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          Dim cellTiendas As New PdfPCell(New Phrase("TIENDAS PARA REALIZAR SU PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTiendas.Border = PdfPCell.NO_BORDER
          cellTiendas.BorderWidthBottom = 0
          cellTiendas.PaddingTop = 5.0F
          cellTiendas.HorizontalAlignment = 1
          cellTiendas.Colspan = 5
          cellTiendas.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellTiendas)

          Dim cellInstrucciones As New PdfPCell(New Phrase("INSTRUCCIONES PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInstrucciones.Border = PdfPCell.NO_BORDER
          cellInstrucciones.BorderWidthBottom = 0
          cellInstrucciones.PaddingTop = 2.0F
          cellInstrucciones.PaddingBottom = 5.0F
          cellInstrucciones.HorizontalAlignment = 1
          cellInstrucciones.Colspan = 5
          cellInstrucciones.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellInstrucciones)

          Dim cellPasps As New PdfPCell(New Phrase("1.- DEBES ELEGIR LA TIENDA QUE MÁS TE CONVENGA ENTRE LAS CADENAS INDICADAS (SOLO SE PUEDE PAGAR EN ESAS TIENDAS).", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          Dim boldText As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          'Dim compania As Chunk = New Chunk("PAYNET OPENPAY", boldText)
          Dim compania As Chunk = New Chunk("CONEKTA", boldText)
          compania.SetUnderline(0.4, -0.8)
          Dim instruccion As String = "2.- AL ACERCARSE AL MOSTRADOR, DEBERÁ MENCIONAR QUE VIENE A PAGAR "
          Dim ph As Phrase = New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          ph.Add(compania)
          ph.Add(",Y MOSTRAR AL CAJERO EL CÓDIGO DE BARRAS O DICTAR LOS NÚMEROS QUE APARECEN EN LA REFERENCIA.")

          cellPasps = New PdfPCell(ph)
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          instruccion = "3.- UNA VEZ REALIZADO EL PAGO EN EFECTIVO, ENVIAREMOS UNA NOTIFICACIÓN DE PAGO EN TIEMPO REAL A SU CORREO Y ¡LISTO!"
          cellPasps = New PdfPCell(New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          'tblPeriodo.AddCell(cellNota2)
          Dim cellGracias As New PdfPCell(New Phrase("¡MUCHAS GRACIAS POR DARNOS LA OPORTUNIDAD DE SERVIRLE!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellGracias.Border = PdfPCell.NO_BORDER
          cellGracias.BorderWidthBottom = 0
          cellGracias.PaddingTop = 20.0F
          cellGracias.PaddingBottom = 30.0F
          cellGracias.HorizontalAlignment = 1
          cellGracias.Colspan = 5
          cellGracias.BorderColorBottom = Color.WHITE

          'tblPeriodo.AddCell(cellGracias)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          ' tblPeriodo.AddCell(cellPeriodo3)

          Dim cellPie1 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie1.Border = PdfPCell.NO_BORDER
          cellPie1.BorderWidthBottom = 0
          cellPie1.PaddingTop = 2.0F
          cellPie1.HorizontalAlignment = 0
          cellPie1.Colspan = 4
          cellPie1.BorderColorBottom = Color.WHITE

          'tblPeriodo.AddCell(cellPie1)
          imagenInfo.ScalePercent(40)
          Dim cellPie2 As New PdfPCell(imagenInfo)
          cellPie2.Border = PdfPCell.NO_BORDER
          cellPie2.BorderWidthBottom = 0
          cellPie2.PaddingTop = 2.0F
          cellPie2.HorizontalAlignment = 2
          cellPie2.Colspan = 1
          cellPie2.BorderColorBottom = Color.WHITE

          'tblPeriodo.AddCell(cellPie2)
          oDoc.Add(tblPeriodo)

          'Fin del flujo de bytes.
          cb.EndText()
          'Forzamos vaciamiento del buffer.
          pdfw.Flush()
          'Cerramos el documento.
          oDoc.Close()
        Catch ex As Exception
          'Si hubo una excepcion y el archivo existe …
          If File.Exists(ruta) Then
            'Cerramos el documento si esta abierto.
            'Y asi desbloqueamos el archivo para su eliminacion.
            If oDoc.IsOpen Then oDoc.Close()
            '… lo eliminamos de disco.
            File.Delete(ruta)
          End If
          'Throw New Exception("Error al generar archivo PDF (" & ex.Message & ")" & ex.Source)
          MsgBox(ex.Message & "--- " & ex.StackTrace)
          'Dim sqlerror As String = "insert into"
          'Dim sql As String = "insert into netcel..Correos(cliente,mensaje,asunto,estatus,respuesta) values('-1','ERROR AL GENERAR ESTADO DE CUENTA DE COMUNICALO  " & cli_id.ToString & ", MENSAJE:" & ex.Message & "<br/> SOURCE: " & ex.Source & " <br/> STACK TRACE:" & ex.StackTrace & "','ERROR ESTADO DE CUENTA ILOXTELECOM" & cli_id.ToString & "','1','sinfante@mail.ilox.mx')"
          'con.ModRegEli(sql)
          'escribir_log("ERROR AL GENERAR ESTADO DE CUENTA DEL CONTRATO_ID " & idcliente.ToString & ", MENSAJE:" & ex.Message & " SOURCE: " & ex.Source & " STACK TRACE:" & ex.StackTrace)
        Finally
          cb = Nothing
          pdfw = Nothing
          oDoc = Nothing
        End Try
      End If
    End If
  End Sub

  Private Sub Generar_pdfOXXO_Before(ByVal id_estado_cuenta As Integer, ByVal id_contrato As Integer, ByVal path As String, ByVal refOxxo As String, ByVal codigoBarraOxxo As String)
    Dim sqledo As String = "select * from ESTADOS_CUENTA where id_estado_cuenta=" & id_estado_cuenta
    Dim dtedo As DataTable = con.ConsultarDT(sqledo)
    If dtedo IsNot Nothing AndAlso dtedo.Rows.Count > 0 Then
      Dim fecha As Date = dtedo(0)("fecha").ToString
      Dim grantotal As Double = Val(dtedo(0)("grantotal").ToString)
      Dim saldo_pendiente As Double = Val(dtedo(0)("saldo_pendiente").ToString)
      Dim total_edo As Double = grantotal - saldo_pendiente
      Dim periodoA As Date = dtedo(0)("periodoA").ToString
      Dim periodoB As Date = dtedo(0)("periodoB").ToString
      Dim totalPlan As Double = Val(dtedo(0)("mensualidad").ToString())
      Dim sqlcli As String = ""
      Dim sqlBalance = "select * from CONTRACTS_BALANCES where id_contrato=" & id_contrato & ";"
      Dim dtBalance = con.ConsultarDT(sqlBalance)
      Dim balance As Double = 0

      If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
        balance = Val(dtBalance(0)("balance").ToString)
      End If

      If tiene_telefonia(id_contrato) Then
        sqlcli = $"SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero,t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM (" &
" SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete FROM(" &
" SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre) AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias" &
" FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e" &
" ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato &
" ) AS t1 INNER JOIN CALLES ca ON ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete)" &
" AS t3 INNER JOIN dbo.EQUIPOS e INNER JOIN EQUIPOS_TELEFONIA et INNER JOIN LINEAS l" &
" ON l.id_linea=et.id_linea ON et.id_equipo=e.id_equipo ON e.id_contrato=t3.id_contrato  where e.estatus=1 AND et.estatus=1"
      Else
        sqlcli = "SELECT upper(nombre) AS nombre,contrato,calle,numext,numint,colonia,cp,municipio,estado,upper(referencias) AS referencias,paquete,numero," &
" t3.id_contrato,id_paquete FROM (" &
" SELECT t2.*,upper(p.nombre) AS paquete FROM" &
" ( SELECT t1.nombre,contrato,upper(ca.nombre) AS calle,numext,numint,colonia,cp,municipio,estado,referencias,id_contrato,id_paquete,numero" &
" FROM( SELECT cli.nombre + ' ' + ap_paterno + ' ' + ap_materno AS nombre,id_contrato,contrato,id_paquete,upper(col.nombre) AS colonia,cp,upper(m.nombre)" &
 " AS municipio,upper(e.nombre) AS estado,id_calle,numext,numint,referencias,telefono AS numero FROM dbo.CLIENTES cli INNER JOIN dbo.CONTRATOS c " &
 " INNER JOIN COLONIAS col INNER JOIN MUNICIPIOS m INNER JOIN ESTADOS e ON e.estado_id=m.estado_id ON m.municipio_id=col.municipio_id ON" &
 " col.colonia_id=c.id_colonia on c.id_cliente=cli.id_cliente WHERE id_contrato=" & id_contrato & " ) AS t1 INNER JOIN CALLES ca ON" &
 " ca.id_calle=t1.id_calle) AS t2 INNER JOIN Paquetes p ON p.id_paquete=t2.id_paquete) AS t3"
      End If

      Dim dtcli As DataTable = con.ConsultarDT(sqlcli)
      If dtcli IsNot Nothing AndAlso dtcli.Rows.Count > 0 Then
        Dim nombre As String = dtcli(0)("nombre").ToString
        Dim contrato As String = dtcli(0)("contrato").ToString
        Dim calle As String = dtcli(0)("calle").ToString
        Dim numext As String = dtcli(0)("numext").ToString
        Dim numint As String = dtcli(0)("numint").ToString
        Dim colonia As String = dtcli(0)("colonia").ToString
        Dim cp As String = dtcli(0)("cp").ToString
        Dim municipio As String = dtcli(0)("municipio").ToString
        Dim estado As String = dtcli(0)("estado").ToString
        Dim referencias As String = dtcli(0)("referencias").ToString
        Dim paquete As String = dtcli(0)("paquete").ToString
        Dim numero As String = dtcli(0)("numero").ToString
        Dim id_paquete As Integer = Val(dtcli(0)("id_paquete").ToString)
        Dim servicios As String = getServicios(id_paquete)
        Dim dtAccount As DataTable = getStpAccount(id_contrato)
        Dim account As String = "s/a"

        If dtAccount IsNot Nothing AndAlso dtAccount.Rows.Count > 0 Then
          account = dtAccount.Rows(0)("clave").ToString()
        End If


        Dim ruta As String = path & "\EstadoCuenta(" & id_estado_cuenta.ToString & ").pdf "
        Dim oDoc As New iTextSharp.text.Document(PageSize.LETTER, 50, 50, 50, 50)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim linea As PdfContentByte
        Dim rectangulo As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Try
          pdfw = PdfWriter.GetInstance(oDoc, New FileStream(ruta,
                    FileMode.Create, FileAccess.Write, FileShare.None))

          Me.PageState = New CustomPageState()
          ''//Wire our event handler and pass in the page state
          pdfw.PageEvent = New MyCustomPdfEvent(Me.PageState)



          'Apertura del documento.
          oDoc.Open()
          cb = pdfw.DirectContent
          linea = pdfw.DirectContent
          rectangulo = pdfw.DirectContent

          'Agregamos una pagina.
          oDoc.NewPage()

          cb.BeginText()
          fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
          cb.SetFontAndSize(fuente, 10) 'fuente definida en la linea anterior y tamaño

          Dim f10 As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f10.SetColor(2, 51, 130)

          Dim f10Bold As New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLUE)
          f10Bold.SetColor(2, 51, 130)


          Dim f14 As New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)
          f14.SetColor(2, 51, 130)

          Dim tblBanner As New PdfPTable(1)
          tblBanner.HorizontalAlignment = 0
          tblBanner.LockedWidth = True
          tblBanner.TotalWidth = 540.0F
          tblBanner.DefaultCell.Border = PdfPCell.NO_BORDER
          tblBanner.DefaultCell.MinimumHeight = 12
          tblBanner.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblBanner.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          'tblBanner.SetWidthPercentage({140.0F, 100.0F, 300.0F}, PageSize.LETTER)

          Dim banner As iTextSharp.text.Image
          banner = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/banner.jpg") 'nombre y ruta de la imagen a insertar
          'imagen.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)

          tblBanner.AddCell(banner)
          oDoc.Add(tblBanner)

          'HEADER
          Dim tblHeader As New PdfPTable(3)
          tblHeader.HorizontalAlignment = 0
          tblHeader.LockedWidth = True
          tblHeader.TotalWidth = 540.0F
          tblHeader.DefaultCell.Border = PdfPCell.NO_BORDER
          tblHeader.DefaultCell.MinimumHeight = 12
          tblHeader.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeader.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblHeader.SetWidthPercentage({140.0F, 100.0F, 300.0F}, PageSize.LETTER)


          'IMAGEN
          Dim imagen As iTextSharp.text.Image 'declaración de imagen
          imagen = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png") 'nombre y ruta de la imagen a insertar
          imagen.ScalePercent(50) 'escala al tamaño de la imagen
          ' imagen.SetAbsolutePosition(50, 700) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)

          tblHeader.AddCell(imagen)
          tblHeader.AddCell(New Paragraph("", FontFactory.GetFont("Helvetica", 8, iTextSharp.text.Font.BOLD)))

          Dim cellInfoEmpresa As New PdfPTable(1)
          cellInfoEmpresa.DefaultCell.Border = PdfPCell.NO_BORDER
          cellInfoEmpresa.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT

          cellInfoEmpresa.AddCell(New Phrase("Comunícalo de México S.A. de C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("Domicilio Fiscal: CONVENTO DE CHURUBUSCO NO. 4,", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("COL. JARDINES DE SANTA MÓNICA, MPIO. TLALNEPANTLA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("DE BAZ, ESTADO DE MÉXICO, C.P. 54050", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellInfoEmpresa.AddCell(New Phrase("RFC: CME0806162SA", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim nesthousing As New PdfPCell(cellInfoEmpresa)
          nesthousing.Border = PdfPCell.NO_BORDER
          nesthousing.Padding = 0F
          nesthousing.HorizontalAlignment = Element.ALIGN_RIGHT
          tblHeader.AddCell(nesthousing)

          oDoc.Add(tblHeader)
          oDoc.Add(New Paragraph(" "))


          'INFO CLIENTE
          Dim tblInfoCliente As New PdfPTable(1)
          tblInfoCliente.HorizontalAlignment = 0
          tblInfoCliente.LockedWidth = True
          tblInfoCliente.TotalWidth = 540.0F
          tblInfoCliente.DefaultCell.Border = PdfPCell.NO_BORDER
          tblInfoCliente.DefaultCell.MinimumHeight = 12
          tblInfoCliente.DefaultCell.HorizontalAlignment = 0
          tblInfoCliente.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblInfoCliente.SetWidthPercentage({540.0F}, PageSize.LETTER)

          Dim cellInfoClient As New PdfPCell(New Phrase(nombre, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInfoClient.Border = PdfPCell.BOX


          'cellInfoClient.BorderWidthBottom = 2
          cellInfoClient.PaddingTop = 0

          tblInfoCliente.AddCell(New Phrase(nombre, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          'tblInfoCliente.AddCell(New Phrase("¡ATENCIÓN!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.RED)))
          tblInfoCliente.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblInfoCliente.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(calle, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblInfoCliente.AddCell(New Phrase("Nueva clabe interbancaria " + account, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          tblInfoCliente.AddCell(New Phrase(referencias & " " & numext & " " & numint, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblInfoCliente.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblInfoCliente.AddCell(New Phrase("¡Si pagas por transferencia olvidate de reportar el pago!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          tblInfoCliente.AddCell(New Phrase(colonia, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblInfoCliente.AddCell(New Phrase(municipio & ", " & estado & ", C.P. " & cp, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          oDoc.Add(tblInfoCliente)
          oDoc.Add(New Paragraph(" "))


          Dim tblPeriodo As New PdfPTable(5)
          tblPeriodo.HorizontalAlignment = 0
          tblPeriodo.LockedWidth = True
          tblPeriodo.TotalWidth = 540.0F
          tblPeriodo.DefaultCell.Border = PdfPCell.NO_BORDER
          tblPeriodo.DefaultCell.MinimumHeight = 12
          tblPeriodo.DefaultCell.HorizontalAlignment = 0
          tblPeriodo.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE
          tblPeriodo.DefaultCell.PaddingLeft = 12.0F
          tblPeriodo.SetWidthPercentage({150.0F, 80.0F, 40.0F, 125.0F, 145.0F}, PageSize.LETTER)


          Dim cellPeriodo3 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPeriodo3.Border = PdfPCell.BOTTOM_BORDER
          cellPeriodo3.BorderWidthBottom = 2
          cellPeriodo3.PaddingTop = 0
          cellPeriodo3.HorizontalAlignment = 0
          cellPeriodo3.Colspan = 5
          cellPeriodo3.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellPeriodo3)
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          tblPeriodo.AddCell(New Phrase("MES DE FACTURACIÓN", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(MonthName(periodoA.Month).ToUpper, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          Dim cell1periodo2 As New PdfPCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cell1periodo2.Border = PdfPCell.RIGHT_BORDER
          cell1periodo2.BorderWidthRight = 2
          cell1periodo2.HorizontalAlignment = 0
          cell1periodo2.BorderColorRight = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))
          tblPeriodo.AddCell(cell1periodo2)

          tblPeriodo.AddCell(New Phrase("TELÉFONO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(numero, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          tblPeriodo.AddCell(New Phrase("FORMA DE PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase("EFECTIVO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("CONTRATO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(contrato, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))


          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("TOTAL A PAGAR", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(grantotal, 2), f10))




          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("PAGAR ANTES DE", f10Bold))
          tblPeriodo.AddCell(New Phrase(periodoA.ToString("dd/MM/yyyy"), f10Bold))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLUE)))

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(cell1periodo2)
          tblPeriodo.AddCell(New Phrase("SALDO VENCIDO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), f10))

          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          'tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 8.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim cellEdocta As New PdfPCell(New Phrase("ESTADO DE CUENTA", f10))
          cellEdocta.Border = PdfPCell.BOTTOM_BORDER
          cellEdocta.BorderWidthBottom = 2
          cellEdocta.PaddingTop = 12.0F
          cellEdocta.PaddingBottom = 5.0F
          cellEdocta.HorizontalAlignment = 1
          cellEdocta.Colspan = 5
          cellEdocta.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellEdocta)

          Dim cellServiciosContratados As New PdfPCell(New Phrase("Servicios contratados", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServiciosContratados.Border = PdfPCell.BOTTOM_BORDER
          cellServiciosContratados.BorderWidthBottom = 2
          cellServiciosContratados.PaddingTop = 5.0F
          cellServiciosContratados.PaddingBottom = 5.0F
          cellServiciosContratados.HorizontalAlignment = 0
          cellServiciosContratados.Colspan = 5
          cellServiciosContratados.BorderColorBottom = New Color(System.Drawing.ColorTranslator.FromHtml("#023382"))

          tblPeriodo.AddCell(cellServiciosContratados)

          Dim cellPaqueteContratado As New PdfPCell(New Phrase(paquete, f10))
          cellPaqueteContratado.Border = PdfPCell.NO_BORDER
          cellPaqueteContratado.BorderWidthBottom = 0
          cellPaqueteContratado.PaddingTop = 5.0F
          cellPaqueteContratado.HorizontalAlignment = 0
          cellPaqueteContratado.Colspan = 5
          cellPaqueteContratado.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPaqueteContratado)

          Dim cellServicios As New PdfPCell(New Phrase(servicios, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellServicios.Border = PdfPCell.NO_BORDER
          cellServicios.BorderWidth = 0
          cellServicios.PaddingTop = 0
          cellServicios.HorizontalAlignment = 0
          cellServicios.Colspan = 4
          cellServicios.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellServicios)
          'tblPeriodo.AddCell(New Phrase(FormatCurrency(total_edo, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(FormatCurrency(totalPlan, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          Dim cellPending As New PdfPCell(New Phrase("SALDO PENDIENTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPending.Border = PdfPCell.NO_BORDER
          cellPending.BorderWidth = 0
          cellPending.PaddingTop = 0
          cellPending.HorizontalAlignment = 0
          cellPending.Colspan = 4
          cellPending.BorderColor = Color.WHITE

          tblPeriodo.AddCell(cellPending)
          tblPeriodo.AddCell(New Phrase(FormatCurrency(saldo_pendiente, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))

          'Show balance'
          Dim auxBalance As Double = If(balance < 0, balance * -1, balance)
          Dim showBalance As Boolean = If(balance <= 0, True, False)

          'If balance < 0 Then
          '  auxBalance = balance * -1
          'End If

          If showBalance Then
            Dim auxShowBalance As Double = auxBalance + grantotal
            Dim cellShowBalance As New PdfPCell(New Phrase("SALDO A FAVOR", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellShowBalance.Border = PdfPCell.NO_BORDER
            cellShowBalance.BorderWidth = 0
            cellShowBalance.PaddingTop = 0
            cellShowBalance.HorizontalAlignment = 0
            cellShowBalance.Colspan = 4
            cellShowBalance.BorderColor = Color.WHITE

            'tblPeriodo.AddCell(cellShowBalance)
            'tblPeriodo.AddCell(New Phrase(FormatCurrency(auxShowBalance, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          End If

          Dim dtCharges As DataTable = getDataBillCharges(id_estado_cuenta)

          If dtCharges IsNot Nothing AndAlso dtCharges.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("OTROS CARGOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtCharges.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtCharges.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtCharges.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim dtDiscount As DataTable = getDataBillDiscounts(id_estado_cuenta)

          If dtDiscount IsNot Nothing AndAlso dtDiscount.Rows.Count > 0 Then
            Dim cellChargesTitle As New PdfPCell(New Phrase("DESCUENTOS", f10))
            cellChargesTitle.Border = PdfPCell.NO_BORDER
            cellChargesTitle.BorderWidthBottom = 0
            cellChargesTitle.PaddingTop = 5.0F
            cellChargesTitle.HorizontalAlignment = 0
            cellChargesTitle.Colspan = 5
            cellChargesTitle.BorderColorBottom = Color.WHITE

            tblPeriodo.AddCell(cellChargesTitle)

            For i = 0 To dtDiscount.Rows.Count - 1
              Dim cellCharges As New PdfPCell(New Phrase(dtDiscount.Rows(0)("nombre").ToString(), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
              cellCharges.Border = PdfPCell.NO_BORDER
              cellCharges.BorderWidth = 0
              cellCharges.PaddingTop = 0
              cellCharges.HorizontalAlignment = 0
              cellCharges.Colspan = 4
              cellCharges.BorderColor = Color.WHITE

              tblPeriodo.AddCell(cellCharges)
              tblPeriodo.AddCell(New Phrase(FormatCurrency(dtDiscount.Rows(0)("importe").ToString(), 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            Next
          End If

          Dim celltotal As New PdfPCell(New Phrase("TOTAL A PAGAR " & FormatCurrency(grantotal, 2), f14))
          celltotal.Border = PdfPCell.NO_BORDER
          celltotal.BorderWidth = 0
          celltotal.PaddingTop = 10.0F
          celltotal.PaddingLeft = 12.0F
          celltotal.HorizontalAlignment = 0
          celltotal.Colspan = 2
          celltotal.BorderColor = Color.WHITE

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(celltotal)

          Dim celltotalLetra As New PdfPCell(New Phrase("(" & totalLetra(grantotal) & ")", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          celltotalLetra.Border = PdfPCell.NO_BORDER
          celltotalLetra.BorderWidth = 0
          celltotalLetra.PaddingTop = 0
          celltotalLetra.PaddingLeft = 12.0F
          celltotalLetra.HorizontalAlignment = 0
          celltotalLetra.Colspan = 2
          celltotalLetra.BorderColor = Color.WHITE

          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(New Phrase(" ", New Font(iTextSharp.text.Font.HELVETICA, 14.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          tblPeriodo.AddCell(celltotalLetra)

          tblPeriodo.AddCell(cellPeriodo3)

          'Balance'

          If showBalance Then
            Dim cellBalance As New PdfPCell(New Phrase("SALDO A FAVOR RESTANTE", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            cellBalance.Border = PdfPCell.NO_BORDER
            cellBalance.BorderWidth = 0
            cellBalance.PaddingTop = 0
            cellBalance.HorizontalAlignment = 0
            cellBalance.Colspan = 4
            cellBalance.BorderColor = Color.WHITE

            tblPeriodo.AddCell(cellBalance)
            tblPeriodo.AddCell(New Phrase(FormatCurrency(auxBalance, 2), New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
            tblPeriodo.AddCell(cellPeriodo3)
          End If

          Dim cellFormasPago As New PdfPCell(New Phrase("FORMAS DE PAGO", f10))
          cellFormasPago.Border = PdfPCell.NO_BORDER
          cellFormasPago.BorderWidthBottom = 0
          cellFormasPago.PaddingTop = 10.0F
          cellFormasPago.HorizontalAlignment = 1
          cellFormasPago.Colspan = 5
          cellFormasPago.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellFormasPago)

          Dim cellStp As New PdfPCell(New Phrase("ATENCIÓN", New Font(iTextSharp.text.Font.HELVETICA, 11.0F, iTextSharp.text.Font.BOLD, Color.RED)))
          cellStp.Border = PdfPCell.NO_BORDER
          cellStp.BorderWidthBottom = 0
          cellStp.PaddingTop = 10.0F
          cellStp.HorizontalAlignment = 0
          cellStp.Colspan = 5
          tblPeriodo.AddCell(cellStp)

          Dim instructions As String = "A PARTIR DE AHORA CADA CLIENTE TENDRÁ UNA CLABE INTERBANCARIA ÚNICA Y PERSONALIZADA POR CONTRATO."
          Dim cellStpInstructions As New PdfPCell(New Phrase(instructions, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellStpInstructions.Border = PdfPCell.NO_BORDER
          cellStpInstructions.BorderWidthBottom = 0
          cellStpInstructions.PaddingTop = 10.0F
          cellStpInstructions.HorizontalAlignment = 0
          cellStpInstructions.Colspan = 5
          cellStpInstructions.BorderColorBottom = Color.BLACK
          tblPeriodo.AddCell(cellStpInstructions)


          Dim cellDeposito As New PdfPCell(New Phrase("DEPOSITO BANCARIO:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellDeposito.Border = PdfPCell.NO_BORDER
          cellDeposito.BorderWidthBottom = 0
          cellDeposito.PaddingTop = 10.0F
          cellDeposito.HorizontalAlignment = 0
          cellDeposito.Colspan = 5
          cellDeposito.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellDeposito)

          Dim cellTransfer As New PdfPCell(New Phrase("TRANSFERENCIA ELECTRÓNICA:", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTransfer.Border = PdfPCell.NO_BORDER
          cellTransfer.BorderWidthBottom = 0
          cellTransfer.PaddingTop = 10.0F
          cellTransfer.HorizontalAlignment = 1
          cellTransfer.Colspan = 2
          cellTransfer.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellTransfer)

          Dim cellFormasPago2 As New PdfPCell(New Phrase("BANCO: STP", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago2.Border = PdfPCell.NO_BORDER
          cellFormasPago2.BorderWidthBottom = 0
          cellFormasPago2.PaddingTop = 2.0F
          cellFormasPago2.HorizontalAlignment = 0
          cellFormasPago2.Colspan = 5
          cellFormasPago2.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago2)


          Dim cellClabe As New PdfPCell(New Phrase("CLABE INTERBANCARIA: 044180256007653656", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellClabe.Border = PdfPCell.NO_BORDER
          cellClabe.BorderWidthBottom = 0
          cellClabe.PaddingTop = 2.0F
          cellClabe.HorizontalAlignment = 2
          cellClabe.Colspan = 2
          cellClabe.BorderColorBottom = Color.WHITE
          'tblPeriodo.AddCell(cellClabe)

          Dim cellFormasPago3 As New PdfPCell(New Phrase("BENEFICIARIO: COMUNICALO DE MÉXICO, S.A. DE C.V.", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago3.Border = PdfPCell.NO_BORDER
          cellFormasPago3.BorderWidthBottom = 0
          cellFormasPago3.PaddingTop = 0.0F
          cellFormasPago3.HorizontalAlignment = 0
          cellFormasPago3.Colspan = 5
          cellFormasPago3.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago3)
          Dim boldTextAccount As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          Dim clabeText As Chunk = New Chunk(account, boldTextAccount)
          Dim indClabe As String = "CLABE INTERBANCARIA PERSONALIZADA: "
          Dim phClabe As Phrase = New Phrase(indClabe, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          phClabe.Add(clabeText)

          Dim cellFormasPago4 As New PdfPCell(New Phrase("CLABE INTERBANCARIA PERSONALIZADA: " + Environment.NewLine + account, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellFormasPago4 = New PdfPCell(phClabe)
          cellFormasPago4.Border = PdfPCell.NO_BORDER
          cellFormasPago4.BorderWidthBottom = 0
          cellFormasPago4.PaddingTop = 0.0F
          cellFormasPago4.HorizontalAlignment = 0
          cellFormasPago4.Colspan = 5
          cellFormasPago4.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellFormasPago4)
          'tblPeriodo.AddCell(phClabe)

          Dim cellEspacio As New PdfPCell(New Phrase("", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellEspacio.Border = PdfPCell.NO_BORDER
          cellEspacio.BorderWidthBottom = 0
          cellEspacio.PaddingTop = 5.0F
          cellEspacio.HorizontalAlignment = 1
          cellEspacio.Colspan = 5
          cellEspacio.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)
          tblPeriodo.AddCell(cellEspacio)

          'Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS PAYNET OPENPAY", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          Dim cellPagoOxxo As New PdfPCell(New Phrase("CÓDIGO PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellPagoOxxo.Border = PdfPCell.NO_BORDER
          cellPagoOxxo.BorderWidthBottom = 0
          cellPagoOxxo.PaddingTop = 5.0F
          cellPagoOxxo.HorizontalAlignment = 1
          cellPagoOxxo.Colspan = 5
          cellPagoOxxo.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPagoOxxo)



          If refOxxo.Trim <> "" And codigoBarraOxxo <> "" Then
            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3
            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
            Dim imgOxxo As iTextSharp.text.Image 'declaración de imagen
            imgOxxo = iTextSharp.text.Image.GetInstance(codigoBarraOxxo) 'nombre y ruta de la imagen a insertar
            'imagen.ScalePercent(50) 'escala al tamaño de la imagen

            Dim cellimgOxxo As New PdfPCell(imgOxxo)
            cellimgOxxo.Border = PdfPCell.NO_BORDER
            cellimgOxxo.BorderWidthBottom = 0
            cellimgOxxo.PaddingTop = 5.0F
            cellimgOxxo.HorizontalAlignment = 1
            cellimgOxxo.Colspan = 5
            cellimgOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellimgOxxo)


            Dim cellrefOxxo As New PdfPCell(New Phrase(refOxxo))
            cellrefOxxo.Border = PdfPCell.NO_BORDER
            cellrefOxxo.BorderWidthBottom = 0
            cellrefOxxo.PaddingTop = 5.0F
            cellrefOxxo.HorizontalAlignment = 1
            cellrefOxxo.Colspan = 5
            cellrefOxxo.BorderColorBottom = Color.WHITE
            tblPeriodo.AddCell(cellrefOxxo)


          End If

          tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)
          'tblPeriodo.AddCell(cellEspacio)


          Dim cellTiendas As New PdfPCell(New Phrase("TIENDAS PARA REALIZAR SU PAGO", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellTiendas.Border = PdfPCell.NO_BORDER
          cellTiendas.BorderWidthBottom = 0
          cellTiendas.PaddingTop = 1.0F
          cellTiendas.HorizontalAlignment = 1
          cellTiendas.Colspan = 5
          cellTiendas.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellTiendas)

          Dim imagenTiendas As iTextSharp.text.Image 'declaración de imagen para las tiendas.

          'imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/tiendasopen.jpg") 'nombre y ruta de la imagen a insertar
          imagenTiendas = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/tiendas.jpeg") 'nombre y ruta de la imagen a insertar
          'imagenTiendas.ScalePercent(44) 'escala al tamaño de la imagen openpay
          imagenTiendas.ScalePercent(40)
          Dim cellimgTiendas As New PdfPCell(imagenTiendas)
          cellimgTiendas.Border = PdfPCell.NO_BORDER
          cellimgTiendas.BorderWidthBottom = 0
          cellimgTiendas.PaddingTop = 5.0F
          cellimgTiendas.HorizontalAlignment = 1  ' 0 para open pay
          cellimgTiendas.Colspan = 5
          cellimgTiendas.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellimgTiendas)

          Dim cellInstrucciones As New PdfPCell(New Phrase("INSTRUCCIONES PARA PAGO EN TIENDAS", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)))
          cellInstrucciones.Border = PdfPCell.NO_BORDER
          cellInstrucciones.BorderWidthBottom = 0
          cellInstrucciones.PaddingTop = 20.0F
          cellInstrucciones.PaddingBottom = 10.0F
          cellInstrucciones.HorizontalAlignment = 1
          cellInstrucciones.Colspan = 5
          cellInstrucciones.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellInstrucciones)

          Dim cellPasps As New PdfPCell(New Phrase("1.- DEBES ELEGIR LA TIENDA QUE MÁS TE CONVENGA ENTRE LAS CADENAS INDICADAS (SOLO SE PUEDE PAGAR EN ESAS TIENDAS).", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          Dim boldText As Font = New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.BOLD, Color.BLACK)
          'Dim compania As Chunk = New Chunk("PAYNET OPENPAY", boldText)
          Dim compania As Chunk = New Chunk("CONEKTA", boldText)
          compania.SetUnderline(0.4, -0.8)
          Dim instruccion As String = "2.- AL ACERCARSE AL MOSTRADOR, DEBERÁ MENCIONAR QUE VIENE A PAGAR "
          Dim ph As Phrase = New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK))
          ph.Add(compania)
          ph.Add(",Y MOSTRAR AL CAJERO EL CÓDIGO DE BARRAS O DICTAR LOS NÚMEROS QUE APARECEN EN LA REFERENCIA.")

          cellPasps = New PdfPCell(ph)
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          instruccion = "3.- UNA VEZ REALIZADO EL PAGO EN EFECTIVO, ENVIAREMOS UNA NOTIFICACIÓN DE PAGO EN TIEMPO REAL A SU CORREO Y ¡LISTO!"
          cellPasps = New PdfPCell(New Phrase(instruccion, New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPasps.Border = PdfPCell.NO_BORDER
          cellPasps.BorderWidthBottom = 0
          cellPasps.PaddingTop = 2.0F
          cellPasps.HorizontalAlignment = 0
          cellPasps.Colspan = 5
          cellPasps.BorderColorBottom = Color.WHITE
          tblPeriodo.AddCell(cellPasps)

          'tblPeriodo.AddCell(cellNota2)

          Dim cellGracias As New PdfPCell(New Phrase("¡MUCHAS GRACIAS POR DARNOS LA OPORTUNIDAD DE SERVIRLE!", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellGracias.Border = PdfPCell.NO_BORDER
          cellGracias.BorderWidthBottom = 0
          cellGracias.PaddingTop = 20.0F
          cellGracias.PaddingBottom = 30.0F
          cellGracias.HorizontalAlignment = 1
          cellGracias.Colspan = 5
          cellGracias.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellGracias)

          tblPeriodo.AddCell(cellPeriodo3)

          Dim cellPie1 As New PdfPCell(New Phrase("ATENCIÓN A CLIENTES: 5526014010", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie1.Border = PdfPCell.NO_BORDER
          cellPie1.BorderWidthBottom = 0
          cellPie1.PaddingTop = 2.0F
          cellPie1.HorizontalAlignment = 0
          cellPie1.Colspan = 3
          cellPie1.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPie1)

          Dim cellPie2 As New PdfPCell(New Phrase("soporte_residencial@comunicalo.mx", New Font(iTextSharp.text.Font.HELVETICA, 10.0F, iTextSharp.text.Font.NORMAL, Color.BLACK)))
          cellPie2.Border = PdfPCell.NO_BORDER
          cellPie2.BorderWidthBottom = 0
          cellPie2.PaddingTop = 2.0F
          cellPie2.HorizontalAlignment = 2
          cellPie2.Colspan = 2
          cellPie2.BorderColorBottom = Color.WHITE

          tblPeriodo.AddCell(cellPie2)

          oDoc.Add(tblPeriodo)


          'Fin del flujo de bytes.
          cb.EndText()
          'Forzamos vaciamiento del buffer.
          pdfw.Flush()
          'Cerramos el documento.
          oDoc.Close()



        Catch ex As Exception
          'Si hubo una excepcion y el archivo existe …
          If File.Exists(ruta) Then
            'Cerramos el documento si esta abierto.
            'Y asi desbloqueamos el archivo para su eliminacion.
            If oDoc.IsOpen Then oDoc.Close()
            '… lo eliminamos de disco.
            File.Delete(ruta)
          End If
          'Throw New Exception("Error al generar archivo PDF (" & ex.Message & ")" & ex.Source)
          MsgBox(ex.Message & "--- " & ex.StackTrace)
          'Dim sqlerror As String = "insert into"
          'Dim sql As String = "insert into netcel..Correos(cliente,mensaje,asunto,estatus,respuesta) values('-1','ERROR AL GENERAR ESTADO DE CUENTA DE COMUNICALO  " & cli_id.ToString & ", MENSAJE:" & ex.Message & "<br/> SOURCE: " & ex.Source & " <br/> STACK TRACE:" & ex.StackTrace & "','ERROR ESTADO DE CUENTA ILOXTELECOM" & cli_id.ToString & "','1','sinfante@mail.ilox.mx')"
          'con.ModRegEli(sql)
          'escribir_log("ERROR AL GENERAR ESTADO DE CUENTA DEL CONTRATO_ID " & idcliente.ToString & ", MENSAJE:" & ex.Message & " SOURCE: " & ex.Source & " STACK TRACE:" & ex.StackTrace)

        Finally
          cb = Nothing
          pdfw = Nothing
          oDoc = Nothing
        End Try


      End If
    End If

  End Sub

  Private Sub generar_Click(sender As Object, e As EventArgs) Handles generar.Click
    Dim sql As String = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente" &
" WHERE contrato='" & txtContrato.Text & "' "
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      For i = 0 To dt.Rows.Count - 1
        Dim id_contrato As Integer = Val(dt(i)("id_contrato").ToString)
        'Dim id_estado_cuenta As Integer = registrarEstado(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
        Dim id_estado_cuenta As Integer = registerBill(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
        If id_estado_cuenta > 0 Then
          Dim ruta As String = ""
          Dim contratos As String = ""

          'ruta = Application.StartupPath & "\EstadosCuenta\" & DateTime.Now.ToString("dd-MM-yyyy")
          ruta = "C:\inetpub\wwwroot\api-comunicalo\Recursos\" & Val(dt(i)("id_cliente").ToString) & "\" & id_contrato & "\Edos"

          System.IO.Directory.CreateDirectory(ruta)
          If Directory.Exists(ruta) Then
            Dim dias As Integer = diasRef(id_contrato)
            Dim refOxxo As Object = referenciaOXXO(id_estado_cuenta, dias)
            Generar_pdfOXXO(id_estado_cuenta, id_contrato, ruta, refOxxo.Referencia, refOxxo.CodigoBarras)
            'Generar_pdfOXXO(id_estado_cuenta, id_contrato, "C:\pdf", "1010102677978684", "https://sandbox-api.openpay.mx/barcode/1010102677978684?width=1&height=45&text=false")
            ''Generar_pdf(id_estado_cuenta, id_contrato, ruta)

            If (File.Exists(ruta & "\EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf")) Then
              Dim mes_facturacion = mesFacturacion(id_estado_cuenta)
              Dim msj As String = crearCorreo(mes_facturacion)
              insertarCorreo(Val(dt(i)("id_cliente").ToString), msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & Val(dt(i)("id_cliente").ToString) & "/" & id_contrato.ToString & "/Edos/EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf", "ltorres@cccard.net")
            Else
              insertarCorreo(-1, "Ocurrio un error al generar el documento del estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar documento de estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
            End If
          End If
        Else
          insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
        End If
      Next
      gvContratos.DataSource = dt
    End If
  End Sub

  Public Function GetMyNumberToWords(ByVal value As String) As String
    Dim str As String = String.Empty
    Select Case Convert.ToDouble(value)
      Case 0 : str = "CERO"
      Case 1 : str = "UN"
      Case 2 : str = "DOS"
      Case 3 : str = "TRES"
      Case 4 : str = "CUATRO"
      Case 5 : str = "CINCO"
      Case 6 : str = "SEIS"
      Case 7 : str = "SIETE"
      Case 8 : str = "OCHO"
      Case 9 : str = "NUEVE"
      Case 10 : str = "DIEZ"
      Case 11 : str = "ONCE"
      Case 12 : str = "DOCE"
      Case 13 : str = "TRECE"
      Case 14 : str = "CATORCE"
      Case 15 : str = "QUINCE"
      Case Is < 20 : str = "DIECI" & GetMyNumberToWords(value - 10)
      Case 20 : str = "VEINTE"
      Case Is < 30 : str = "VEINTI" & GetMyNumberToWords(value - 20)
      Case 30 : str = "TREINTA"
      Case 40 : str = "CUARENTA"
      Case 50 : str = "CINCUENTA"
      Case 60 : str = "SESENTA"
      Case 70 : str = "SETENTA"
      Case 80 : str = "OCHENTA"
      Case 90 : str = "NOVENTA"
      Case Is < 100 : str = GetMyNumberToWords(Int(value \ 10) * 10) & " Y " & GetMyNumberToWords(value Mod 10)
      Case 100 : str = "CIEN"
      Case Is < 200 : str = "CIENTO " & GetMyNumberToWords(value - 100)
      Case 200, 300, 400, 600, 800 : str = GetMyNumberToWords(Int(value \ 100)) & "CIENTOS"
      Case 500 : str = "QUINIENTOS"
      Case 700 : str = "SETECIENTOS"
      Case 900 : str = "NOVECIENTOS"
      Case Is < 1000 : str = GetMyNumberToWords(Int(value \ 100) * 100) & " " & GetMyNumberToWords(value Mod 100)
      Case 1000 : str = "MIL"
      Case Is < 2000 : str = "MIL " & GetMyNumberToWords(value Mod 1000)
      Case Is < 1000000 : str = GetMyNumberToWords(Int(value \ 1000)) & " MIL"
        If value Mod 1000 Then str = str & ” ” & GetMyNumberToWords(value Mod 1000)
      Case 1000000 : str = "UN MILLON"
      Case Is < 2000000 : str = "UN MILLON " & GetMyNumberToWords(value Mod 1000000)
      Case Is < 1000000000000.0# : str = GetMyNumberToWords(Int(value / 1000000)) & " MILLONES "
        If (value - Int(value / 1000000) * 1000000) Then str = str & " " & GetMyNumberToWords(value - Int(value / 1000000) * 1000000)
      Case 1000000000000.0# : str = "UN BILLON"
      Case Is < 2000000000000.0# : str = "UN BILLON " & GetMyNumberToWords(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
      Case Else : str = str(Int(value / 1000000000000.0#)) & " BILLONES"
        If (value - Int(value / 1000000000000.0#) * 1000000000000.0#) Then str = str & " " & GetMyNumberToWords(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
    End Select

    Return str
  End Function
  Public Function totalLetra(ByVal total As Double) As String
    Dim res As String = ""
    Dim largo = Len(CStr(Format(CDbl(total.ToString), “#,###.00”)))
    Dim decimales = Mid(CStr(Format(CDbl(total.ToString), “#,###.00”)), largo - 2)
    res = GetMyNumberToWords(total.ToString - decimales) & " PESOS " & Mid(decimales, Len(decimales) - 1) & "/100 M.N."
    Return res
  End Function

  Private Sub Generador_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'MessageBox.Show(Now.Hour.ToString())
    'MessageBox.Show(Now.Minute.ToString())
    'Dim refOxxo As Object = referenciaOXXO(105487, 29)
    'MessageBox.Show(refOxxo.Referencia)
    'MessageBox.Show(refOxxo.CodigoBarras)

    'Dim id_contrato As Integer = 6749
    'Dim grantotal As Double = 350
    'Dim balance As Decimal = 0
    'Dim sqlBalance As String = "select top 1 coalesce(b.balance,0) as balance from CONTRATOS c WITH (NOLOCK) left join CONTRACTS_BALANCES b WITH (NOLOCK) on c.id_contrato=b.id_contrato " &
    '     "where c.id_contrato=" & id_contrato & " order by b.id desc;"
    'Dim dtBalance As DataTable = con.ConsultarDT(sqlBalance)

    'If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
    'balance = Val(dtBalance(0)("balance").ToString)
    'End If

    'Dim auxTotal As Double = grantotal + balance
    'Dim auxTotalBill = 0

    'If auxTotal <= 0 Then
    'auxTotalBill = 0
    'End If
    'Dim id_estado_cuenta As Integer = registerBill(137, 112, 2)
    'MsgBox(id_estado_cuenta)
    'MessageBox.Show(id_estado_cuenta)
    'MessageBox.Show(auxTotal)
    'MessageBox.Show(auxTotalBill)
    'MessageBox.Show(balance)
    'Dim msj As String = crearCorreo("NOVIEMBRE")
    'MessageBox.Show(msj)
    'Console.WriteLine(msj)
    Generar_pdfOXXO_Rediseno(105479, 7805, "C:\pdf", "1010102677978684", "https://sandbox-api.openpay.mx/barcode/1010102677978684?width=1&height=45&text=false")
    'Dim msj As String = crearCorreo("AGOSTO")
    'insertarCorreo(267, msj, "Comunícalo, estado de cuenta ", "http://localhost/api-comunicalo/Resources/267/242/Edos/EstadoCuenta(16162).pdf", "")
  End Sub

  Private Sub todos_Click(sender As Object, e As EventArgs) Handles todos.Click
    ' MsgBox(dtpFecha.Value.ToString("dd/MM/yyyy"))
    GenerarTodosRefPorFecha(dtpFecha.Value.ToString("dd/MM/yyyy"))
    'MsgBox(System.IO.File.Exists("C:\Users\sergi\Documents\cambios_lorenas.txt"))
  End Sub
  Private Function mesFacturacion(ByVal id_estado_cuenta As Integer) As String
    Dim res As String = MonthName(Now.Date.Month) & " " & Now.Date.Year.ToString
    Dim periodoa As Date
    Dim sql As String = "select periodoA from estados_cuenta where id_estado_cuenta=" & id_estado_cuenta
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      periodoa = dt(0)("periodoa")
      res = MonthName(periodoa.Month).ToUpper & " " & periodoa.Year.ToString
    End If
    Return res
  End Function

  Private Sub GenerarTodos(ByVal fecha As String)
    Dim sql As String = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente WHERE fecha_edo_cta='" & fecha & "' AND c.estatus in(2,3)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      For i = 0 To dt.Rows.Count - 1
        Try
          Dim id_contrato As Integer = Val(dt(i)("id_contrato").ToString)
          'Dim id_estado_cuenta As Integer = registrarEstado(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
          Dim id_estado_cuenta As Integer = registerBill(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
          If id_estado_cuenta > 0 Then
            Dim ruta As String = ""
            Dim contratos As String = ""

            'ruta = Application.StartupPath & "\EstadosCuenta\" & DateTime.Now.ToString("dd-MM-yyyy")
            ruta = "C:\inetpub\wwwroot\api-comunicalo\Recursos\" & Val(dt(i)("id_cliente").ToString) & "\" & id_contrato & "\Edos"

            System.IO.Directory.CreateDirectory(ruta)
            If Directory.Exists(ruta) Then
              '    For Each Row As DataGridViewRow In gvcontratos.Rows
              'If historial_edoscuenta(txtcuenta.Text) < 1 Then
              'registrar_primer_historial(txtcuenta.Text)
              'End If
              Generar_pdf(id_estado_cuenta, id_contrato, ruta)
              '  AddPageNumber()
              If (File.Exists(ruta & "\EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf")) Then
                Dim mes_facturacion = mesFacturacion(id_estado_cuenta)
                Dim msj As String = "<html>" &
    "<head>" &
        "<title></title>" &
    "</head>" &
    "<body>" &
        "<p>" &
            "Estimado cliente," &
        "</p>" &
        "<p>" &
            "Adjunto encontrará su estado de cuenta correspondiente a " & mes_facturacion & ", le pedimos que en cuanto realice su pago nos envíe su comprobante al correo ltorres@comunicalo.mx  o vía WhatsApp 5564161055." &
        "</p>" &
        "<p>" &
            "De antemano le agradecemos mantener su cuenta al corriente y nos permita seguir brindándole nuestros servicios." &
        "</p>" &
        "<p>" &
            "Gracias" &
        "</p>" &
        "<p>" &
            "<img src=|http://201.158.105.66:1518/recursos/LOGOCOMUNICALO.PNG| width=|200px|>" &
        "</p>" &
                                "</body>" &
                                "</html>"
                insertarCorreo(Val(dt(i)("id_cliente").ToString), msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & Val(dt(i)("id_cliente").ToString) & "/" & id_contrato.ToString & "/Edos/EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf", "")
              Else
                insertarCorreo(-1, "Ocurrio un error al generar el documento del estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar documento de estado de cuenta", "", "ltorres@cccard.net;njimenez@comunicalo.mx")
              End If
            End If
          Else
            insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
          End If
        Catch ex As Exception
          insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString & " error: " & ex.Message.ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
        End Try
      Next
    End If
  End Sub
  Private Sub GenerarTodosRefPorFecha(ByVal fecha As String)
    'Dim sql As String = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente WHERE fecha_edo_cta='" & fecha & "' AND c.estatus in(2,3)"
    Dim sql As String = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente" &
" WHERE (id_contrato in(SELECT distinct id_contrato FROM pagos WHERE convert(varchar(10),dateadd(day,-9,periodob),103)='" & fecha & "') or id_contrato in( SELECT DISTINCT id_contrato FROM  pagos WHERE convert(varchar(10),dateadd(day,-9, dateadd(MONTH,1,periodob)),103)='" & fecha & "')) and c.estatus in(2,3)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      For i = 0 To dt.Rows.Count - 1
        Try
          Dim id_contrato As Integer = Val(dt(i)("id_contrato").ToString)
          If revisarEstados(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString)) Then
            Dim id_cliente As Integer = Val(dt(i)("id_cliente").ToString)
            'Dim id_estado_cuenta As Integer = registrarEstado(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
            Dim id_estado_cuenta As Integer = registerBill(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
            If id_estado_cuenta > 0 Then
              Dim ruta As String = ""
              Dim contratos As String = ""

              'ruta = Application.StartupPath & "\EstadosCuenta\" & DateTime.Now.ToString("dd-MM-yyyy")
              ruta = "C:\inetpub\wwwroot\api-comunicalo\Recursos\" & id_cliente.ToString & "\" & id_contrato.ToString & "\Edos"

              System.IO.Directory.CreateDirectory(ruta)
              If Directory.Exists(ruta) Then
                Dim dias As Integer = diasRef(id_contrato)
                Dim refOxxo As Object = referenciaOXXO(id_estado_cuenta, dias)
                Generar_pdfOXXO(id_estado_cuenta, id_contrato, ruta, refOxxo.Referencia, refOxxo.CodigoBarras)
                'Generar_pdf(id_estado_cuenta, id_contrato, ruta)

                If (File.Exists(ruta & "\EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf")) Then
                  Dim mes_facturacion = mesFacturacion(id_estado_cuenta)
                  Dim msj As String = crearCorreo(mes_facturacion)
                  insertarCorreo(id_cliente, msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & id_cliente.ToString & "/" & id_contrato.ToString & "/Edos/EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf", "")
                Else
                  insertarCorreo(-1, "Ocurrio un error al generar el documento del estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar documento de estado de cuenta", "", "ltorres@cccard.net;njimenez@comunicalo.mx")
                End If
              End If
            Else
              insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
            End If
          Else
            insertarCorreo(-1, "No se generó el estado de cuenta del contrato: " & dt(i)("contrato").ToString & " debido a que tiene estados de cuenta pendientes", "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
          End If
        Catch ex As Exception
          insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString & " error: " & ex.Message.ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
        End Try
      Next
      gvContratos.DataSource = dt
    End If
  End Sub

  Private Function revisarEstados(ByVal id_cliente As Integer, ByVal id_contrato As Integer, ByVal estatus As Integer) As Boolean
    Dim cont As Integer = 0
    Dim resp As Boolean = False
    Dim sqlStr = $"select count(*) as cont from ESTADOS_CUENTA where id_cliente=" & id_cliente & " and id_contrato=" & id_contrato & " and estatus=1;"
    Dim dtEstatos = con.ConsultarDT(sqlStr)
    If dtEstatos IsNot Nothing AndAlso dtEstatos.Rows.Count > 0 Then
      cont = Val(dtEstatos(0)("cont").ToString)
    End If

    If estatus = 2 AndAlso cont < 1 Then
      resp = True
    ElseIf estatus = 3 AndAlso cont < 2 Then
      resp = True
    End If

    Return resp
  End Function

  Private Function crearCorreo(ByVal mes_facturacion As String) As String
    Dim msj As String = "<html>" &
        "<head>" &
            "<title></title>" &
        "</head>" &
        "<body>" &
            "<p>" &
                "Estimado cliente," &
            "</p>" &
            "<p>" &
                "Adjunto encontrará su estado de cuenta correspondiente a " & mes_facturacion &
            "</p>" &
            "<p><strong>COMUNICADO OFICIAL</strong></p>" &
          "<p>Estimado Cliente:</p>" &
          "<p>Agradecemos sinceramente su lealtad y confianza en COMUNÍCALO DE MÉXICO, S.A. DE C.V. Nuestro compromiso es continuar brindándole un servicio de calidad, a la altura de sus expectativas.</p>" &
          "<p>Por este medio le informamos lo siguiente:</p>" &
          "<p>Al momento de realizarse una visita para instalación, reparación y/o mantenimiento de su servicio, nuestro personal deberá:</p>" &
          "<ul>" &
              "<li>Presentarse debidamente identificado con credencial vigente con fotografía que lo acredite como trabajador de COMUNÍCALO DE MÉXICO, S.A. DE C.V.</li>" &
              "<li>Portar en todo momento el uniforme oficial de la empresa.</li>" &
              "<li>Llevar consigo la orden de servicio correspondiente, así como la herramienta necesaria para la realización de sus labores.</li>" &
              "<li>Dirigirse en todo momento con respeto y profesionalismo, asegurando un trato cordial y responsable hacia nuestros clientes.</li>" &
          "</ul>" &
          "<p><strong>Es importante destacar que:</strong></p>" &
          "<p>Los trabajadores de COMUNÍCALO DE MÉXICO, S.A. DE C.V. tienen estrictamente prohibido solicitar, recibir o cobrar cualquier tipo de remuneración económica (en efectivo, depósito, transferencia bancaria o por cualquier otro medio) por los trabajos, instalaciones de equipos y/o accesorios relacionados con la prestación del servicio.</p>" &
          "<p>Asimismo, los trabajadores tienen estrictamente prohibido vender servicios adicionales no oficiales como plataformas, equipos o dispositivos ajenos a los que COMUNÍCALO DE MÉXICO, S.A. DE C.V. ofrece como parte de sus servicios oficiales, siendo estos cualquiera que intenten cobrar directamente y no a través del estado de cuenta que le emitimos mensualmente.</p>" &
          "<p>Puede consultar toda la información relacionada con nuestros paquetes, servicios, así como los términos y condiciones aplicables, en nuestro sitio web <a href=""https://www.comunicalo.mx"">www.comunicalo.mx</a>, dentro del apartado ""Términos y Condiciones"" o en <a href=""https://tarifas.ift.org.mx/ift_visor/"">https://tarifas.ift.org.mx/ift_visor/</a></p>" &
          "<p>Todos los cargos que pudieran generarse por concepto de servicios, reparación, cambio y/o instalación de equipos y accesorios complementarios se verán reflejados directamente en su próximo Estado de Cuenta, en donde se detallarán los conceptos y montos correspondientes.</p>" &
          "<p>Finalmente, COMUNÍCALO DE MÉXICO, S.A. DE C.V. se deslinda totalmente de cualquier pago que un suscriptor realice directamente a un trabajador, por cualquier medio o en especie, al margen del procedimiento aquí establecido.</p>" &
          "<p>En caso de identificar o sospechar cualquier intento de ofrecimiento de servicios no oficiales, cobros realizados por medios no autorizados o cualquier irregularidad en las actividades de nuestro personal, le invitamos a reportarlo de inmediato a través de nuestros canales oficiales de atención: teléfono 55 2601 4010 o al correo electrónico <a href=""mailto:residencial@comunicalo.mx"">residencial@comunicalo.mx</a>, donde con gusto le brindaremos la asistencia correspondiente.</p>" &
          "<p>Agradecemos de antemano su atención y colaboración para mantener una relación transparente y de confianza.</p>" &
          "<p><strong>Atentamente,<br>COMUNÍCALO DE MÉXICO, S.A. DE C.V.</strong></p>" &
            "<p></p>" &
            "<p>Apreciaríamos conocer su opinión sobre la calidad del servicio que le hemos ofrecido. Por favor, completa nuestra breve encuesta haciendo clic en el siguiente enlace: <br /> <a href=""https://forms.gle/Ura9byQGj6rwToXZ6"" target=""_blank"" style=""text-decoration:underline;""><strong>https://forms.gle/Ura9byQGj6rwToXZ6</strong></a></p>" &
            "<hr />" &
            "<p style=""font-weight:bold; font-size: 18px;""><strong>HORARIO DE ATENCIÓN</strong></p>" &
            "<p>Lunes a Viernes 9:00 am a 6:00 pm</p>" &
            "<p>Para atender sus inquietudes, dudas o quejas puede enviarnos un correo: <a href=""mailto:atc_clientes@comunicalo.mx"">atc_clientes@comunicalo.mx</a></p>" &
            "<p>Puedes consultar la carta de derechos mínimos del usuario en nuestra página de Internet <a href=""https://www.comunicalo.mx"">www.comunicalo.mx</a></p>" &
            "<p></p>" &
            "<p style=|text-align:justify;|>*Nota: No olvide que para consultarlo, es necesario tener instalado el software Adobe Acrobat Reader 5.0 o superior, si aún no lo tiene puede instalarlo de manera gratuita en: <a href=""http://www.adobe.com/products/acrobat/readstep2.html"">http://www.adobe.com/products/acrobat/readstep2.html</a> </p>" &
            "<p>" &
                "<img src=""https://www.comunicalo.mx/img/logo.png"" width=""200px"">" &
            "</p>" &
                                    "</body>" &
                                    "</html>"

    Dim newContent As String =
        "<p><strong>COMUNICADO OFICIAL</strong></p>" &
        "<p>Estimado Cliente:</p>" &
        "<p>Agradecemos sinceramente su lealtad y confianza en COMUNÍCALO DE MÉXICO, S.A. DE C.V. Nuestro compromiso es continuar brindándole un servicio de calidad, a la altura de sus expectativas.</p>" &
        "<p>Por este medio le informamos lo siguiente:</p>" &
        "<p>Al momento de realizarse una visita para instalación, reparación y/o mantenimiento de su servicio, nuestro personal deberá:</p>" &
        "<ul>" &
            "<li>Presentarse debidamente identificado con credencial vigente con fotografía que lo acredite como trabajador de COMUNÍCALO DE MÉXICO, S.A. DE C.V.</li>" &
            "<li>Portar en todo momento el uniforme oficial de la empresa.</li>" &
            "<li>Llevar consigo la orden de servicio correspondiente, así como la herramienta necesaria para la realización de sus labores.</li>" &
            "<li>Dirigirse en todo momento con respeto y profesionalismo, asegurando un trato cordial y responsable hacia nuestros clientes.</li>" &
        "</ul>" &
        "<p><strong>Es importante destacar que:</strong></p>" &
        "<p>Los trabajadores de COMUNÍCALO DE MÉXICO, S.A. DE C.V. tienen estrictamente prohibido solicitar, recibir o cobrar cualquier tipo de remuneración económica (en efectivo, depósito, transferencia bancaria o por cualquier otro medio) por los trabajos, instalaciones de equipos y/o accesorios relacionados con la prestación del servicio.</p>" &
        "<p>Asimismo, los trabajadores tienen estrictamente prohibido vender servicios adicionales no oficiales como plataformas, equipos o dispositivos ajenos a los que COMUNÍCALO DE MÉXICO, S.A. DE C.V. ofrece como parte de sus servicios oficiales, siendo estos cualquiera que intenten cobrar directamente y no a través del estado de cuenta que le emitimos mensualmente.</p>" &
        "<p>Puede consultar toda la información relacionada con nuestros paquetes, servicios, así como los términos y condiciones aplicables, en nuestro sitio web <a href=""https://www.comunicalo.mx"">www.comunicalo.mx</a>, dentro del apartado ""Términos y Condiciones"" o en <a href=""https://tarifas.ift.org.mx/ift_visor/"">https://tarifas.ift.org.mx/ift_visor/</a></p>" &
        "<p>Todos los cargos que pudieran generarse por concepto de servicios, reparación, cambio y/o instalación de equipos y accesorios complementarios se verán reflejados directamente en su próximo Estado de Cuenta, en donde se detallarán los conceptos y montos correspondientes.</p>" &
        "<p>Finalmente, COMUNÍCALO DE MÉXICO, S.A. DE C.V. se deslinda totalmente de cualquier pago que un suscriptor realice directamente a un trabajador, por cualquier medio o en especie, al margen del procedimiento aquí establecido.</p>" &
        "<p>En caso de identificar o sospechar cualquier intento de ofrecimiento de servicios no oficiales, cobros realizados por medios no autorizados o cualquier irregularidad en las actividades de nuestro personal, le invitamos a reportarlo de inmediato a través de nuestros canales oficiales de atención: teléfono 55 2601 4010 o al correo electrónico <a href=""mailto:residencial@comunicalo.mx"">residencial@comunicalo.mx</a>, donde con gusto le brindaremos la asistencia correspondiente.</p>" &
        "<p>Agradecemos de antemano su atención y colaboración para mantener una relación transparente y de confianza.</p>" &
        "<p><strong>Atentamente,<br>COMUNÍCALO DE MÉXICO, S.A. DE C.V.</strong></p>"

    Dim marker As String = "<p>Nuestras formas de pago son las siguientes:</p>"
    Dim init As Integer = msj.IndexOf(marker)

    'If init >= 0 Then
    'msj = msj.Substring(0, init + marker.Length) & newContent & "</body></html>"
    'End If

    Return msj
  End Function

  Private Sub GenerarTodosRef()
    'Dim sql As String = "Select cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c On c.id_cliente = cli.id_cliente WHERE fecha_edo_cta='" & fecha & "' AND c.estatus in(2,3)"
    'Dim sql As String = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente" &
    '" WHERE id_contrato in(SELECT distinct id_contrato FROM pagos WHERE convert(varchar(10),dateadd(day,-9,periodob),103)=convert(varchar(10),getdate(),103)) and c.estatus in(2,3)"
    Dim sql As String = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente" &
 " WHERE id_contrato in(SELECT distinct id_contrato FROM pagos WHERE convert(varchar(10),dateadd(day,-9,periodob),103)=convert(varchar(10),getdate(),103)) and c.estatus=2" &
 " UNION all" &
 " SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente" &
 " WHERE id_contrato in( SELECT DISTINCT id_contrato FROM  pagos WHERE convert(varchar(10),dateadd(day,-9, dateadd(MONTH,1,periodob)),103)=convert(varchar(10),getdate(),103)) AND c.estatus=3"

    'sql = "SELECT cli.id_cliente,c.id_contrato,contrato,c.estatus FROM clientes cli INNER JOIN dbo.CONTRATOS c ON c.id_cliente = cli.id_cliente" &
    '" WHERE (id_contrato in(SELECT distinct id_contrato FROM pagos WHERE convert(varchar(10),dateadd(day,-9,periodob),103)=convert(varchar(10),getdate(),103)) or id_contrato in( SELECT DISTINCT id_contrato FROM  pagos WHERE convert(varchar(10),dateadd(day,-9, dateadd(MONTH,1,periodob)),103)=convert(varchar(10),getdate(),103))) and c.estatus in(2,3)"

    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      For i = 0 To dt.Rows.Count - 1
        Try
          Dim id_contrato As Integer = Val(dt(i)("id_contrato").ToString)
          If revisarEstados(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString)) Then
            'Dim id_estado_cuenta As Integer = registrarEstado(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
            Dim id_estado_cuenta As Integer = registerBill(Val(dt(i)("id_cliente").ToString), id_contrato, Val(dt(i)("estatus").ToString))
            If id_estado_cuenta > 0 Then
              Dim ruta As String = ""
              Dim contratos As String = ""
              ruta = "C:\inetpub\wwwroot\api-comunicalo\Recursos\" & Val(dt(i)("id_cliente").ToString) & "\" & id_contrato & "\Edos"
              'ruta = "C:\inetput\wwwroot\api-comunicalo\Recursos\" & Val(dt(i)("id_cliente").ToString) & "\" & Val(dt(i)("id_contrato").ToString) & "\Edos"
              System.IO.Directory.CreateDirectory(ruta)
              If Directory.Exists(ruta) Then
                Dim dias As Integer = diasRef(id_contrato)
                Dim refOxxo As Object = referenciaOXXO(id_estado_cuenta, dias)
                Generar_pdfOXXO(id_estado_cuenta, id_contrato, ruta, refOxxo.Referencia, refOxxo.CodigoBarras)
                'Generar_pdf(id_estado_cuenta, id_contrato, ruta)

                If (File.Exists(ruta & "\EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf")) Then
                  Dim mes_facturacion = mesFacturacion(id_estado_cuenta)
                  Dim msj As String = crearCorreo(mes_facturacion)
                  insertarCorreo(Val(dt(i)("id_cliente").ToString), msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & Val(dt(i)("id_cliente").ToString) & "/" & id_contrato.ToString & "/Edos/EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf", "ltorres@cccard.net")
                Else
                  insertarCorreo(-1, "Ocurrio un error al generar el documento del estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar documento de estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
                End If
              End If
            Else
              insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
            End If
          Else
            insertarCorreo(-1, "No se generó el estado de cuenta del contrato: " & dt(i)("contrato").ToString & " debido a que tiene estados de cuenta pendientes", "No se generó estado de cuenta", "", "njimenez@comunicalo.mx")
          End If
        Catch ex As Exception
          insertarCorreo(-1, "Ocurrio un error al registrar el estado de cuenta del contrato: " & dt(i)("contrato").ToString & " error: " & ex.Message.ToString, "Error al generar estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
        End Try

      Next
      gvContratos.DataSource = dt
    End If
  End Sub
  Private Function diasRef(ByVal id_contrato As Integer) As Integer
    Dim dias As Integer = 10
    Dim sql As String = "SELECT convert(varchar(10),fecha_edo_cta,103) AS fecha_edo_cta FROM contratos WHERE id_Contrato=" & id_contrato
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      Dim fecha_edo_cta As Date = dt(0)("fecha_edo_cta").ToString
      dias = 29 'DateDiff(DateInterval.Day, Now.Date, fecha_edo_cta) - 1
      If dias < 10 Then
        dias = 10
      End If
    End If
    Return dias
  End Function

  Private Sub insertarCorreo(ByVal id_cliente As Integer, ByVal msj As String, ByVal asunto As String, ByVal adjuntos As String, ByVal extras As String)
    Dim sql As String = "insert into Alertas values(" & id_cliente & ",'" & msj & "','" & asunto & "','" & adjuntos & "','" & extras & "','default',1,'',getdate())"
    con.ModRegEli(sql)
  End Sub

  Private Function referenciaOXXO(ByVal id_estado_cuenta As Integer, ByVal dias As Integer) As Object
    Dim referencia As String = ""
    Dim codigoBarras As String = ""
    Try
      Dim sql As String = "SELECT id_estado_cuenta,cli.id_cliente,c.id_contrato,upper(nombre + ' ' + ap_paterno) AS nombre,telefono,email,grantotal" &
" FROM dbo.ESTADOS_CUENTA ec INNER JOIN dbo.CONTRATOS c INNER JOIN CLIENTES cli" &
" ON cli.id_cliente=c.id_cliente ON c.id_contrato=ec.id_contrato WHERE ec.id_estado_cuenta=" & id_estado_cuenta
      Dim dt As DataTable = con.ConsultarDT(sql)
      If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        Dim id_cliente As Integer = Val(dt(0)("id_cliente").ToString)
        Dim id_contrato As Integer = Val(dt(0)("id_contrato").ToString)
        Dim nombre As String = dt(0)("nombre").ToString.Replace("ñ", "n").Replace("'", "")
        Dim telefono As String = dt(0)("telefono").ToString
        Dim email As String = dt(0)("email").ToString.Replace("ñ", "n")
        Dim grantotal As Double = Val(dt(0)("grantotal").ToString)

        nombre = Eliminar_Acentos(nombre)

        Dim webC As New WebClient()
        webC.Headers.Add("Content-Type", "application/json")
        webC.Headers.Add("token", "024D8C18-1EE6-47FD-ABD4-545BF169F708")


        Dim info() As Byte = webC.UploadData("http://localhost/api-comunicalo/api/CargosOnline/Oxxo", "POST", Encoding.Default.GetBytes("{""Servicio"": ""Mensualidad Comunicalo"", ""Cantidad"": 1, ""Costo"": " & FormatNumber(grantotal, 2).Replace(",", "") & ", ""Nombre"": """ & nombre & """, ""Duracion"": " & dias & ",""IdCliente"": " & id_cliente & ",""IdContrato"": " & id_contrato & ",""IdEstadoCuenta"": " & id_estado_cuenta & ", ""Telefono"": """ & telefono & """,""Email"": """ & email & """}"))
        'Dim info() As Byte = webC.UploadData("http://localhost/api-comunicalo/api/CargosOnline/OpenPay", "POST", Encoding.Default.GetBytes("{""Servicio"": ""Mensualidad Comunicalo"", ""Cantidad"": 1, ""Costo"": " & FormatNumber(grantotal, 2).Replace(",", "") & ", ""Nombre"": """ & nombre & """, ""Duracion"": " & dias & ",""IdCliente"": " & id_cliente & ",""IdContrato"": " & id_contrato & ",""IdEstadoCuenta"": " & id_estado_cuenta & ", ""Telefono"": """ & telefono & """,""Email"": """ & email & """}"))
        'Open pay local test'
        'Dim info() As Byte = webC.UploadData("http://localhost:5000/api/CargosOnline/OpenPay", "POST", Encoding.Default.GetBytes("{""Servicio"": ""Mensualidad Comunicalo"", ""Cantidad"": 1, ""Costo"": " & FormatNumber(grantotal, 2).Replace(",", "") & ", ""Nombre"": """ & nombre & """, ""Duracion"": " & dias & ",""IdCliente"": " & id_cliente & ",""IdContrato"": " & id_contrato & ",""IdEstadoCuenta"": " & id_estado_cuenta & ", ""Telefono"": """ & telefono & """,""Email"": """ & email & """}"))
        'Dim info() As Byte = webC.UploadData("http://comunicalodemexico.com.mx:1518/api-comunicalo/api/CargosOnline/Oxxo", "POST", Encoding.Default.GetBytes("{""Servicio"": ""Mensualidad Comunicalo"", ""Cantidad"": 1, ""Costo"": " & FormatNumber(grantotal, 2).Replace(",", "") & ", ""Nombre"": """ & nombre & """, ""Duracion"": " & dias & ",""IdCliente"": " & id_cliente & ",""IdContrato"": " & id_contrato & ",""IdEstadoCuenta"": " & id_estado_cuenta & ", ""Telefono"": """ & telefono & """,""Email"": """ & email & """}"))
        'Dim info() As Byte = webC.UploadData("http://comunicalodemexico.com.mx:1518/api-comunicalo/api/CargosOnline/Oxxo", "POST", Encoding.Default.GetBytes("{""Servicio"": ""Mensualidad Comunicalo"", ""Cantidad"": 1, ""Costo"": " & FormatNumber(grantotal, 2) & ", ""Nombre"": """ & nombre & """, ""Duracion"": " & dias & ",""IdCliente"": " & id_cliente & ",""IdContrato"": " & id_contrato & ",""IdEstadoCuenta"": " & id_estado_cuenta & ", ""Telefono"": """ & telefono & """,""Email"": """ & email & """}"))
        Dim i As String = Encoding.Default.GetString(info)
        Dim datos As JObject = JObject.Parse(i)
        'Console.WriteLine("pago online")
        'Console.WriteLine(datos.SelectToken("mensaje"))
        'MsgBox(datos.SelectToken("mensaje"))
        If Boolean.Parse(datos.SelectToken("exito")) Then
          referencia = datos.SelectToken("info.referencia")
          codigoBarras = datos.SelectToken("info.codigoBarras")
          'MsgBox("Referenci: " & referencia & ControlChars.NewLine & "Codigo Barras: " & codigoBarras)
        Else
          ' MsgBox(datos.SelectToken("mensaje"))
          insertarCorreo(-1, "Ocurrio un error al generar la referencia OXXO del estado de cuenta " & id_estado_cuenta.ToString & " con el msj:" & datos.SelectToken("mensaje").ToString, "Error al generar referencia OXXO", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
        End If
      End If
      Return New With {.Referencia = referencia, .CodigoBarras = codigoBarras}
    Catch ex As Exception
      insertarCorreo(-1, "Ocurrio una excepción al generar la referencia OXXO del estado de cuenta " & id_estado_cuenta.ToString & " con el msj:" & ex.Message, "Error al generar referencia OXXO", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
      Return Nothing
    End Try
  End Function
  Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    Try
      Dim ruta As String = "C:\inetput\wwwroot\api-comunicalo\Recursos\" & 60 & "\" & 58 & "\Edos"
      Dim refOxxo As Object = referenciaOXXO(85, 10)
      System.IO.Directory.CreateDirectory(ruta)
      If Directory.Exists(ruta) Then
        Generar_pdfOXXO(85, 58, ruta, refOxxo.Referencia, refOxxo.CodigoBarras)
      End If
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try
    'Dim nombre As String = "pasó áctivo"
    'MsgBox(Eliminar_Acentos(nombre))
  End Sub
  Private Function Eliminar_Acentos(ByVal accentedStr As String) As String
    Dim tempBytes As Byte()
    tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(accentedStr)
    Return System.Text.Encoding.UTF8.GetString(tempBytes)
  End Function
  Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    If Now.Hour = 7 And Now.Minute = 0 AndAlso Not Generados() Then
      Timer1.Enabled = False
      GenerarTodosRef()
      'Threading.Thread.Sleep(5000)
      Me.Close()
      'Timer1.Enabled = True
    End If
  End Sub
  Private Function Generados() As Boolean
    Dim res As Boolean = False
    Dim sql As String = "select count(*) from ESTADOS_CUENTA where convert(varchar(10),fecha,103)=convert(varchar(10),getdate(),103)"
    Dim dt As DataTable = con.ConsultarDT(sql)
    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
      If Val(dt(0)(0).ToString) > 0 Then
        res = True
      End If
    End If
    Return res
  End Function
  Private Sub btnEstado_Click(sender As Object, e As EventArgs) Handles btnEstado.Click
    Dim id_estado_cuenta As Integer = Val(txtEstado.Text)
    If id_estado_cuenta > 0 Then

      Dim sql As String = "SELECT * FROM dbo.ESTADOS_CUENTA ec WHERE ec.id_estado_cuenta=" & id_estado_cuenta
      Dim dt As DataTable = con.ConsultarDT(sql)
      If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

        Dim ruta As String = ""
        Dim contratos As String = ""
        Dim id_contrato As Integer = Val(dt(0)("id_contrato").ToString)

        'ruta = Application.StartupPath & "\EstadosCuenta\" & DateTime.Now.ToString("dd-MM-yyyy")
        ruta = "C:\inetput\wwwroot\api-comunicalo\Recursos\" & Val(dt(0)("id_cliente").ToString) & "\" & Val(dt(0)("id_contrato").ToString) & "\Edos"

        System.IO.Directory.CreateDirectory(ruta)
        If Directory.Exists(ruta) Then
          Dim dias As Integer = diasRef(id_contrato)
          Dim refOxxo As Object = referenciaOXXO(id_estado_cuenta, dias)
          Generar_pdfOXXO(id_estado_cuenta, id_contrato, ruta, refOxxo.Referencia, refOxxo.CodigoBarras)
          'Generar_pdf(id_estado_cuenta, id_contrato, ruta)

        End If

      Else
        MsgBox("no se pudo obtener la informacion del contrato")
      End If

    Else
      MsgBox("Ingresa el estado de cuenta")
    End If
  End Sub

  Private Sub regenerar_pdf_por_id(ByVal id As Integer)
    Dim contador As Integer = 0
    Try
      Dim sql As String = "SELECT * FROM dbo.ESTADOS_CUENTA ec WHERE ec.id_estado_cuenta >=" & id
      Dim dt As DataTable = con.ConsultarDT(sql)
      If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        For i = 0 To dt.Rows.Count - 1
          Dim ruta As String = ""
          Dim contratos As String = ""
          Dim id_contrato As Integer = Val(dt(i)("id_contrato").ToString)
          Dim id_cliente As Integer = Val(dt(i)("id_cliente").ToString)
          Dim id_estado_cuenta As Integer = Val(dt(i)("id_estado_cuenta").ToString)
          ruta = "C:\inetpub\wwwroot\api-comunicalo\Recursos\" & Val(dt(i)("id_cliente").ToString) & "\" & Val(dt(i)("id_contrato").ToString) & "\Edos"
          ''ruta = "C:\pdf"
          System.IO.Directory.CreateDirectory(ruta)
          If Directory.Exists(ruta) Then
            Dim referencia As String = ""
            Dim codigoBarras As String = ""

            Dim sqlref As String = "SELECT po.referencia,po.codigo_barras FROM dbo.PAGOS_ONLINE po WHERE po.id_edo_cta=" & id_estado_cuenta
            Dim dtref As DataTable = con.ConsultarDT(sqlref)
            If dtref IsNot Nothing AndAlso dtref.Rows.Count > 0 Then
              referencia = dtref(0)("referencia").ToString
              codigoBarras = dtref(0)("codigo_barras").ToString
              Generar_pdfOXXO(id_estado_cuenta, id_contrato, ruta, referencia, codigoBarras)
              ' Reenvia el correo con el estado de cuenta y el archivo.
              If (File.Exists(ruta & "\EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf")) Then
                Dim mes_facturacion = mesFacturacion(id_estado_cuenta)
                Dim msj As String = crearCorreo(mes_facturacion)
                insertarCorreo(id_cliente, msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & id_cliente.ToString & "/" & id_contrato.ToString & "/Edos/EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf", "")
                ' MsgBox("Archivo generado y enviado")
                contador += 1
              Else
                ' MsgBox("No se encontro el archivo o no se pudo generar")
                insertarCorreo(-1, "Ocurrio un error al generar el documento del estado de cuenta del contrato: " & id_contrato.ToString, "Error al generar documento de estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
              End If
            Else
              MsgBox("No se encontro la referencia asignada al estado de cuenta " & id_estado_cuenta.ToString)
            End If
          End If
        Next
      Else
        MsgBox("no se pudieron obtener registros en ese rango")
      End If
    Catch ex As Exception
      MsgBox("Ocurrió un error " + ex.Message.ToString)
    Finally
      MsgBox("Se enviaron/regeneraron " & contador.ToString & " Estados de cuenta")
    End Try
  End Sub

  Private Sub regenerar_pdf_Click(sender As Object, e As EventArgs) Handles regenerar_pdf.Click
    Dim id_estado_cuenta As Integer = Val(txtEstado.Text)
    If id_estado_cuenta > 0 Then

      Dim sql As String = "SELECT * FROM dbo.ESTADOS_CUENTA ec WHERE ec.id_estado_cuenta=" & id_estado_cuenta
      Dim dt As DataTable = con.ConsultarDT(sql)
      If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

        Dim ruta As String = ""
        Dim contratos As String = ""
        Dim id_contrato As Integer = Val(dt(0)("id_contrato").ToString)
        Dim id_cliente As Integer = Val(dt(0)("id_cliente").ToString)
        'ruta = Application.StartupPath & "\EstadosCuenta\" & DateTime.Now.ToString("dd-MM-yyyy")
        ruta = "C:\inetpub\wwwroot\api-comunicalo\Recursos\" & Val(dt(0)("id_cliente").ToString) & "\" & Val(dt(0)("id_contrato").ToString) & "\Edos"
        ''ruta = "C:\pdf"
        System.IO.Directory.CreateDirectory(ruta)
        If Directory.Exists(ruta) Then
          Dim referencia As String = ""
          Dim codigoBarras As String = ""

          Dim sqlref As String = "SELECT po.referencia,po.codigo_barras FROM dbo.PAGOS_ONLINE po WHERE po.id_edo_cta=" & id_estado_cuenta
          Dim dtref As DataTable = con.ConsultarDT(sqlref)
          If dtref IsNot Nothing AndAlso dtref.Rows.Count > 0 Then
            referencia = dtref(0)("referencia").ToString
            codigoBarras = dtref(0)("codigo_barras").ToString
            Generar_pdfOXXO(id_estado_cuenta, id_contrato, ruta, referencia, codigoBarras)
            ' Reenvia el correo con el estado de cuenta y el archivo.
            If (File.Exists(ruta & "\EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf")) Then
              Dim mes_facturacion = mesFacturacion(id_estado_cuenta)
              Dim msj As String = crearCorreo(mes_facturacion)
              insertarCorreo(id_cliente, msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & id_cliente.ToString & "/" & id_contrato.ToString & "/Edos/EstadoCuenta(" + id_estado_cuenta.ToString + ").pdf", "")
              MsgBox("Archivo generado y enviado")
            Else
              MsgBox("No se encontro el archivo o no se pudo generar")
              insertarCorreo(-1, "Ocurrio un error al generar el documento del estado de cuenta del contrato: " & id_contrato.ToString, "Error al generar documento de estado de cuenta", "", "njimenez@comunicalo.mx;dcastillo@comunicalo.mx")
            End If
          Else
            MsgBox("No se encontro la referencia asignada al estado de cuenta ")
          End If

          'Generar_pdf(id_estado_cuenta, id_contrato, ruta)
        End If

      Else
        MsgBox("no se pudo obtener la informacion del contrato")
      End If

    Else
      MsgBox("Ingresa el estado de cuenta")
    End If
  End Sub

  Private Sub enviar_Click(sender As Object, e As EventArgs) Handles enviar.Click
    Dim id_estado_cuenta As Integer = Val(txtEstado.Text)
    If id_estado_cuenta > 0 Then

      Dim sql As String = " SELECT * FROM dbo.ESTADOS_CUENTA ec WHERE ec.id_estado_cuenta>" & id_estado_cuenta
      Dim dt As DataTable = con.ConsultarDT(sql)
      If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        For i = 0 To dt.Rows.Count - 1

          Dim mes_facturacion = mesFacturacion(Val(dt(i)("id_estado_cuenta")))
          Dim msj As String = "<html>" &
        "<head>" &
        "<title></title>" &
        "</head>" &
        "<body>" &
        "<p>" &
        "Estimado cliente," &
        "</p>" &
        "<p>" &
        "Adjunto encontrará su estado de cuenta correspondiente a " & mes_facturacion &
        "</p>" &
        "<p>" &
        "Ahora su pago en <b>Oxxo es referenciado</b>, solo tiene que seguir estos sencillos pasos:" &
        "</p>" &
        "<p>" &
        "1.- Indique al cajero que va a realizar un pago de Oxxopay.<br/>" &
        "2.- Muestre el código de barras que aparece en el estado de cuenta adjunto. (puede imprimirlo o desde su teléfono celular)<br/>" &
        "3.- Realice su pago.<br/>" &
        "4.- Le llegará un correo de confirmación y en automático queda reflejado en nuestro sistema." &
        "</p>" &
        "<p>" &
        "Si usted prefiere realizar su pago por deposito bancario o transferencia electrónica le pedimos por favor nos envíe su comprobante al correo ltorres@comunicalo.mx o vía WhatsApp 5564161055." &
        "</p>" &
        "<p>" &
        "De antemano le agradecemos mantener su cuenta al corriente y nos permita seguir brindándole nuestros servicios." &
        "</p>" &
        "<p>" &
        "*Nota: No olvide que para consultarlo, es necesario tener instalado el software Adobe Acrobat Reader 5.0 o superior, si aún no lo tiene puede instalarlo de manera gratuita en: <a href=|http://www.adobe.com/products/acrobat/readstep2.html|>http://www.adobe.com/products/acrobat/readstep2.html</a>" &
        "</p>" &
        "<p>" &
        "<img src=|http://201.158.105.66:1518/recursos/LOGOCOMUNICALO.PNG| width=|200px|>" &
        "</p>" &
                "</body>" &
                "</html>"
          insertarCorreo(Val(dt(i)("id_cliente").ToString), msj, "Comunícalo, estado de cuenta " & mes_facturacion, "http://localhost/api-comunicalo/Resources/" & Val(dt(i)("id_cliente").ToString) & "/" & Val(dt(i)("id_contrato").ToString) & "/Edos/EstadoCuenta(" + dt(i)("id_estado_cuenta").ToString + ").pdf", "")
        Next
      End If

    Else
      MsgBox("Ingresa el estado de cuenta")
    End If
  End Sub

  Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    GenerarTodosRef()

  End Sub

  Private Sub btnReenvio_Click(sender As Object, e As EventArgs) Handles btnReenvio.Click
    Dim id_edo As Integer = Val(txtEdoId.Text)
    If id_edo > 0 Then
      btnReenvio.Enabled = False
      regenerar_pdf_por_id(id_edo)
      txtEdoId.Text = ""
      btnReenvio.Enabled = True
    Else
      MsgBox("Ingresa un valor mayor a 0")
    End If
  End Sub

  Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

  End Sub

  Private Sub gvContratos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles gvContratos.CellContentClick

  End Sub
End Class

Public Class CustomPageState
    Public Property IsLastPage As Boolean = False
End Class

Public Class MyCustomPdfEvent
    Implements IPdfPageEvent
    ''//Reference to the state container
    Private PageState As CustomPageState
    Public template As PdfTemplate
    Public cb As PdfContentByte

    Public Sub New(ByRef pageState As CustomPageState)
        Me.PageState = pageState
    End Sub

    Public Sub OnEndPage(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document) Implements iTextSharp.text.pdf.IPdfPageEvent.OnEndPage
    If Me.PageState.IsLastPage Then
      ''//Last page, do something different
    Else
      Dim fuente As iTextSharp.text.pdf.BaseFont
      cb.BeginText()
      Dim imageInfo As iTextSharp.text.Image
      imageInfo = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/LOGOCOMUNICALO.png")
      imageInfo.ScalePercent(30)
      imageInfo.SetAbsolutePosition(500, 20)
      fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.BOLD).BaseFont
      cb.SetFontAndSize(fuente, 8)
      ' drawing a line'
      cb.SetColorStroke(New Color(System.Drawing.ColorTranslator.FromHtml("#023382")))
      cb.SetLineWidth(2)
      cb.MoveTo(30, 50)
      cb.LineTo(590, 50)
      cb.Stroke()
      ' line
      cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "soporte_residencial@comunicalo.mx", 30, 40, 0)
      cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Atención a clientes: 5526014010", 30, 30, 0)
      cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Horario de atención de 9 a 18 hrs", 30, 20, 0)
      cb.AddImage(imageInfo)
      cb.EndText()

      ' If document.PageNumber >= 10 Then
      'cb.AddTemplate(template, 588, 20)
      'Else
      'cb.AddTemplate(template, 585, 20)
      'En'd If
    End If
  End Sub

    Public Sub OnChapter(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal paragraphPosition As Single, ByVal title As iTextSharp.text.Paragraph) Implements iTextSharp.text.pdf.IPdfPageEvent.OnChapter

    End Sub

    Public Sub OnChapterEnd(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal paragraphPosition As Single) Implements iTextSharp.text.pdf.IPdfPageEvent.OnChapterEnd

    End Sub

  Public Sub OnCloseDocument(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document) Implements iTextSharp.text.pdf.IPdfPageEvent.OnCloseDocument
    Dim fuente As iTextSharp.text.pdf.BaseFont
    template.BeginText()
    fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.BOLD).BaseFont
    template.SetFontAndSize(fuente, 6) 'fuente definida en la linea anterior y tamaño
    template.ShowTextAligned(PdfContentByte.ALIGN_LEFT, (writer.PageNumber - 1).ToString, 0, 0, 0)
    template.EndText()

  End Sub

  Public Sub OnGenericTag(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal rect As iTextSharp.text.Rectangle, ByVal text As String) Implements iTextSharp.text.pdf.IPdfPageEvent.OnGenericTag

    End Sub

  Public Sub OnOpenDocument(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document) Implements iTextSharp.text.pdf.IPdfPageEvent.OnOpenDocument
    cb = writer.DirectContent
    template = cb.CreateTemplate(50, 50)
  End Sub

  Public Sub OnParagraph(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal paragraphPosition As Single) Implements iTextSharp.text.pdf.IPdfPageEvent.OnParagraph

  End Sub

  Public Sub OnParagraphEnd(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal paragraphPosition As Single) Implements iTextSharp.text.pdf.IPdfPageEvent.OnParagraphEnd

  End Sub

  Public Sub OnSection(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal paragraphPosition As Single, ByVal depth As Integer, ByVal title As iTextSharp.text.Paragraph) Implements iTextSharp.text.pdf.IPdfPageEvent.OnSection

    End Sub

  Public Sub OnSectionEnd(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document, ByVal paragraphPosition As Single) Implements iTextSharp.text.pdf.IPdfPageEvent.OnSectionEnd

  End Sub

  Public Sub OnStartPage(ByVal writer As iTextSharp.text.pdf.PdfWriter, ByVal document As iTextSharp.text.Document) Implements iTextSharp.text.pdf.IPdfPageEvent.OnStartPage
    Dim tblBanner As New PdfPTable(1)
    tblBanner.HorizontalAlignment = 0
    tblBanner.LockedWidth = True
    tblBanner.TotalWidth = 550.0F
    tblBanner.DefaultCell.Border = PdfPCell.NO_BORDER
    tblBanner.DefaultCell.MinimumHeight = 12
    tblBanner.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER
    tblBanner.DefaultCell.BackgroundColor = iTextSharp.text.Color.WHITE

    Dim banner As iTextSharp.text.Image
    banner = iTextSharp.text.Image.GetInstance(Application.StartupPath & "/imgs/banner_01.jpg")
    tblBanner.AddCell(banner)
    document.Add(tblBanner)
    document.Add(New Paragraph(" "))
  End Sub
End Class
