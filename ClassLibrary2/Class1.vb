Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Web
Imports HtmlAgilityPack
Imports Tesseract
Imports System.Drawing
Imports System.Text
Public Class Class1
    '    Public Function ConsultaSUNAT(RUC As String) As SunatEN
    '        Try
    '            ValidarRUC(RUC)
    'ReiniciarProceso:
    '            Dim miCookie As New CookieContainer
    '            Dim ImageCache As System.Drawing.Image
    '            Dim TextoCache, URL As String
    '            URL = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image"
    '            Dim wrImage As HttpWebRequest = WebRequest.Create(URL)
    '            wrImage.CookieContainer = miCookie
    '            ImageCache = Image.FromStream(wrImage.GetResponse.GetResponseStream)
    '            'Dim eng As New TesseractEngine("", "eng", EngineMode.Default)
    '            Using engine As New TesseractEngine("./tessdata", "eng", EngineMode.Default)
    '                Dim page As Page = engine.Process(New Bitmap(ImageCache))
    '                TextoCache = page.GetText().ToUpper
    '                TextoCache = Regex.Replace(TextoCache, "[^a-zA-Z]+", String.Empty)
    '                'TextoCache = Regex.Replace(TextoCache, "[^\w\.@-]", String.Empty)
    '            End Using
    '            If TextoCache Is Nothing OrElse TextoCache.Length <> 4 Then
    '                GoTo ReiniciarProceso
    '            End If

    '            Dim DATA As String
    '            URL = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRuc&nroRuc={RUC}&codigo={TextoCache}&tipdoc=1"
    '            DATA = ReadAllHTML(URL, miCookie)

    '            If Not DATA.Contains("El codigo ingresado es incorrecto") Then
    '                Dim docEntity As New HtmlDocument
    '                docEntity.LoadHtml(DATA)
    '                Dim Query = (From ItemNode In (From Item In docEntity.DocumentNode.SelectNodes("//table[1]").Elements("tr")
    '                                               Select Item.SelectNodes("td")).ToList
    '                             Where ItemNode IsNot Nothing AndAlso ItemNode.Count > 1
    '                             Select ItemNode).ToList

    '                Dim ClearNode = Function(Node As String) As String
    '                                    Node = Node.Replace(vbCrLf, "").Trim
    '                                    Node = Node.Replace(vbTab, "")
    '                                    Node = Node.Replace("  ", "")
    '                                    'Node = Regex.Replace(Node, "[^\w\.@-]", String.Empty)
    '                                    Return Node
    '                                End Function
    '                Dim fGet = Function(Index As Integer, IndexValor As Integer) As String
    '                               Dim ValorNode As String = "-"
    '                               Try
    '                                   Dim selectNode = Query(Index)(IndexValor)
    '                                   If selectNode IsNot Nothing Then ValorNode = ClearNode(selectNode.InnerText)
    '                               Catch ex As Exception
    '                                   ValorNode = ex.Message
    '                               End Try
    '                               Return ValorNode
    '                           End Function
    '                Dim fGetList = Function(Index As Integer) As List(Of String)
    '                                   Dim selectNode = Query(Index)(1).Element("select")
    '                                   Dim List As New List(Of String) From {"-"}
    '                                   If selectNode IsNot Nothing Then
    '                                       List = (From Item In Query(Index)(1).Element("select").SelectNodes("option")
    '                                               Select ClearNode(Item.InnerText)).ToList
    '                                   End If
    '                                   Return List
    '                               End Function

    '                Dim NumeroDeRUC, RazonSocial As String
    '                Dim Array = fGet(0, 1).Split("-")
    '                NumeroDeRUC = Array(0).Trim
    '                RazonSocial = Array(1).Trim

    '                Dim objE As SunatEN

    '                If RUC.StartsWith("2") Then
    '                    objE = New SunatEN With {
    '                                            .NumeroDeRUC = NumeroDeRUC,
    '                                            .RazonSocial = RazonSocial,
    '                                            .TipoDeContribuyente = fGet(1, 1),
    '                                            .NombreComercial = fGet(2, 1),
    '                                            .FechaDeInscripcion = fGet(3, 1),
    '                                            .FechaDeInicioDeActividades = fGet(3, 3),
    '                                            .EstadoDeContribuyente = fGet(4, 1),
    '                                            .CondicionDeContribuyente = fGet(5, 1),
    '                                            .DomicilioFiscal = fGet(6, 1),
    '                                            .SistemaDeEmisionDeComprobante = fGet(7, 1),
    '                                            .ActividadDeComercioExterior = fGet(7, 3),
    '                                            .SistemaDeContabilidad = fGet(8, 1),
    '                                            .ActividadesEconomicas = fGetList(9),
    '                                            .ComprobantesDePago = fGetList(10),
    '                                            .SistemaDeEmisionElectronica = fGetList(11),
    '                                            .EmisorElectronicoDesde = fGet(12, 1),
    '                                            .ComprobantesElectronicos = (From Item In fGet(13, 1).Split(",") Select Item.Trim).ToList,
    '                                            .AfiliadoAlPLE = fGet(14, 1)
    '                                            }
    '                Else
    '                    objE = New SunatEN With {
    '                                           .NumeroDeRUC = NumeroDeRUC,
    '                                           .RazonSocial = RazonSocial,
    '                                           .TipoDeContribuyente = fGet(1, 1),
    '                                           .TipoDeDocumento = fGet(2, 1),
    '                                           .NombreComercial = fGet(3, 1),
    '                                           .FechaDeInscripcion = fGet(4, 1),
    '                                           .FechaDeInicioDeActividades = fGet(4, 3),
    '                                           .EstadoDeContribuyente = fGet(5, 1),
    '                                           .CondicionDeContribuyente = fGet(6, 1),
    '                                           .ProfesionUOficio = fGet(6, 3),
    '                                           .DomicilioFiscal = fGet(7, 1),
    '                                           .SistemaDeEmisionDeComprobante = fGet(8, 1),
    '                                           .ActividadDeComercioExterior = fGet(8, 3),
    '                                           .SistemaDeContabilidad = fGet(9, 1),
    '                                           .ActividadesEconomicas = fGetList(10),
    '                                           .ComprobantesDePago = fGetList(11),
    '                                           .SistemaDeEmisionElectronica = fGetList(12),
    '                                           .EmisorElectronicoDesde = fGet(13, 1),
    '                                           .ComprobantesElectronicos = (From Item In fGet(14, 1).Split(",") Select Item.Trim).ToList,
    '                                           .AfiliadoAlPLE = fGet(15, 1)
    '                                           }
    '                End If

    '                Dim newRazonSocial As String = RazonSocial.Replace(" ", "+")
    '                If objE.NumeroDeRUC.StartsWith("2") Then
    '                    'If boolRepLeg Then
    '                    '    URL = $"http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=getRepLeg&desRuc={newRazonSocial}&nroRuc={objE.NumeroDeRUC}"
    '                    '    DATA = ReadAllHTMLEx(URL)

    '                    '    Dim docRepLegal As New HtmlDocument
    '                    '    docRepLegal.LoadHtml(DATA)

    '                    '    Dim NodeError = docRepLegal.DocumentNode.SelectSingleNode("//div[@class='cuerpo']")
    '                    '    If NodeError Is Nothing Then
    '                    '        Dim RepLegal = (From ItemNode In (From Item In docRepLegal.DocumentNode.SelectSingleNode("//td[@class='beta']").ChildNodes(1).SelectNodes("tr")
    '                    '                                          Select Item.SelectNodes("td")).ToList
    '                    '                        Where ItemNode IsNot Nothing
    '                    '                        Select New RepLegalEN With {
    '                    '                                    .Documento = ClearNode(ItemNode(0).InnerText),
    '                    '                                    .NumeroDeDocumento = ClearNode(ItemNode(1).InnerText),
    '                    '                                    .Nombre = ClearNode(ItemNode(2).InnerText),
    '                    '                                    .Cargo = ClearNode(ItemNode(3).InnerText),
    '                    '                                    .FechaDesde = ClearNode(ItemNode(4).InnerText)
    '                    '                                    }).ToList
    '                    '        objE.RepresentanteLegal = RepLegal
    '                    '    End If
    '                    'End If

    '                    'If boolEstablecimientos Then
    '                    '    URL = $"http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?submit=Establecimientos+Anexos&nroRuc"
    '                    '    URL &= $"={objE.NumeroDeRUC}&accion=getLocAnex&desRuc={newRazonSocial}&tamanioPagina=100&pagina=0"
    '                    '    DATA = ReadAllHTMLEx(URL)

    '                    '    Dim docLocAnex As New HtmlDocument
    '                    '    docLocAnex.LoadHtml(DATA)

    '                    '    Dim LocAnex = (From ItemNode In (From Item In docLocAnex.DocumentNode.SelectSingleNode("//td[@class='beta']").ChildNodes(1).Elements("tr")
    '                    '                                     Select Item.SelectNodes("td")).ToList
    '                    '                   Where ItemNode IsNot Nothing
    '                    '                   Select New EstablecimientosEN With {
    '                    '                        .Codigo = ClearNode(ItemNode(0).InnerText),
    '                    '                        .TipoDeEstablecimiento = ClearNode(ItemNode(1).InnerText),
    '                    '                        .Direccion = ClearNode(ItemNode(2).InnerText),
    '                    '                        .ActividadEconomica = ClearNode(ItemNode(3).InnerText)}).ToList

    '                    '    objE.Establecimientos = LocAnex
    '                    'End If

    '                    'If boolCantTrab Then
    '                    '    URL = $"https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=getCantTrab&nroRuc={objE.NumeroDeRUC}&desRuc={newRazonSocial}"
    '                    '    DATA = ReadAllHTMLEx(URL)

    '                    '    If Not DATA.Contains("No existen declaraciones") Then

    '                    '        Dim docCantTrab As New HtmlDocument
    '                    '        docCantTrab.LoadHtml(DATA)
    '                    '        Dim CantTrab = (From ItemNode In (From Item In docCantTrab.DocumentNode.SelectSingleNode("//td[1]/table[1]").Elements("tr")
    '                    '                                          Select Item.SelectNodes("td")).ToList
    '                    '                        Where ItemNode IsNot Nothing AndAlso ItemNode.Count = 4
    '                    '                        Select New TrabajadoresEN With {
    '                    '                           .Periodo = ClearNode(ItemNode(0).InnerText),
    '                    '                           .NroDeTrabajadores = ClearNode(ItemNode(1).InnerText),
    '                    '                           .NroDePensionistas = ClearNode(ItemNode(2).InnerText),
    '                    '                           .NroDePrestadoresDeServicio = ClearNode(ItemNode(3).InnerText)}).ToList

    '                    '        objE.CantidadDeTrabajadores = CantTrab
    '                    '    End If
    '                    'End If

    '                End If
    '                Return objE
    '            Else
    '                Throw New Exception("Error en el código captcha.")
    '            End If
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '    Private Function ReadAllHTML(URL As String, Optional Cookie As CookieContainer = Nothing) As String
    '        Try
    '            Dim Data As String
    '            Dim wrURL As HttpWebRequest = WebRequest.Create(URL)
    '            wrURL.CookieContainer = Cookie
    '            Using SR As New StreamReader(wrURL.GetResponse.GetResponseStream, System.Text.Encoding.Default)
    '                Data = HttpUtility.HtmlDecode(SR.ReadToEnd)
    '            End Using
    '            Return Data
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '    Private Function ReadAllHTMLEx(URL As String) As String
    '        Try
    '            Dim Data As String
    '            Data = New WebClient().DownloadString(URL)
    '            Return Data
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '    Private Sub ValidarRUC(ByVal NroDocumento As String)
    '        Try
    '            NroDocumento = NroDocumento.Trim
    '            If Not IsNumeric(NroDocumento) Then
    '                Throw New Exception("El documento ingresado no contiene el formato correcto.")
    '            End If
    '            If NroDocumento.Length = 11 Then
    '                Dim Factores = {5, 4, 3, 2, 7, 6, 5, 4, 3, 2}, Resultado%
    '                For i = 0 To 9
    '                    Dim Valor% = Mid(NroDocumento, i + 1, 1)
    '                    Factores(i) = Valor * Factores(i)
    '                Next
    '                Resultado = 11 - (Factores.Sum Mod 11)
    '                Resultado = IIf(Resultado = 10, 0, IIf(Resultado = 11, 1, Resultado))
    '                If Resultado > 11 Then Resultado = Right(Resultado, 1)
    '                If Resultado <> Right(NroDocumento, 1) Then
    '                    Throw New Exception("El Número de RUC es incorrecto.")
    '                End If
    '            Else
    '                Throw New Exception("La cantidad de dígitos tiene que ser igual a 11.")
    '            End If
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Sub
End Class
