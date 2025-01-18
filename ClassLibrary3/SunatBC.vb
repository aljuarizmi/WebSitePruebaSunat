Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Web
Imports HtmlAgilityPack
Imports Tesseract
Imports System.Drawing
Imports System.Text
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports RestSharp.RestClient
Imports System.Runtime.Serialization.Json
Imports System.Runtime.Serialization
Imports System.Security.Cryptography
Imports BUND.GRE

Public Class SunatBC
    Enum EnumTipoDato
        RUC = 1
        DNI = 2
        TipoCambio = 3
    End Enum
    Public Function ConsultaSUNAT(DATO As String, tipo As EnumTipoDato) As SunatEN
        Dim objE As New SunatEN
        objE.mensaje = ""
        Try
            If ValidarRUC(DATO, tipo, objE.mensaje) Then
                Dim miCookie As New CookieContainer
                'Dim TextoCache, URL As String
                Dim URL As String = ""
                'If tipo.Trim = "" Then
                '    tipo = "1"
                'End If
                If tipo <> EnumTipoDato.RUC And tipo <> EnumTipoDato.DNI And tipo <> EnumTipoDato.TipoCambio Then
                    objE.mensaje = "Solo se permite consultas para tipo 1 (RUC), 2 (DNI) Y 3 (Tipo Cambio)."
                    objE.out_band = 1
                    Return objE
                End If
                Dim DATA As String
                '''URL = $"http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRuc&nroRuc={RUC}&codigo={TextoCache}&tipdoc=1"
                If tipo = EnumTipoDato.RUC Then
                    'Consulta de RUC
                    URL = $"https://api.apis.net.pe/v1/ruc?numero=" & DATO
                End If
                If tipo = EnumTipoDato.DNI Then
                    URL = $"https://api.apis.net.pe/v1/dni?numero=" & DATO
                End If
                If tipo = EnumTipoDato.TipoCambio Then
                    URL = $"https://api.apis.net.pe/v1/tipo-cambio-sunat?fecha=" & DATO
                End If
                DATA = ReadAllHTML(URL, miCookie)
                Dim objsunat As New SunatEN
                objsunat = JsonConvert.DeserializeObject(Of SunatEN)(DATA)
                Return objsunat
            Else
                objE.out_band = 1
            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objE.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objE.mensaje = ex.Message
            End If
            objE.out_band = 1
        End Try
        Return objE
    End Function
    Private Function ReadAllHTML(URL As String, Optional Cookie As CookieContainer = Nothing) As String
        Try
            Dim Data As String
            Dim wrURL As HttpWebRequest = WebRequest.Create(URL)
            wrURL.CookieContainer = Cookie
            Using SR As New StreamReader(wrURL.GetResponse.GetResponseStream, System.Text.Encoding.Default)
                Data = HttpUtility.HtmlDecode(SR.ReadToEnd)
            End Using
            Return Data
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidarRUC(ByRef NroDocumento As String, tipo As EnumTipoDato, ByRef mensaje As String) As Boolean
        Dim blnOK As Boolean = True
        Try
            NroDocumento = NroDocumento.Trim
            If tipo = EnumTipoDato.RUC Or tipo = EnumTipoDato.DNI Then
                If Not IsNumeric(NroDocumento) Then
                    mensaje = "El documento ingresado no contiene el formato correcto."
                    blnOK = False
                End If
            End If
            If tipo = EnumTipoDato.TipoCambio Then
                If Not IsDate(NroDocumento) Then
                    mensaje = "La fecha NO contiene el formato correcto yyyy-MM-dd."
                    blnOK = False
                Else
                    Dim fecha As Date = CDate(NroDocumento)
                    NroDocumento = fecha.Year.ToString & "-" & fecha.Month.ToString.PadLeft(2, "0") & "-" & fecha.Day.ToString.PadLeft(2, "0")
                End If
            End If
            If tipo = EnumTipoDato.RUC Then
                If NroDocumento.Length = 11 Then
                    Dim Factores = {5, 4, 3, 2, 7, 6, 5, 4, 3, 2}, Resultado%
                    For i = 0 To 9
                        Dim Valor% = Mid(NroDocumento, i + 1, 1)
                        Factores(i) = Valor * Factores(i)
                    Next
                    Resultado = 11 - (Factores.Sum Mod 11)
                    Resultado = IIf(Resultado = 10, 0, IIf(Resultado = 11, 1, Resultado))
                    If Resultado > 11 Then Resultado = Right(Resultado, 1)
                    If Resultado <> Right(NroDocumento, 1) Then
                        mensaje = "El Número de RUC es incorrecto."
                        blnOK = False
                    End If
                Else
                    mensaje = "La cantidad de dígitos tiene que ser igual a 11."
                    blnOK = False
                End If
            End If
            If tipo = EnumTipoDato.DNI Then
                If NroDocumento.Length <> 8 Then
                    mensaje = "La cantidad de dígitos tiene que ser igual a 8."
                    blnOK = False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return blnOK
    End Function

    Private Function F_LeerCaptcha(ByRef miCookie As CookieContainer) As String
        Dim TextoCache As String
        Try
            Dim ImageCache As System.Drawing.Image
            Dim URL As String
            URL = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image"
            Dim wrImage As HttpWebRequest = WebRequest.Create(URL)
            wrImage.CookieContainer = miCookie
            ImageCache = Image.FromStream(wrImage.GetResponse.GetResponseStream)
            Dim ruta As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly.CodeBase)
            ruta = Path.Combine(ruta, "tessdata")
            ruta = ruta.Replace("file:\\", "")
            ruta = ruta.Replace("file:\", "")
            Using engine As New TesseractEngine(ruta, "eng", EngineMode.Default)
                Dim page As Page = engine.Process(New Bitmap(ImageCache))
                TextoCache = page.GetText().ToUpper
                TextoCache = Regex.Replace(TextoCache, "[^a-zA-Z]+", String.Empty)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return TextoCache
    End Function

    Public Function sendBill(objRequestEnvio As RequestEnvio) As ResponseStatus
        Dim objResponseStatus As New ResponseStatus
        objResponseStatus.mensaje = ""
        Try
            Dim msgError As String = ""
            Dim clientessol = New RestSharp.RestClient()
            Dim requestclientessol = New RestRequest("https://api-seguridad.sunat.gob.pe/v1/clientessol/" & objRequestEnvio.client_id & "/oauth2/token/", Method.Post)
            requestclientessol.AddHeader("Content-Type", "application/x-www-form-urlencoded")
            'request.AddHeader("Cookie", "TS019e7fc2=014dc399cb267fc8a442af4917ebbb259353545b162ef7387d93e64a91107371a372395932bd6ebf05e5a4d9c8e02c20e111d21d7c")
            requestclientessol.AddParameter("grant_type", "password")
            requestclientessol.AddParameter("scope", "https://api-cpe.sunat.gob.pe")
            'requestclientessol.AddParameter("client_id", "1148c8da-bd67-4ff8-ac1f-3fa014873463")
            'requestclientessol.AddParameter("client_secret", "TQoUyrrXYOijNYKV0JctHA==")
            requestclientessol.AddParameter("client_id", objRequestEnvio.client_id)
            requestclientessol.AddParameter("client_secret", objRequestEnvio.client_secret)
            requestclientessol.AddParameter("username", objRequestEnvio.Doi & objRequestEnvio.Username)
            requestclientessol.AddParameter("password", objRequestEnvio.Password)
            Dim responseclientessol As RestResponse = clientessol.Execute(requestclientessol)
            Dim objToken As New Token
            objToken = JsonConvert.DeserializeObject(Of Token)(responseclientessol.Content)
            If objToken IsNot Nothing Then
                If objToken.access_token IsNot Nothing Then
                    If objToken.access_token.Trim <> "" Then
                        Dim contribuyente = New RestSharp.RestClient()
                        Dim requestcontribuyente = New RestRequest("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/" & objRequestEnvio.Doi & "-" & objRequestEnvio.type_doc_id & "-" & objRequestEnvio.seri_doc_id & "-" & objRequestEnvio.corr_doc_id, Method.Post)
                        requestcontribuyente.AddHeader("Authorization", "Bearer " & objToken.access_token.Trim)
                        requestcontribuyente.AddHeader("Content-Type", "application/json")
                        Dim body As String = "{
                    " +
                    "    ""archivo"" : {
                    " +
                    "        ""nomArchivo"" : """ & objRequestEnvio.fileName & """,
                    " +
                    "        ""arcGreZip"" : """ & ConvertBase64(objRequestEnvio.contentFile) & """,
                    " +
                    "        ""hashZip"" : """ & ConvertSha256(objRequestEnvio.contentFile) & """
                    " +
                    "    }}"
                        requestcontribuyente.AddParameter("application/json", body, ParameterType.RequestBody)
                        Dim responsecontribuyente As RestResponse = contribuyente.Execute(requestcontribuyente)
                        If responsecontribuyente IsNot Nothing Then
                            If responsecontribuyente.Content IsNot Nothing Then
                                If responsecontribuyente.Content.Trim <> "" Then
                                    'Dim objResponse As Response = JsonDeserialize(Of Response)(jsonstring)
                                    Dim objResponse As Response = JsonDeserialize(Of Response)(responsecontribuyente.Content)
                                    If objResponse IsNot Nothing Then
                                        If objResponse.numTicket IsNot Nothing Then
                                            If objResponse.numTicket.Trim <> "" Then
                                                'Obtenemos el ticket y luego consultamos el estado del proceso
                                                Dim contribuyenteConsulta = New RestSharp.RestClient()
                                                Dim requestcontribuyenteConsulta = New RestRequest("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/envios/" & objResponse.numTicket.Trim, Method.Get)
                                                objResponseStatus.numTicket = objResponse.numTicket
                                                requestcontribuyenteConsulta.AddHeader("Authorization", "Bearer " & objToken.access_token.Trim)
                                                Dim responsecontribuyenteConsulta As RestResponse = contribuyenteConsulta.Execute(requestcontribuyenteConsulta)
                                                If responsecontribuyenteConsulta IsNot Nothing Then
                                                    If responsecontribuyenteConsulta.Content IsNot Nothing Then
                                                        If responsecontribuyenteConsulta.Content.Trim <> "" Then
                                                            Dim jsonstring As String = responsecontribuyenteConsulta.Content.Replace(":{", ":[{").Replace("},", "}],")
                                                            Dim objResponseConsulta As ResponseEnvio = JsonDeserialize(Of ResponseEnvio)(jsonstring)
                                                            If objResponseConsulta IsNot Nothing Then
                                                                If objResponseConsulta.codRespuesta IsNot Nothing Then
                                                                    If objResponseConsulta.codRespuesta = "0" Then
                                                                        objResponseStatus.codRespuesta = 0
                                                                        objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                                        If objResponseConsulta.indCdrGenerado = "1" Then
                                                                            If objResponseConsulta.arcCdr.Trim <> "" Then
                                                                                objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                                            End If
                                                                        End If
                                                                        objResponseStatus.out_band = 0
                                                                    Else
                                                                        If objResponseConsulta.codRespuesta = "99" Then
                                                                            objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                                            objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                                            If objResponseConsulta.indCdrGenerado = "1" Then
                                                                                If objResponseConsulta.arcCdr.Trim <> "" Then
                                                                                    objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                                                End If
                                                                            End If
                                                                            If objResponseConsulta.error IsNot Nothing Then
                                                                                Dim lstErrors As List(Of [error]) = Nothing
                                                                                lstErrors = objResponseConsulta.error
                                                                                If lstErrors.Count > 0 Then
                                                                                    objResponseStatus.mensaje = lstErrors(0).desError
                                                                                    objResponseStatus.numError = lstErrors(0).numError
                                                                                End If
                                                                            End If
                                                                            objResponseStatus.out_band = 0
                                                                        End If
                                                                        If objResponseConsulta.codRespuesta = "98" Then
                                                                            objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                                            objResponseStatus.mensaje = "Envío en proceso"
                                                                            objResponseStatus.out_band = 0
                                                                        End If
                                                                    End If
                                                                Else
                                                                    msgError = "Código de error " & objResponseConsulta.status & ": " & objResponseConsulta.message
                                                                    objResponseStatus.codRespuesta = objResponseConsulta.status
                                                                    objResponseStatus.mensaje = objResponseConsulta.message
                                                                    objResponseStatus.out_band = 1
                                                                End If
                                                            Else
                                                                objResponseStatus.mensaje = "Error B14: No se pudo deserializar la respuesta del ticket."
                                                                objResponseStatus.out_band = 1
                                                            End If
                                                        Else
                                                            objResponseStatus.mensaje = "Error B13: El contenido de la respuesta del ticket es vacío."
                                                            objResponseStatus.out_band = 1
                                                        End If
                                                    Else
                                                        objResponseStatus.mensaje = "Error B12: El contenido de la respuesta del ticket es NULO."
                                                        objResponseStatus.out_band = 1
                                                    End If
                                                Else
                                                    'El objeto responsecontribuyenteConsulta es NULO
                                                    objResponseStatus.mensaje = "Error B11: No se pudo obtener la respuesta del ticket " & objResponse.numTicket
                                                    objResponseStatus.out_band = 1
                                                End If
                                            Else
                                                'EL objeto objResponse.numTicket es una cadena vacia
                                                objResponseStatus.mensaje = "Error B10: No se pudo obtener el ticket."
                                                objResponseStatus.out_band = 1
                                            End If
                                        Else
                                            'Verificamos el campos status
                                            If objResponse.status IsNot Nothing Then
                                                msgError = "Código de error " & objResponse.status & ": " & objResponse.message
                                                objResponseStatus.codRespuesta = objResponse.status
                                                objResponseStatus.mensaje = objResponse.message
                                                objResponseStatus.out_band = 1
                                            Else
                                                'EL objeto objResponse.numTicket es NULO
                                                'Como es NULO entonces verificamos si el json trae error
                                                Dim objResponseError As ResponseError = JsonDeserialize(Of ResponseError)(responsecontribuyente.Content)
                                                Dim lstErrors As List(Of errors) = Nothing
                                                If objResponseError IsNot Nothing Then
                                                    If objResponseError.cod.Trim <> "" Then
                                                        'Obtenemos el codigo de error
                                                        msgError = "Código de error " & objResponseError.cod & ": " & objResponseError.msg
                                                        objResponseStatus.codRespuesta = objResponseError.cod
                                                        objResponseStatus.mensaje = objResponseError.msg
                                                        objResponseStatus.out_band = 1
                                                        If objResponseError.errors IsNot Nothing Then
                                                            lstErrors = objResponseError.errors
                                                            If lstErrors.Count > 0 Then
                                                                'Throw New Exception("Código de error " & lstErrors(0).cod & ": " & lstErrors(0).msg)
                                                                msgError = "Código de error " & lstErrors(0).cod & ": " & lstErrors(0).msg
                                                                objResponseStatus.codRespuesta = lstErrors(0).cod
                                                                objResponseStatus.mensaje = lstErrors(0).msg
                                                            End If
                                                        End If
                                                    Else
                                                        objResponseStatus.mensaje = "Error B09: El código de error esta vacío."
                                                        objResponseStatus.out_band = 1
                                                    End If
                                                Else
                                                    'EL objeto objResponseError500 es NULO
                                                    objResponseStatus.mensaje = "Error B08: No se pudo deserializar la respuesta del error."
                                                    objResponseStatus.out_band = 1
                                                End If
                                            End If

                                        End If
                                    Else
                                        'EL objeto objResponse es NULO
                                        objResponseStatus.mensaje = "Error B07: No se pudo deserializar la respuesta del envio."
                                        objResponseStatus.out_band = 1
                                    End If
                                Else
                                    'EL objeto responsecontribuyente.Content es una cadena vacia
                                    objResponseStatus.mensaje = "Error B06: El contenido de la respuesta es vacio."
                                    objResponseStatus.out_band = 1
                                End If
                            Else
                                'EL objeto responsecontribuyente.Content es NULO
                                objResponseStatus.mensaje = "Error B05: El contenido de la respuesta es nulo."
                                objResponseStatus.out_band = 1
                            End If
                        Else
                            'El objeto responsecontribuyente es NULO
                            objResponseStatus.mensaje = "Error B04: No se pudo obtener la respuesta del envío."
                            objResponseStatus.out_band = 1
                        End If
                    Else
                        'El token esta vacio
                        objResponseStatus.mensaje = "Error B03: El token esta vacio."
                        objResponseStatus.out_band = 1
                    End If
                Else
                    'El token es nulo
                    msgError = "El token es nulo."
                    objResponseStatus.mensaje = "Error B02: El token es nulo."
                    objResponseStatus.out_band = 1
                    If objToken.error_description IsNot Nothing Then
                        If objToken.error_description.Trim <> "" Then
                            msgError = objToken.error_description.Trim
                            objResponseStatus.mensaje = "Error B02: " & objToken.error_description.Trim
                        End If
                    End If
                End If
            Else
                objResponseStatus.out_band = 1
                objResponseStatus.mensaje = "Error B01: No se pudo deserializar el contenido del token."
            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objResponseStatus.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objResponseStatus.mensaje = ex.Message
            End If
            objResponseStatus.out_band = 1
        End Try
        Return objResponseStatus
    End Function

    Public Function sendBillTest(fileName As String, type_doc_id As String, seri_doc_id As String, corr_doc_id As String, Username As String, Password As String, doi As String, client_id As String, client_secret As String, arcGreZip As String, hashZip As String) As ResponseStatus
        Dim objResponseStatus As New ResponseStatus
        objResponseStatus.mensaje = ""
        Try
            Dim msgError As String = ""
            Dim clientessol = New RestSharp.RestClient()
            Dim requestclientessol = New RestRequest("https://api-seguridad.sunat.gob.pe/v1/clientessol/" & client_id & "/oauth2/token/", Method.Post)
            requestclientessol.AddHeader("Content-Type", "application/x-www-form-urlencoded")
            'request.AddHeader("Cookie", "TS019e7fc2=014dc399cb267fc8a442af4917ebbb259353545b162ef7387d93e64a91107371a372395932bd6ebf05e5a4d9c8e02c20e111d21d7c")
            requestclientessol.AddParameter("grant_type", "password")
            requestclientessol.AddParameter("scope", "https://api-cpe.sunat.gob.pe")
            'requestclientessol.AddParameter("client_id", "1148c8da-bd67-4ff8-ac1f-3fa014873463")
            'requestclientessol.AddParameter("client_secret", "TQoUyrrXYOijNYKV0JctHA==")
            requestclientessol.AddParameter("client_id", client_id)
            requestclientessol.AddParameter("client_secret", client_secret)
            requestclientessol.AddParameter("username", doi & Username)
            requestclientessol.AddParameter("password", Password)
            Dim responseclientessol As RestResponse = clientessol.Execute(requestclientessol)
            Dim objToken As New Token
            objToken = JsonConvert.DeserializeObject(Of Token)(responseclientessol.Content)
            If objToken IsNot Nothing Then
                If objToken.access_token IsNot Nothing Then
                    If objToken.access_token.Trim <> "" Then
                        Dim contribuyente = New RestSharp.RestClient()
                        Dim requestcontribuyente = New RestRequest("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/" & doi & "-" & type_doc_id & "-" & seri_doc_id & "-" & corr_doc_id, Method.Post)
                        requestcontribuyente.AddHeader("Authorization", "Bearer " & objToken.access_token.Trim)
                        requestcontribuyente.AddHeader("Content-Type", "application/json")
                        Dim body As String = "{
                    " +
                    "    ""archivo"" : {
                    " +
                    "        ""nomArchivo"" : """ & fileName & """,
                    " +
                    "        ""arcGreZip"" : """ & arcGreZip & """,
                    " +
                    "        ""hashZip"" : """ & hashZip & """
                    " +
                    "    }}"
                        requestcontribuyente.AddParameter("application/json", body, ParameterType.RequestBody)
                        Dim responsecontribuyente As RestResponse = contribuyente.Execute(requestcontribuyente)
                        If responsecontribuyente IsNot Nothing Then
                            If responsecontribuyente.Content IsNot Nothing Then
                                If responsecontribuyente.Content.Trim <> "" Then
                                    'Dim objResponse As Response = JsonDeserialize(Of Response)(jsonstring)
                                    Dim objResponse As Response = JsonDeserialize(Of Response)(responsecontribuyente.Content)
                                    If objResponse IsNot Nothing Then
                                        If objResponse.numTicket IsNot Nothing Then
                                            If objResponse.numTicket.Trim <> "" Then
                                                'Obtenemos el ticket y luego consultamos el estado del proceso
                                                Dim contribuyenteConsulta = New RestSharp.RestClient()
                                                'Dim contribuyenteConsulta = New RestSharp.RestClient("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/envios/" & objResponse.numTicket.Trim)
                                                Dim requestcontribuyenteConsulta = New RestRequest("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/envios/" & objResponse.numTicket.Trim, Method.Get)
                                                objResponseStatus.numTicket = objResponse.numTicket
                                                requestcontribuyenteConsulta.AddHeader("Authorization", "Bearer " & objToken.access_token.Trim)
                                                Dim responsecontribuyenteConsulta As RestResponse = contribuyenteConsulta.Execute(requestcontribuyenteConsulta)
                                                If responsecontribuyenteConsulta IsNot Nothing Then
                                                    If responsecontribuyenteConsulta.Content IsNot Nothing Then
                                                        If responsecontribuyenteConsulta.Content.Trim <> "" Then
                                                            Dim jsonstring As String = responsecontribuyenteConsulta.Content.Replace(":{", ":[{").Replace("},", "}],")
                                                            Dim objResponseConsulta As ResponseEnvio = JsonDeserialize(Of ResponseEnvio)(jsonstring)
                                                            'Dim js As New System.Web.Script.Serialization.JavaScriptSerializer
                                                            'Dim blogObject = js.Deserialize(Of ResponseEnvio)(responsecontribuyenteConsulta.Content)
                                                            'Dim blogObject As Dynamic = js.Deserialize(Of Dynamic)("")
                                                            'Dim obj3 As New BUND.GRE.ResponseEnvio
                                                            'Dim obj4 As New JsonResponse
                                                            'obj3 = obj4.JsonDeserialize(jsonstring)
                                                            If objResponseConsulta IsNot Nothing Then
                                                                If objResponseConsulta.codRespuesta IsNot Nothing Then
                                                                    If objResponseConsulta.codRespuesta = "0" Then
                                                                        objResponseStatus.codRespuesta = 0
                                                                        objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                                        If objResponseConsulta.indCdrGenerado = "1" Then
                                                                            If objResponseConsulta.arcCdr.Trim <> "" Then
                                                                                objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                                            End If
                                                                        End If
                                                                        objResponseStatus.out_band = 0
                                                                    Else
                                                                        If objResponseConsulta.codRespuesta = "99" Then
                                                                            objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                                            objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                                            If objResponseConsulta.indCdrGenerado = "1" Then
                                                                                If objResponseConsulta.arcCdr.Trim <> "" Then
                                                                                    objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                                                End If
                                                                            End If
                                                                            If objResponseConsulta.error IsNot Nothing Then
                                                                                Dim lstErrors As List(Of [error]) = Nothing
                                                                                lstErrors = objResponseConsulta.error
                                                                                If lstErrors.Count > 0 Then
                                                                                    objResponseStatus.mensaje = lstErrors(0).desError
                                                                                    objResponseStatus.numError = lstErrors(0).numError
                                                                                End If
                                                                            End If
                                                                            objResponseStatus.out_band = 0
                                                                        End If
                                                                        If objResponseConsulta.codRespuesta = "98" Then
                                                                            objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                                            objResponseStatus.mensaje = "Envío en proceso"
                                                                            objResponseStatus.out_band = 0
                                                                        End If
                                                                    End If
                                                                Else
                                                                    msgError = "Código de error " & objResponseConsulta.status & ": " & objResponseConsulta.message
                                                                    objResponseStatus.codRespuesta = objResponseConsulta.status
                                                                    objResponseStatus.mensaje = objResponseConsulta.message
                                                                    objResponseStatus.out_band = 1
                                                                    'Throw New Exception(msgError)
                                                                End If
                                                            Else
                                                                objResponseStatus.mensaje = "Error B14: No se pudo deserializar la respuesta del ticket."
                                                                objResponseStatus.out_band = 1
                                                                'Throw New Exception("No se pudo deserializar la respuesta del ticket.")
                                                            End If
                                                        Else
                                                            objResponseStatus.mensaje = "Error B13: El contenido de la respuesta del ticket es vacío."
                                                            objResponseStatus.out_band = 1
                                                            'Throw New Exception("El contenido de la respuesta del ticket es vacío.")
                                                        End If
                                                    Else
                                                        objResponseStatus.mensaje = "Error B12: El contenido de la respuesta del ticket es NULO."
                                                        objResponseStatus.out_band = 1
                                                        'Throw New Exception("El contenido de la respuesta del ticket es NULO.")
                                                    End If
                                                Else
                                                    'El objeto responsecontribuyenteConsulta es NULO
                                                    objResponseStatus.mensaje = "Error B11: No se pudo obtener la respuesta del ticket " & objResponse.numTicket
                                                    objResponseStatus.out_band = 1
                                                    'Throw New Exception("No se pudo obtener la respuesta del ticket " & objResponse.numTicket)
                                                End If
                                            Else
                                                'EL objeto objResponse.numTicket es una cadena vacia
                                                objResponseStatus.mensaje = "Error B10: No se pudo obtener el ticket."
                                                objResponseStatus.out_band = 1
                                                'Throw New Exception("No se pudo obtener el ticket.")
                                            End If
                                        Else
                                            'Verificamos el campos status
                                            If objResponse.status IsNot Nothing Then
                                                msgError = "Código de error " & objResponse.status & ": " & objResponse.message
                                                objResponseStatus.codRespuesta = objResponse.status
                                                objResponseStatus.mensaje = objResponse.message
                                                objResponseStatus.out_band = 1
                                                'Throw New Exception(msgError)
                                            Else
                                                'EL objeto objResponse.numTicket es NULO
                                                'Como es NULO entonces verificamos si el json trae error
                                                'Dim objResponseError As ResponseError = JsonDeserialize(Of ResponseError)(jsonstring)
                                                Dim objResponseError As ResponseError = JsonDeserialize(Of ResponseError)(responsecontribuyente.Content)
                                                Dim lstErrors As List(Of errors) = Nothing
                                                'Dim objResponseError422 As ResponseError422 = Nothing
                                                'Dim errors As String = ""
                                                'If jsonstring.IndexOf("[") >= 0 And jsonstring.IndexOf("]") >= 0 Then
                                                '    errors = jsonstring.Substring(jsonstring.IndexOf("["), jsonstring.IndexOf("]") - jsonstring.IndexOf("[") + 1)
                                                '    errors = errors.Replace("[", "").Replace("]", "")
                                                '    objResponseError422 = JsonDeserialize(Of ResponseError422)(errors)
                                                'End If
                                                If objResponseError IsNot Nothing Then
                                                    If objResponseError.cod.Trim <> "" Then
                                                        'Obtenemos el codigo de error
                                                        msgError = "Código de error " & objResponseError.cod & ": " & objResponseError.msg
                                                        objResponseStatus.codRespuesta = objResponseError.cod
                                                        objResponseStatus.mensaje = objResponseError.msg
                                                        objResponseStatus.out_band = 1
                                                        If objResponseError.errors IsNot Nothing Then
                                                            lstErrors = objResponseError.errors
                                                            If lstErrors.Count > 0 Then
                                                                'Throw New Exception("Código de error " & lstErrors(0).cod & ": " & lstErrors(0).msg)
                                                                msgError = "Código de error " & lstErrors(0).cod & ": " & lstErrors(0).msg
                                                                objResponseStatus.codRespuesta = lstErrors(0).cod
                                                                objResponseStatus.mensaje = lstErrors(0).msg
                                                            End If
                                                        End If
                                                        'Throw New Exception(msgError)
                                                    Else
                                                        objResponseStatus.mensaje = "Error B09: El código de error esta vacío."
                                                        objResponseStatus.out_band = 1
                                                        'Throw New Exception("El código de error esta vacío.")
                                                    End If
                                                Else
                                                    'EL objeto objResponseError500 es NULO
                                                    objResponseStatus.mensaje = "Error B08: No se pudo deserializar la respuesta del error."
                                                    objResponseStatus.out_band = 1
                                                    'Throw New Exception("No se pudo deserializar la respuesta del error.")
                                                End If
                                            End If

                                        End If
                                    Else
                                        'EL objeto objResponse es NULO
                                        objResponseStatus.mensaje = "Error B07: No se pudo deserializar la respuesta del envio."
                                        objResponseStatus.out_band = 1
                                        'Throw New Exception("No se pudo deserializar la respuesta del envio.")
                                    End If
                                Else
                                    'EL objeto responsecontribuyente.Content es una cadena vacia
                                    objResponseStatus.mensaje = "Error B06: El contenido de la respuesta es vacio."
                                    objResponseStatus.out_band = 1
                                    'Throw New Exception("El contenido de la respuesta es vacio.")
                                End If
                            Else
                                'EL objeto responsecontribuyente.Content es NULO
                                objResponseStatus.mensaje = "Error B05: El contenido de la respuesta es nulo."
                                objResponseStatus.out_band = 1
                                'Throw New Exception("El contenido de la respuesta es nulo.")
                            End If
                        Else
                            'El objeto responsecontribuyente es NULO
                            objResponseStatus.mensaje = "Error B04: No se pudo obtener la respuesta del envío."
                            objResponseStatus.out_band = 1
                            'Throw New Exception("No se pudo obtener la respuesta del envío.")
                        End If
                    Else
                        'El token esta vacio
                        objResponseStatus.mensaje = "Error B03: El token esta vacio."
                        objResponseStatus.out_band = 1
                        'Throw New Exception("El token esta vacio.")
                    End If
                Else
                    'El token es nulo
                    msgError = "El token es nulo."
                    objResponseStatus.mensaje = "Error B02: El token es nulo."
                    objResponseStatus.out_band = 1
                    If objToken.error_description IsNot Nothing Then
                        If objToken.error_description.Trim <> "" Then
                            msgError = objToken.error_description.Trim
                            objResponseStatus.mensaje = "Error B02: " & objToken.error_description.Trim
                        End If
                    End If
                    'Throw New Exception(msgError)
                End If
            Else
                objResponseStatus.out_band = 1
                objResponseStatus.mensaje = "Error B01: No se pudo deserializar el contenido del token."
                'Throw New Exception("No se pudo deserializar el contenido del token.")
            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objResponseStatus.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objResponseStatus.mensaje = ex.Message
            End If
            objResponseStatus.out_band = 1
            'Throw New Exception(objResponseStatus.mensaje)
        End Try
        Return objResponseStatus
    End Function

    Public Function getStatusCdr(objRequestEnvio As RequestEnvio) As ResponseStatus
        Dim objResponseStatus As New ResponseStatus
        objResponseStatus.mensaje = ""
        Try
            Dim clientessol = New RestSharp.RestClient()
            Dim requestclientessol = New RestRequest("https://api-seguridad.sunat.gob.pe/v1/clientessol/" & objRequestEnvio.client_id & "/oauth2/token/", Method.Post)
            requestclientessol.AddHeader("Content-Type", "application/x-www-form-urlencoded")
            requestclientessol.AddParameter("grant_type", "password")
            requestclientessol.AddParameter("scope", "https://api-cpe.sunat.gob.pe")
            requestclientessol.AddParameter("client_id", objRequestEnvio.client_id)
            requestclientessol.AddParameter("client_secret", objRequestEnvio.client_secret)
            requestclientessol.AddParameter("username", objRequestEnvio.Doi & objRequestEnvio.Username)
            requestclientessol.AddParameter("password", objRequestEnvio.Password)
            Dim responseclientessol As RestResponse = clientessol.Execute(requestclientessol)
            Dim objToken As New Token
            objToken = JsonConvert.DeserializeObject(Of Token)(responseclientessol.Content)
            If objToken IsNot Nothing Then
                If objToken.access_token IsNot Nothing Then
                    If objToken.access_token.Trim <> "" Then
                        Dim contribuyenteConsulta = New RestSharp.RestClient()
                        Dim requestcontribuyenteConsulta = New RestRequest("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/envios/" & objRequestEnvio.numTicket.Trim, Method.Get)
                        objResponseStatus.numTicket = objRequestEnvio.numTicket
                        requestcontribuyenteConsulta.AddHeader("Authorization", "Bearer " & objToken.access_token.Trim)
                        Dim responsecontribuyenteConsulta As RestResponse = contribuyenteConsulta.Execute(requestcontribuyenteConsulta)
                        If responsecontribuyenteConsulta IsNot Nothing Then
                            If responsecontribuyenteConsulta.Content IsNot Nothing Then
                                If responsecontribuyenteConsulta.Content.Trim <> "" Then
                                    Dim jsonstring As String = responsecontribuyenteConsulta.Content.Replace(":{", ":[{").Replace("},", "}],")
                                    Dim objResponseConsulta As ResponseEnvio = JsonDeserialize(Of ResponseEnvio)(jsonstring)
                                    If objResponseConsulta IsNot Nothing Then
                                        If objResponseConsulta.codRespuesta IsNot Nothing Then
                                            If objResponseConsulta.codRespuesta = "0" Then
                                                objResponseStatus.codRespuesta = 0
                                                objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                If objResponseConsulta.indCdrGenerado = "1" Then
                                                    If objResponseConsulta.arcCdr.Trim <> "" Then
                                                        objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                    End If
                                                End If
                                                objResponseStatus.out_band = 0
                                            Else
                                                If objResponseConsulta.codRespuesta = "99" Then
                                                    objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                    objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                    If objResponseConsulta.indCdrGenerado = "1" Then
                                                        If objResponseConsulta.arcCdr.Trim <> "" Then
                                                            objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                        End If
                                                    End If
                                                    If objResponseConsulta.error IsNot Nothing Then
                                                        Dim lstErrors As List(Of [error]) = Nothing
                                                        lstErrors = objResponseConsulta.error
                                                        If lstErrors.Count > 0 Then
                                                            objResponseStatus.mensaje = lstErrors(0).desError
                                                            objResponseStatus.numError = lstErrors(0).numError
                                                        End If
                                                    End If
                                                    objResponseStatus.out_band = 0
                                                End If
                                                If objResponseConsulta.codRespuesta = "98" Then
                                                    objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                    objResponseStatus.mensaje = "Envío en proceso"
                                                    objResponseStatus.out_band = 0
                                                End If
                                            End If
                                        Else
                                            'msgError = "Código de error " & objResponseConsulta.status & ": " & objResponseConsulta.message
                                            objResponseStatus.codRespuesta = objResponseConsulta.status
                                            objResponseStatus.mensaje = objResponseConsulta.message
                                            objResponseStatus.out_band = 1
                                        End If
                                    Else
                                        objResponseStatus.mensaje = "Error B07: No se pudo deserializar la respuesta del ticket."
                                        objResponseStatus.out_band = 1
                                    End If
                                Else
                                    objResponseStatus.mensaje = "Error B06: El contenido de la respuesta del ticket es vacío."
                                    objResponseStatus.out_band = 1
                                End If
                            Else
                                objResponseStatus.mensaje = "Error B05: El contenido de la respuesta del ticket es NULO."
                                objResponseStatus.out_band = 1
                            End If
                        Else
                            'El objeto responsecontribuyenteConsulta es NULO
                            objResponseStatus.mensaje = "Error B04: No se pudo obtener la respuesta del ticket " & objRequestEnvio.numTicket
                            objResponseStatus.out_band = 1
                        End If
                    Else
                        'El token esta vacio
                        objResponseStatus.mensaje = "Error B03: El token esta vacio."
                        objResponseStatus.out_band = 1
                    End If
                Else
                    'El token es nulo
                    'msgError = "El token es nulo."
                    objResponseStatus.mensaje = "Error B02: El token es nulo."
                    objResponseStatus.out_band = 1
                    If objToken.error_description IsNot Nothing Then
                        If objToken.error_description.Trim <> "" Then
                            'msgError = objToken.error_description.Trim
                            objResponseStatus.mensaje = "Error B02: " & objToken.error_description.Trim
                        End If
                    End If
                End If
            Else
                objResponseStatus.out_band = 1
                objResponseStatus.mensaje = "Error B01: No se pudo deserializar el contenido del token."
            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objResponseStatus.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objResponseStatus.mensaje = ex.Message
            End If
            objResponseStatus.out_band = 1
        End Try
        Return objResponseStatus
    End Function

    Public Function getStatusCdrTest(numTicket As String, Username As String, Password As String, doi As String, client_id As String, client_secret As String) As ResponseStatus
        Dim objResponseStatus As New ResponseStatus
        objResponseStatus.mensaje = ""
        Try
            Dim clientessol = New RestSharp.RestClient()
            Dim requestclientessol = New RestRequest("https://api-seguridad.sunat.gob.pe/v1/clientessol/" & client_id & "/oauth2/token/", Method.Post)
            requestclientessol.AddHeader("Content-Type", "application/x-www-form-urlencoded")
            requestclientessol.AddParameter("grant_type", "password")
            requestclientessol.AddParameter("scope", "https://api-cpe.sunat.gob.pe")
            requestclientessol.AddParameter("client_id", client_id)
            requestclientessol.AddParameter("client_secret", client_secret)
            requestclientessol.AddParameter("username", doi & Username)
            requestclientessol.AddParameter("password", Password)
            Dim responseclientessol As RestResponse = clientessol.Execute(requestclientessol)
            Dim objToken As New Token
            objToken = JsonConvert.DeserializeObject(Of Token)(responseclientessol.Content)
            If objToken IsNot Nothing Then
                If objToken.access_token IsNot Nothing Then
                    If objToken.access_token.Trim <> "" Then
                        Dim contribuyenteConsulta = New RestSharp.RestClient()
                        Dim requestcontribuyenteConsulta = New RestRequest("https://api-cpe.sunat.gob.pe/v1/contribuyente/gem/comprobantes/envios/" & numTicket.Trim, Method.Get)
                        objResponseStatus.numTicket = numTicket
                        requestcontribuyenteConsulta.AddHeader("Authorization", "Bearer " & objToken.access_token.Trim)
                        Dim responsecontribuyenteConsulta As RestResponse = contribuyenteConsulta.Execute(requestcontribuyenteConsulta)
                        If responsecontribuyenteConsulta IsNot Nothing Then
                            If responsecontribuyenteConsulta.Content IsNot Nothing Then
                                If responsecontribuyenteConsulta.Content.Trim <> "" Then
                                    Dim jsonstring As String = responsecontribuyenteConsulta.Content.Replace(":{", ":[{").Replace("},", "}],")
                                    Dim objResponseConsulta As ResponseEnvio = JsonDeserialize(Of ResponseEnvio)(jsonstring)
                                    If objResponseConsulta IsNot Nothing Then
                                        If objResponseConsulta.codRespuesta IsNot Nothing Then
                                            If objResponseConsulta.codRespuesta = "0" Then
                                                objResponseStatus.codRespuesta = 0
                                                objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                If objResponseConsulta.indCdrGenerado = "1" Then
                                                    If objResponseConsulta.arcCdr.Trim <> "" Then
                                                        objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                    End If
                                                End If
                                                objResponseStatus.out_band = 0
                                            Else
                                                If objResponseConsulta.codRespuesta = "99" Then
                                                    objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                    objResponseStatus.indCdrGenerado = objResponseConsulta.indCdrGenerado
                                                    If objResponseConsulta.indCdrGenerado = "1" Then
                                                        If objResponseConsulta.arcCdr.Trim <> "" Then
                                                            objResponseStatus.arcCdr = ConvertFromBase64(objResponseConsulta.arcCdr)
                                                        End If
                                                    End If
                                                    If objResponseConsulta.error IsNot Nothing Then
                                                        Dim lstErrors As List(Of [error]) = Nothing
                                                        lstErrors = objResponseConsulta.error
                                                        If lstErrors.Count > 0 Then
                                                            objResponseStatus.mensaje = lstErrors(0).desError
                                                            objResponseStatus.numError = lstErrors(0).numError
                                                        End If
                                                    End If
                                                    objResponseStatus.out_band = 0
                                                End If
                                                If objResponseConsulta.codRespuesta = "98" Then
                                                    objResponseStatus.codRespuesta = objResponseConsulta.codRespuesta
                                                    objResponseStatus.mensaje = "Envío en proceso"
                                                    objResponseStatus.out_band = 0
                                                End If
                                            End If
                                        Else
                                            'msgError = "Código de error " & objResponseConsulta.status & ": " & objResponseConsulta.message
                                            objResponseStatus.codRespuesta = objResponseConsulta.status
                                            objResponseStatus.mensaje = objResponseConsulta.message
                                            objResponseStatus.out_band = 1
                                        End If
                                    Else
                                        objResponseStatus.mensaje = "Error B07: No se pudo deserializar la respuesta del ticket."
                                        objResponseStatus.out_band = 1
                                    End If
                                Else
                                    objResponseStatus.mensaje = "Error B06: El contenido de la respuesta del ticket es vacío."
                                    objResponseStatus.out_band = 1
                                End If
                            Else
                                objResponseStatus.mensaje = "Error B05: El contenido de la respuesta del ticket es NULO."
                                objResponseStatus.out_band = 1
                            End If
                        Else
                            'El objeto responsecontribuyenteConsulta es NULO
                            objResponseStatus.mensaje = "Error B04: No se pudo obtener la respuesta del ticket " & numTicket
                            objResponseStatus.out_band = 1
                        End If
                    Else
                        'El token esta vacio
                        objResponseStatus.mensaje = "Error B03: El token esta vacio."
                        objResponseStatus.out_band = 1
                    End If
                Else
                    'El token es nulo
                    'msgError = "El token es nulo."
                    objResponseStatus.mensaje = "Error B02: El token es nulo."
                    objResponseStatus.out_band = 1
                    If objToken.error_description IsNot Nothing Then
                        If objToken.error_description.Trim <> "" Then
                            'msgError = objToken.error_description.Trim
                            objResponseStatus.mensaje = "Error B02: " & objToken.error_description.Trim
                        End If
                    End If
                End If
            Else
                objResponseStatus.out_band = 1
                objResponseStatus.mensaje = "Error B01: No se pudo deserializar el contenido del token."
            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objResponseStatus.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objResponseStatus.mensaje = ex.Message
            End If
            objResponseStatus.out_band = 1
        End Try
        Return objResponseStatus
    End Function
    Public Shared Function JsonDeserialize(Of T)(ByVal jsonString As String) As T
        Dim obj As T = Activator.CreateInstance(Of T)
        Dim ms As MemoryStream = New MemoryStream(Encoding.Unicode.GetBytes(jsonString))
        Dim serializer As New DataContractJsonSerializer(obj.GetType())
        obj = serializer.ReadObject(ms)
        ms.Close()
        Return obj
    End Function

    Public Shared Function ConvertSha256(contentFile As Byte()) As String
        Dim mySHA256 As SHA256 = SHA256Managed.Create()
        Dim archivohash As Byte()
        Dim sBuilder As New StringBuilder
        archivohash = mySHA256.ComputeHash(contentFile)
        For i As Int32 = 0 To archivohash.Length - 1
            sBuilder.Append(archivohash(i).ToString("x2"))
        Next
        Return sBuilder.ToString()
    End Function

    Public Shared Function ConvertBase64(contentFile As Byte()) As String
        Return Convert.ToBase64String(contentFile)
    End Function

    Public Shared Function ConvertFromBase64(arcCdr As String) As Byte()
        Return Convert.FromBase64String(arcCdr)
    End Function
End Class