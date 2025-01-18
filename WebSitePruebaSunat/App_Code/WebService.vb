Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports GEN.RUC

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class WebService
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function Genconsruc(ruc As String, tipo As SunatBC.EnumTipoDato) As SunatEN
        Dim objE As New SunatEN
        objE.out_band = 0
        Try
            Dim clases As New SunatBC
            objE = clases.ConsultaSUNAT(ruc, tipo)
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

    <WebMethod()>
    Public Function sendBill(objRequestEnvio As RequestEnvio) As ResponseStatus
        Dim objE As New ResponseStatus
        'objE.out_band = 0
        Try
            Dim clases As New SunatBC
            objE = clases.sendBill(objRequestEnvio)
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objE.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objE.mensaje = ex.Message
            End If
            'objE.out_band = 1
        End Try
        Return objE
    End Function

    <WebMethod()>
    Public Function sendBillTest(fileName As String, type_doc_id As String, seri_doc_id As String, corr_doc_id As String, Username As String, Password As String, doi As String, client_id As String, client_secret As String, arcGreZip As String, hashZip As String) As ResponseStatus
        Dim objE As New ResponseStatus
        'objE.out_band = 0
        Try
            Dim clases As New SunatBC
            objE = clases.sendBillTest(fileName, type_doc_id, seri_doc_id, corr_doc_id, Username, Password, doi, client_id, client_secret, arcGreZip, hashZip)
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                objE.mensaje = ex.Message & " " & ex.InnerException.Message
            Else
                objE.mensaje = ex.Message
            End If
            'objE.out_band = 1
        End Try
        Return objE
    End Function

    <WebMethod()>
    Public Function getStatusCdr(objRequestEnvio As RequestEnvio) As ResponseStatus
        Dim objE As New ResponseStatus
        objE.out_band = 0
        Try
            Dim clases As New SunatBC
            objE = clases.getStatusCdr(objRequestEnvio)
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

    <WebMethod()>
    Public Function getStatusCdrTest(numTicket As String, Username As String, Password As String, doi As String, client_id As String, client_secret As String) As ResponseStatus
        Dim objE As New ResponseStatus
        objE.out_band = 0
        Try
            Dim clases As New SunatBC
            objE = clases.getStatusCdrTest(numTicket, Username, Password, doi, client_id, client_secret)
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
End Class