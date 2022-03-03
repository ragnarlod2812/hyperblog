<%@Language=VBScript %>
<%  Option Explicit
    
    dim nDocumento, nSiniestro, TipoDoc
    dim intValorRetornado
    dim oRs2, oRs1,sObj3, NRDBobj3, logOut
    nDocumento = Request.QueryString("nDocumento")
    nSiniestro = Request.QueryString("nSiniestro")
    TipoDoc = Request.QueryString("TipoDoc")

    call LoginWs

if oRs1("success") = "Falso" then
        intValorRetornado = 1
elseif oRs1("success") = "No Paso Invoke" THEN
    intValorRetornado = 1
elseif oRs1("success") = "Verdadero" THEN
    if oRs1("IsAutenticated") = "true" then
    Session("IsAutenticated") = oRs1("IsAutenticated")
    Session("SessionID") = oRs1("SessionID")
    Session("NoAutenticatedReason") = oRs1("NoAutenticatedReason")
    Session("Username") = oRs1("Username")
    end if
end if

function LoginWs
    sObj3               = "NR_ConsGrales.GetDatPolizas"
Set NRDBobj3 = Server.CreateObject(sObj3)
set oRs1 = NRDBobj3.WSDocumentosLogin

                                                                                                                                                                      
  set sObj3  = nothing                                                                                                                                                                                                                     
  set NRDBobj3 =nothing                                                                                                                                                                                                                    
                                                                                                                                                                                                                                           
end function                                                                                                                                                                                                                               
                                                                                                                                                                                                                                           
function LogoutWs                                                                                                                                                                                                                          
    sObj3               = "NR_ConsGrales.GetDatPolizas"                                                                                                                                                                                    
Set NRDBobj3 = Server.CreateObject(sObj3)                                                                                                                                                                                                  
    logOut = NRDBobj3.WSDocumentosLogout(Session("SessionID"))                                                                                                                                                                             
                                                                                                                                                                                                                                           
                                                                                                                                                                                                                                           
  set sObj3  = nothing                                                                                                                                                                                                                     
  set NRDBobj3 =nothing                                                                                                                                                                                                                    
                                                                                                                                                                                                                                           
end function                                                                                                                                                                                                                               
                                                                                                                                                                                                                                           
    call ObtieneDocumento                                                                                                                                                                                                                  
    call LogoutWs                                                                                                                                                                                                                          
                                                                                                                                                                                                                                           

  If intValorRetornado = 0 Then

     dim tmpDoc,nodeB64,base64String, Extension, Base64Decodex

     base64String = oRs2("Base64")
     Extension = oRs2("Extension")

    On Error Resume Next
    Set tmpDoc = Server.CreateObject("MSXML2.DomDocument")
    Set nodeB64 = tmpDoc.CreateElement("bs64")
    nodeB64.DataType = "bin.base64" 
    Base64Decodex = Base64Decode(base64String)
    nodeB64.Text = Mid(Base64Decodex, InStr(Base64Decodex, ",") + 1) 
    
    If Err.Number <> 0 Then
        Base64Decodex = Base64ToBSTR(base64String)    
       With Response
        .Clear
        .ContentType ="application/octet-stream"
        .Charset = ""
        .AddHeader "Content-Disposition", "attachment; filename="& TipoDoc & "_" & nSiniestro & "." & Extension    
        .BinaryWrite Base64Decodex'nodeB64.NodeTypedValue 'get bytes and write
        .Flush
        .End
       End With
    else
        With Response
        .Clear
        .ContentType ="application/octet-stream"
        .Charset = ""
        .AddHeader "Content-Disposition", "attachment; filename="& TipoDoc & "_" & nSiniestro & "." & Extension    
        .BinaryWrite nodeB64.NodeTypedValue 'get bytes and write
        .Flush
        .End
        End With
    End If
end if

Function Base64ToBSTR(sBase64)
dim i, ByteArray,w1,w2,w3,w4
    For i = 1 To Len(sBase64) Step 4
        w1 = FindPos(Mid(sBase64, i, 1))
        w2 = FindPos(Mid(sBase64, i + 1, 1))
        w3 = FindPos(Mid(sBase64, i + 2, 1))
        w4 = FindPos(Mid(sBase64, i + 3, 1))
        If (w2 >= 0) Then ByteArray = ByteArray & chrB((w1 * 4 + Int(w2 / 16)) And 255)
        If (w3 >= 0) Then ByteArray = ByteArray & chrB((w2 * 16 + Int(w3 / 4)) And 255)
        If (w4 >= 0) Then ByteArray = ByteArray & chrB((w3 * 64 + w4) And 255)
    Next
    Base64ToBSTR = ByteArray
End Function

Function FindPos(sChar)
Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    If (Len(sChar) = 0) Then
        FindPos = -1
    Else
        FindPos = InStr(Base64, sChar) - 1
    End If
End Function
 
Function Base64Decode(ByVal base64String)

  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin

  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")

  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    nGroup = Hex(nGroup)
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

function ObtieneDocumento()

                sObj3           = "NR_ConsGrales.GetDatPolizas"
                Set NRDBobj3 = Server.CreateObject(sObj3)

        set NRDBobj3 = Server.CreateObject(sObj3)
        set  oRs2 = NRDBobj3.GetDocumentosSiniestros(nDocumento,"2", Session("SessionID")) 
        if oRs2 is nothing then
                intValorRetornado = 1
        else
            intValorRetornado = 0
        end if

        set sObj3  = nothing
        set NRDBobj3 =nothing

end function

%>

<html>
<head>
    <title>Consorcio Seguros</title>
<script language="JavaScript">
    function CerrarVentana() {
        window.close();
    } 
</script>
</head>
<body onload="CerrarVentana()" onfocus="window.close()">

</body>
</html>
