<%@ Language=VBScript %>

<html>
<head>
<title>Priore CryptoWEB ASP Sample</title>
</head>
<body>

<%
	Response.Expires = -1
	
	on error resume next
	' create object cryptoweb
	set oCrypt = createobject("CryptoWEB.Functions")
	if err.number <> 0 then
%>
		<br>
		<b>ERROR: The CryptoWEB Control not be installed, or you have<br>
		a SHAREWARE version with time limit period terminated!</b><br>
		<br>
		<b>SOLUTION</b><br>
		<br>
		Try a installation of CryptoWEB Control.<br>
		<small>you can download SHAREWARE version form here <a href="http://www.prioregroup.com">http://www.prioregroup.com</a>.</small><br>
		<br>
		On MS-IIS WEB it is necessary to set the process IIS in low modality, as visible in figure
		<br>
		<br>		
		<img src="iissetup.gif" WIDTH="486" HEIGHT="453">
<%
	else
		with oCrypt
			' set the ecnryption mode
			.HashingType = CRYPT_MD5    ' (not required)
			.EncryptionType = CRYPT_RC4 ' (not required)
			select case request("mode")
				case "1"	' encryption
					a = request("pwd1")
					b = request("txt1")
					.Password = a
					out1 = .Encrypt(b, true) ' true for HEX output
				case "2"	' decryption
					.Password = request("pwd2")
					out2 = .Decrypt(request("txt2"), true) ' true for HEX output
			end select
		end with
		' free resources
		set oCrypt = nothing
%>
		<b><u>Encrypt Data</u></b><br>
		<br>
		<form action="default.asp" method="get">
		Data to encrypt : <input type="text" name="txt1" value="<%=request("txt1")%>">
		<input type="submit" value="Run"><br>
		Password : <input type="text" name="pwd1" value="<%=request("pwd1")%>"><br>
		Output:<br>
		<textarea cols="30" rows="5" readonly><%=out1%></textarea>
		<input type="hidden" name="mode" value="1">
		</form>
		<br>
		<hr>
		<br>
		<b><u>Decrypt Data</u></b><br>
		<br>
		<form action="default.asp" method="get">
		Data to decrypt : <input type="text" name="txt2" value="<%=request("txt2")%>">
		<input type="submit" value="Run"><br>
		Password : <input type="text" name="pwd2" value="<%=request("pwd2")%>"><br>
		Output:<br>
		<textarea cols="30" rows="5" readonly><%=out2%></textarea>
		<input type="hidden" name="mode" value="2">
		</form>
<%
	end if
%>

</body>