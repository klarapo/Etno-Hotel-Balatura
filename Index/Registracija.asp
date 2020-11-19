<html>
<head><meta http-equiv="refresh" content="5;url=index.html"></head>
<body>

<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\Etno Hotel Balatura\Index\Gosti.mdb"

sql="INSERT INTO Rezervacije([Ime],[Prezime],"
sql=sql & "[Adresa],[Postanski broj],[Email],[Od datuma],[Do datuma],[Broj osoba],[Broj soba],[Ime sobe],[Komentar])"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("name") & "',"
sql=sql & "'" & Request.Form("surname") & "',"
sql=sql & "'" & Request.Form("add") & "',"
sql=sql & "'" & Request.Form("ptc") & "',"
sql=sql & "'" & Request.Form("email") & "',"
sql=sql & "'" & Request.Form("date1") & "',"
sql=sql & "'" & Request.Form("date2") & "',"
sql=sql & "'" & Request.Form("number1") & "',"
sql=sql & "'" & Request.Form("number2") & "',"
sql=sql & "'" & Request.Form("room_name") & "',"
sql=sql & "'" & Request.Form("comments") & "')"


conn.Execute sql,recaffected
if err<>0 then
  Response.Write("No update permissions!")
else
  Response.Write("<h3>" & recaffected & " record added</h3>")
  end if
conn.close
%>

</body>
</html>