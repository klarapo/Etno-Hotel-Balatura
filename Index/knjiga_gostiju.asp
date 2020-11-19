<%@ Language="VBScript" %>
<!DOCTYPE html>

<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Ethno Hotel Balatura</title>
        <link href="stil.css" rel="stylesheet" type="text/css" />
<script src="js/jquery-1.10.2.min.js"></script>
<script src="js/lightbox-2.6.min.js"></script>
<link href="css/lightbox.css" rel="stylesheet" />
    </head>
    <body>
        
<div id="container">

<div id="header">
</div>

<div id="navigacija">
<table width="1000">
  <tr>
    <th scope="col"><a href="index.html">Početna</a></th>
    <th scope="col"><a href="o_nama.html">O nama</a></th>
    <th scope="col"><a href="galerija.html">Galerija</a></th>
    <th scope="col"><a href="rezervacije.html">Rezervacija</a></th>
    <th scope="col"><a href="knjiga_gostiju.asp">Knjiga gostiju</a></th>
    <th scope="col"><a href="kontakt.html">Kontakt</a></th>
  </tr>
</table>
</div>
    <div id="sadrzaj">
<br /><br />

    <%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\Etno Hotel Balatura\Index\Gosti.mdb"

set rs = Server.CreateObject("ADODB.recordset")
sql="SELECT Ime,Prezime,Email,[Ime sobe],Komentar FROM Rezervacije"
rs.Open sql, conn
%>

<table border="1" width="98%">
  <tr>
  <%for each x in rs.Fields
    response.write("<th>" & x.name & "</th>")
  next%>
  </tr>
  <%do until rs.EOF%>
    <tr>
    <%for each x in rs.Fields%>
      <td><%Response.Write(x.value)%></td>
    <%next
    rs.MoveNext%>
    </tr>
  <%loop
  rs.close
  conn.close
  %>
</table>
     
     </div> 
       <div id="footer">
     Klara Popović &copy; 2014 
      </div>  
     </div>
     
    </body>
</html>