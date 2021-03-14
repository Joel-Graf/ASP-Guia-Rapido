<%
  Conectar ao banco de dados PGSQL
  Set db_conn = Server.CreateObject("ADODB.Connection")
  str_conn = "Driver={PostgreSQL};Server=127.0.0.1;Port=5432;Database=asp;Uid=postgres;Pwd=postgres;"
  db_conn.connectionstring = str_conn
  db_conn.Open
%>