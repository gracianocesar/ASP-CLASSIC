<% 

' Cria uma nova inst�ncia da Classe formando o Objeto
 Set conDB = Server.CreateObject("ADODB.Connection")
  stringConexao =    "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password=123456;Initial Catalog=ASP; Data Source=GRACIANOTI"

' Abre a conex�o com o Banco de dados
conDB.ConnectionString = stringConexao
conDB.Open

    
%>

