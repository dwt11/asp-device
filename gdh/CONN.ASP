<%


db="jlhs.mdb"
Set connjlhs = Server.CreateObject("ADODB.Connection")
Connas = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
    connjlhs.Open connas



%>