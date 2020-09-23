Attribute VB_Name = "Module1"
Public Const gcDBNaam = "klanten.mdb"
Public Enum recstatus
recnieuw = 1
RecWijzigen = 2
recopgeslagen = 3
End Enum
Public blnsearchlistopen As Boolean

Public Function gfhaalpad(pstrpad As String) As String
    If Right$(pstrpad, 1) = "\" Then
        gfhaalpad = pstrpad
    Else
        gfhaalpad = pstrpad & "\"
    End If
    
End Function


Public Sub gsClearText(frm As Form)
    Dim ctrl As Control
        For Each ctrl In frm.Controls
    If TypeOf ctrl Is TextBox Then
        ctrl.Text = vbNullString
    End If
    Next ctrl
    
End Sub
Public Sub gsvulgrid(pgrdsearch As MSFlexGrid, pstrdb As String, pstrsql As String)
Dim dbsql As DAO.Database, rssql As DAO.Recordset
Dim fld As DAO.Field
Dim strdata As String, intteller As Integer
Dim intbreedte As Integer, intveldnr As Integer
Set dbsql = OpenDatabase(gfhaalpad(App.Path) & pstrdb)
Set rssql = dbsql.OpenRecordset(pstrsql, dbOpenSnapshot)
intveldnr = 0
With rssql
    If Not .BOF And Not .EOF Then
        With pgrdsearch
        .FixedCols = 0
        .FixedRows = 0
        .rows = 0
        .cols = rssql.Fields.Count
        
        For Each fld In rssql.Fields
            strdata = strdata & fld.Name & vbTab
            Next fld
            .AddItem strdata
            Do Until rssql.EOF
            strdata = vbNullString
            For Each fld In rssql.Fields
            strdata = strdata & fld.Value & vbTab
            Next fld
            .AddItem strdata
            For intteller = 0 To (rssql.Fields.Count - 1)
            .ColWidth(intteller) = pgrdsearch.Parent.TextWidth(rssql.Fields(intteler).Value) * 1.6
            intbreedte = intbreedte + .ColWidth(intteller)
            Next intteller
            rssql.MoveNext
            Loop
            .FixedRows = 1
            For intteller = 0 To .cols - 1
            intbreedte = intbreedte + 120
            End With
            Else
            End If
            End With
            rssql.Close
            dbsql.Close
            Set rssql = Nothing
            Set dbsql = Nothing
            
            
            
End Sub




Public Sub Sleesrec(rslees As DAO.Recordset)
    Dim dbklant As DAO.Database
    Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)

    txtrijksregisternr.Text = rslees.Fields("lngrijksregisternr").Value
    txtnaam.Text = rslees.Fields("strnaam").Value
    txtvoornaam.Text = rslees.Fields("strvoornaam").Value
    txtstraat.Text = rslees.Fields("strstraat").Value
    txthuisnr.Text = rslees.Fields("strhuisnummer").Value
    txtpostcode.Text = rslees.Fields("strpostcode").Value
    txtgemeente.Text = rslees.Fields("strgemeente").Value
    txtrizivnr.Text = rslees.Fields("lngrizivr").Value
    txtkg1.Text = rslees.Fields("strkg1").Value
    txtkg2.Text = rslees.Fields("strkg2").Value
    txtfederatie.Text = rslees.Fields("strfederatie").Value
    txtkaartnr.Text = rslees.Fields("lngkaartnummer").Value
    txtarts.Text = rslees.Fields("strarts").Value


End Sub



