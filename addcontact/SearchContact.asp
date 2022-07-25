<%Response.Buffer%>
<!-- #include file ="../lib/include.inc" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../lib/styles/hd.css" rel="stylesheet" type="text/css">
</head>



<body>
<p>Contacts Found</p>

<%
'| TableArrayContents(QueryResults) 'Draw table from array
'| Displays(DisplayName,RecordSet) ' Display RecordSet Data in Displa type set by displayname
'| DisplayNavBar(PageNumber,MaxPages) ' Display NavBar (HTML and code) -- Customizable
'|
'| DeformatIntoArray(FFArray,Delimiter,FFArrayDestRow) ' Remore Delimiters From Row 0
'| ProcessFormFields(FormFields,FFArray) Create Fields Array
'| Delimit(FFString,Delimiter) ' Delimit Strings
'|
'| CreateRS(RecordSet,ReqString,DSN) 'Create recordset from query string in Null then = "Nothing"
'| CloseRS(RecordSet) ' Delete RS object
'| CursorLocaton(RecordSet,IndexName) ' Find Current location of RAS cursor
'| SetCursor(RecordSet,RecordPos) ' Move Cursor to Positon
'|
'| QueryString(ReqString) (Function) ' Insert form values into query string
'| Function RmvCharStr(Str,Chars) ' Remove selected Chars from string
'| GenSqlInsert(TableName,FormFields) ' Generate insert String
'| GenSqlSelect(TableName,FormFields,TranID) ' Generated select String
'| GenTranIDInput(TableName,DSN) ' Generated Transaction ID in HTML
'| InsertFormData(IDName,FormFields,TableName,DSN)
'| 
'| PageRanges(RecordSet,ByVal DisplayLines) ' Returns String containg renges required to construct pages
'|
'| RSDisplay(RecordSet,ContentName,HeaderName,TailName,DisplayLines) ' Display Contents of RS in HTML
'| DisplayNavBar(PageNumber,MaxPages,DisplayName)
'| NavDisplays(DisplayName,DisplaySeg,PageNumber,MaxPages)
'|
'| RSColData(ColName,RecordSet) 'Return colum data from recordset

%>


<%Response.Write CreateRS(RecordSet,"Select * from Contacts;",DSN)  ' Run Query Against DB%>
<table>

<%Call RSDisplay(RecordSet,"ContactList","NavHead","NavTail",20)%>
</table>
<%Call CloseRS(RecordSet)%>
  



  


</body>

</html>
