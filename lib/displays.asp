
<%


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Sub Displays(DisplayName,RecordSet)

Select Case DisplayName

'Included Displays



'-----------------------------------------------------------------------------
	Case "Header"
	%>
<table><TR><TD>
<%
	
'-----------------------------------------------------------------------------
	
	Case "Tail"
	
	Response.Write "</TABLE>"
	
 
    
'-----------------------------------------------------------------------------
 

Case "GuestBook" %>
 
 <!-- #include file ="./displays/guestbook.asp" -->

 
  <%
  

   
 '-----------------------------------------------------------------------------
    
 case "DZOption"
 %>
<option value="<%Response.Write RSColData("Dropzone",RecordSet)%>"><%Response.Write RSColData("Dropzone",RecordSet)%></option>
 <%
 
 
 
 
  '-----------------------------------------------------------------------------
    
 case "Gallery" %>
 
 <!-- #include file ="./displays/gallery.asp" -->

<%
 
 '-----------------------------------------------------------------------------
    
 case "GalleryOption"
 %>
<option value="<%Response.Write RSColData("GalleryID",RecordSet)%>"><%Response.Write RSColData("GalleryName",RecordSet)%></option>
 <%
 
 '-----------------------------------------------------------------------------
 
Case "ContactList" %>
 
 <!-- #include file ="./displays/contactlist.asp" -->

 
  <%
  

   
 '-----------------------------------------------------------------------------
 
 case else
    

'-----------------------------------------------------------------------------
	
End Select

end sub




%>






























