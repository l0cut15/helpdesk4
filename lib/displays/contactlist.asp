  <tr>
    <td nowrap align="left"><%
    
    Response.Write RScolData("FirstName",RecordSet) 
    Response.Write " "
    Response.Write RScolData("LastName",RecordSet) 
    %></td>
    <td align="left">
    <%  
    Response.Write RScolData("Email",RecordSet) 
    %></td>
    <td align="left"><% 
    If RScolData("DirectCode",RecordSet) <> "" then
     Response.Write "("
     Response.Write RScolData("DirectCode",RecordSet) 
     Response.Write ")"
    end if
    Response.Write RScolData("Direct",RecordSet) 
    %></td>
    <td align="left"><%  
    Response.Write RScolData("Cell",RecordSet) 
    %></td>
  </tr>



