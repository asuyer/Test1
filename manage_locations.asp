<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <script src="includes/jquery-ui-1.12.1/external/jquery/jquery.js" type="text/javascript"></script>
<head>
    <title>iCentrix Corp. Electronic Medical Records | Manage Locations</title>
    
    <%


         search_is_active = 1

      
           




        if Request("page") = "search_locations" or Request.QueryString("sf") <> "" THEN
            didSearch = 1
            foundMatch = 0

         if Request("search_is_active") <> "on"  then
              search_is_active = 0
             else
              search_is_active = 1
             end if

            if Request.QueryString("is_active") = "1" then
               search_is_active = 1
            end if
       

            
            if Request.QueryString("sf") <> "" THEN
                search_for = Request.QueryString("sf")
            else
                search_for = Request("find_name")
            end if
           
            'FIND LIST OF LOCATIONS THAT MATCH THIS NAME
            Set SQLStmtF = Server.CreateObject("ADODB.Command")
  	        Set rsF = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtF.CommandText = "exec find_locations_by_name '" & search_for & "',"  &  search_is_active
  	        SQLStmtF.CommandType = 1
  	        Set SQLStmtF.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmtF.CommandText
  	        rsF.Open SQLStmtF
            
            Do Until rsF.EOF

             '  if rsF("Is_Active") = "1" then
               ' search_is_active = 1
             '  else
                ' search_is_active = 0
              ' end if

	            foundMatch = 1
	        rsF.MoveNext
            Loop
            
        elseif Request("page") = "create_location" THEN
            
            'CHECK IF THIS FIRST,LAST,DOB EXISTS IF SO REJECT OR NOT BASED ON COMPANY CHOICE
            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt2.CommandText = "exec check_location_dup '" & Request("location_name") & "'"  
  	        SQLStmt2.CommandType = 1
  	        Set SQLStmt2.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rs2.Open SQLStmt2
  	        
  	        if rs2("location_exists") = "Yes" THEN
  	        
  	            'CHOICE DEPENDS ON COMPANY
  	            Response.write "Location exists in the system already."
  	        
  	        else
  	        
  	                      
  	            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	            Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmt2.CommandText = "exec create_location '" & Request("location_name") & "','" & Request("address") & "','" & Request("city") & "','" & Request("state") & "','" & Request("zip") & "','" & Request("phone") & "','" & Request("phone_ext") & "'"
  	            SQLStmt2.CommandType = 1
  	            Set SQLStmt2.ActiveConnection = conn
  	            'response.write "SQL = " & SQLStmt2.CommandText
  	            rs2.Open SQLStmt2
  	            
  	            close_window = 1
  	            
  	        end if
            
        end if
    %>
    
    <script type="text/javascript">
        if('<%=close_window%>' == '1')
        {

           if($("input[name=search_is_active]:checked").val() == '1') {
             window.location.href =  "manage_locations.asp?sf=<%=Request("location_name")%>&is_active=1";
          } else {
             window.location.href =  "manage_locations.asp?sf=<%=Request("location_name")%>>&is_active=0";

          }

        }
        
        function submitForSearch()
        {
            document.locationForm.page.value="search_locations";
            document.locationForm.submit();
        }
        
        function submitForSave()
        {
            document.locationForm.page.value="create_location";
            document.locationForm.submit();
        }
        
      
        function popupwindow(thehref,height,width,windowname)
        {
		    var windowFeatures ="menubar=no,scrollbars=yes,location=no,favorites=no,resizable=no,status=yes,toolbar=no,directories=no";
		    var test = "'"; 
		    winLeft = (screen.width-width)/2; 
		    winTop = (screen.height-(height+110))/2; 
		    myWin= window.open(thehref,windowname,"width=" + width +",height=" + height + ",left=" + winLeft + ",top=" + winTop + test + windowFeatures + test);
	
		    return false; 
	    }
	    
    </script>
<link href="includes/styles.css" rel="stylesheet" type="text/css" />
      <script type="text/javascript" src="includes/jquery-latest.pack.js"></script>
</head>
<body>
     <script type="text/javascript">
    $(document).ready(function(){
	    $("#show_all_locations").click(function () {
//alert($("input[name=search_is_active]:checked").val());
 event.preventDefault();
if($("input[name=search_is_active]:checked").val() == 'on') {
window.location.href =  "manage_locations.asp?sf=All&is_active=1";
} else {
window.location.href =  "manage_locations.asp?sf=All&is_active=0";
}
});
});
</script>
<form action="manage_locations.asp" name="locationForm" method="post">
<input type="hidden" name="page" value="search_locations" />
<!--#include file="includes/header_client.asp" -->
    <table cellpadding="4" cellspacing="0" border="0" width="968" align="center" id="box">
	  <tr><td colspan="6"><h1>Manage Locations:</h1></td></tr>
	  <tr>
        <td class="pod_title" colspan="6"><div style="float:left;">Location Name:&nbsp; <input type="text" name="find_name" value="<%=search_for%>" />&nbsp; <input type="button" onclick="submitForSearch();" value="Search" class="submit" />&nbsp;&nbsp; - or - &nbsp;&nbsp; <a href="" id="show_all_locations"><img src="images/Search.gif" width="16" height="16" border="0" alt="Show all Locations" title="Show all Locations" /> Show All Locations</a></div><div style="float:left;padding-left:30px;padding-top:5px;">Is Active:<input type="checkbox" name="search_is_active" <% if search_is_active = 1 then response.write "checked" %> /></div></td>
	  </tr>
	    <% if foundMatch = 1 THEN %>
	    <tr>
	        <td colspan="6"><hr width="100%" /></td>
	    </tr>
	    <tr>
	        <td colspan="6">
	            <table width="100%" border="0">
	                <tr>
	                    <td colspan="8"><b>Choose From Matches Found:</b></td>
	                </tr>

	                <tr>
	                    <td nowrap bgcolor="#d0d5e9"><b>ID</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Name</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Address</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>City</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>State</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Zip</b></td>
	                    <td nowrap bgcolor="#d0d5e9" align="center"><b>Actions</b></td>
	                </tr>
	                <%
	                rsF.MoveFirst
	                Do Until rsF.EOF


                        if cur_color = "" or cur_color="#99CCFF" then
  	                        cur_color = "#CCCCCC"
  	                    else
  	                        cur_color = "#99CCFF"
  	                    end if
	                %>
	                <tr style="background-color:<%=cur_color%>;">
	                    <td class="blackFontSmall" nowrap><%=rsF("location_id")%></td>
	                    <td class="blueFontSmall"><b><%=rsF("location_name")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("address")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("city")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("state")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("zip")%></b></td>
	                    <td class="blueFontSmall" colspan="2" nowrap align="center">
	                    <input type="button" onclick="window.location.href='edit_location.asp?lid=<%=rsF("location_id")%>';" value="Edit" class="submit" />
	                    </td>
	                </tr>
	                <%
	                rsF.MoveNext
                    Loop
	                %>
	            </table>
	        </td>
	    </tr>
	    <% end if %>
	    <% if didSearch = 1 THEN %>
	    <tr>
	        <td colspan="7"><hr width="100%" /></td>
	    </tr>
	    <tr>
	        <td colspan="7" class="pod_title">Or Enter New Location:</td>
	    </tr>
	    
	    <tr>
   
	        <td align="right" class="blackFontLarge" nowrap>Location Name:</td>
	        <td align="left">
	            <input type="text" name="location_name" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap>Address:</td>
	        <td align="left">
	            <input type="text" name="address" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap>City:</td>
	        <td align="left">
	            <input type="text" name="city" size="50" />
	        </td>
 	        
	        <td></td>
	    </tr>
	    <tr>	    
	        <td align="right" class="blackFontLarge" nowrap>State:</td>
	        <td align="left">
	            <input type="text" name="state" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap>Zip:</td>
	        <td align="left">
	            <input type="text" name="zip" size="50" />
	        </td>
	    
	        <td align="right" class="blackFontLarge" nowrap>Phone:</td>
             
            <td align="left">
	            <input type="text" name="phone" size="50" />
	        </td>
	    </tr>
         <tr>	    
	        <td align="right" class="blackFontLarge" nowrap>Phone Ext:</td>
	        <td align="left">
	            <input type="text" name="phone_ext" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap></td>
	        <td align="left">
	          
	        </td>
	    
	        <td align="right" class="blackFontLarge" nowrap></td>
             
            <td align="left">
	           
	        </td>
	    </tr>
	    <tr>
	        <td colspan="7" align="center"><input type="button" onclick="submitForSave();" value="Submit" class="submit" /></td>
	    </tr>   
	    <% end if %>
    </table>
</form>
</body>
</html>
