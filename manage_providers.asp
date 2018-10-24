<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>iCentrix Corp. Electronic Medical Records | Manage Providers</title>
    
    <%
        if Request("page") = "search_providers" or Request.QueryString("sf") <> "" THEN
            didSearch = 1
            foundMatch = 0
            
            if Request.QueryString("sf") <> "" THEN
                search_for = Request.QueryString("sf")
            else
                search_for = Request("find_name")
            end if
           
            'FIND LIST OF PROVIDERS THAT MATCH THIS NAME
            Set SQLStmtF = Server.CreateObject("ADODB.Command")
  	        Set rsF = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtF.CommandText = "exec find_providers_by_name '" & search_for & "'"  
  	        SQLStmtF.CommandType = 1
  	        Set SQLStmtF.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmtF.CommandText
  	        rsF.Open SQLStmtF
            
            Do Until rsF.EOF
	            foundMatch = 1
	        rsF.MoveNext
            Loop
            
        elseif Request("page") = "create_provider" THEN
            
            'CHECK IF THIS FIRST,LAST,DOB EXISTS IF SO REJECT OR NOT BASED ON COMPANY CHOICE
            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt2.CommandText = "exec check_provider_dup '" & Request("last_name") & "','" & Request("first_name") & "'"  
  	        SQLStmt2.CommandType = 1
  	        Set SQLStmt2.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rs2.Open SQLStmt2
  	        
  	        if rs2("provider_exists") = "Yes" THEN
  	        
  	            'CHOICE DEPENDS ON COMPANY
  	            Response.write "Provider exists in the system already."
  	        
  	        else
  	            
  	            'if Request("is_active") <> "" THEN
  	                cur_active = 1
  	            'else
  	            '    cur_active = 0
  	            'end if
  	                      
  	            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	            Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmt2.CommandText = "exec create_provider '" & Request("last_name") & "','" & Request("first_name") & "','" & Request("specialty") & "'," & cur_active & ",'" & Request("address") & "','" & Request("city") & "','" & Request("state") & "','" & Request("zip") & "','" & Request("phone") & "','" & Request("phone_ext") & "','" & Request("phone_type") & "','" & Request("phone2") & "','" & Request("phone_ext2") & "','"  & Request("phone_type2") & "','" & Request("fax") & "','" & Request("email") & "','" & Request("concat_name") & "','" & Request("npi_number") & "','" & Request("provider_care_number") & "','" & Request("web_site") & "'"
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
            window.location.href =  "manage_providers.asp?sf=<%=Request("last_name")%>";
        }
        
        function submitForSearch()
        {
            document.providerForm.page.value="search_providers";
            document.providerForm.submit();
        }
        
        function submitForSave()
        {
            if(document.providerForm.concat_name.value == "")
            {
                alert("A Description for the Provider is required.");
                return false;
            }
            else
            {
                document.providerForm.page.value="create_provider";
                document.providerForm.submit();
            }
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
</head>
<body>
<form action="manage_providers.asp" name="providerForm" method="post">
<input type="hidden" name="page" value="search_providers" />
<!--#include file="includes/header_client.asp" -->
    <table cellpadding="4" cellspacing="0" border="0" width="968" align="center" id="box">
	  <tr><td colspan="6"><h1>Manage Providers:</h1></td></tr>
	  <tr>
        <td class="pod_title" colspan="6">Provider Last Name:&nbsp; <input type="text" name="find_name" value="<%=search_for%>" />&nbsp; <input type="button" onclick="submitForSearch();" value="Search" class="submit" />&nbsp;&nbsp; - or - &nbsp;&nbsp; <a href="manage_providers.asp?sf=All"><img src="images/Search.gif" width="16" height="16" border="0" alt="Show all Providers" title="Show all Providers" /> Show All Providers</a></td>
	  </tr>
	    <% if foundMatch = 1 THEN %>
	    <tr>
	        <td colspan="6"><hr width="100%" /></td>
	    </tr>
	    <tr>
	        <td colspan="6">
	            <table width="100%" border="0">
	                <tr>
	                    <td colspan="8"><b>Choose From Matches Found</b></td>
	                </tr>

	                <tr>
	                    <td nowrap bgcolor="#d0d5e9"><b>Name</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Specialty</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Address</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>City</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>State</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Zip</b></td>
	                    <td nowrap bgcolor="#d0d5e9"><b>Action</b></td>
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
	                    <td class="blueFontSmall" nowrap><b><%=rsF("Provider_name")%></b></td>
	                    <td class="blueFontSmall"><b><%=rsF("specialty")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("address")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("city")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("state")%></b></td>
	                    <td class="blueFontSmall" ><b><%=rsF("zip")%></b></td>
	                    <td class="blueFontSmall" colspan="2" nowrap>
	                    <input type="button" onclick="window.location.href='edit_provider.asp?prid=<%=rsF("provider_id")%>';" value="Edit" class="submit" />
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
	        <td colspan="7" class="pod_title">Or Enter New Provider:</td>
	    </tr>
	    
	    <tr>
	        <td align="right" class="blackFontLarge" nowrap>Provider Description:</td>
	        <td align="left">
	            <input type="text" name="concat_name" size="50" />
	        </td>
   
	        <td align="right" class="blackFontLarge" nowrap>First Name:</td>
	        <td align="left">
	            <input type="text" name="first_name" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap>Last Name:</td>
	        <td align="left">
	            <input type="text" name="last_name" size="50" />
	        </td>	         
 	        
	        <td></td>
	    </tr>
	    <tr>	    
	        <td align="right" class="blackFontLarge" nowrap>Address:</td>
	        <td align="left">
	            <input type="text" name="address" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap>City:</td>
	        <td align="left">
	            <input type="text" name="city" size="50" />
	        </td>
	        
	        <td align="right" class="blackFontLarge" nowrap>State:</td>
	        <td align="left">
	           <select name="state" >
  <option value="AK">AK</option>
  <option value="AL">AL</option>
  <option value="AR">AR</option>
  <option value="AZ">AZ</option>
  <option value="CA">CA</option>
  <option value="CO">CO</option>
  <option value="CT">CT</option>
  <option value="DC">DC</option>
  <option value="DE">DE</option>
  <option value="FL">FL</option>
  <option value="GA">GA</option>
  <option value="HI">HI</option>
  <option value="IA">IA</option>
  <option value="ID">ID</option>
  <option value="IL">IL</option>
  <option value="IN">IN</option>
  <option value="KS">KS</option>
  <option value="KY">KY</option>
  <option value="LA">LA</option>
  <option value="MA" selected>MA</option>
  <option value="MD">MD</option>
  <option value="ME">ME</option>
  <option value="MI">MI</option>
  <option value="MN">MN</option>
  <option value="MO">MO</option>
  <option value="MS">MS</option>
  <option value="MT">MT</option>
  <option value="NC">NC</option>
  <option value="ND">ND</option>
  <option value="NE">NE</option>
  <option value="NH">NH</option>
  <option value="NJ">NJ</option>
  <option value="NM">NM</option>
  <option value="NV">NV</option>
  <option value="NY">NY</option>
  <option value="OH">OH</option>
  <option value="OK">OK</option>
  <option value="OR">OR</option>
  <option value="PA">PA</option>
  <option value="RI">RI</option>
  <option value="SC">SC</option>
  <option value="SD">SD</option>
  <option value="TN">TN</option>
  <option value="TX">TX</option>
  <option value="UT">UT</option>
  <option value="VA">VA</option>
  <option value="VT">VT</option>
  <option value="WA">WA</option>
  <option value="WI">WI</option>
  <option value="WV">WV</option>
  <option value="WY">WY</option>
</select>

	        </td>        
	    
	        <td></td>
	    </tr>
	    <tr>
	        <td align="right" class="blackFontLarge" nowrap>Zip:</td>
	        <td align="left">
	            <input type="text" name="zip" size="50" />
	        </td>
	        <td align="right" class="blackFontLarge">Email:</td>
	        <td align="left">
	            <input type="text" name="email" size="50"/>
	        </td>
	        <td align="right" class="blackFontLarge" nowrap>Specialty:</td>
	        <td align="left">
	            <select name="specialty" id="specialty">
	                <%
	                Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                SQLStmt2.CommandText = "select code_name, short_desc from code_map where code_type = (select code_type from code_def where form_sid = 'Specialist_Type') order by short_desc"  
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                'response.write "SQL = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
  	                Do Until rs2.EOF
	                %>
	                    <option value="<%=rs2("code_name")%>"><%=UCASE(rs2("short_desc"))%></option>
	                <%
	                rs2.MoveNext
                    Loop
	                %>
	            </select>
	        </td>	       
	       
	        
	        <td colspan="1"></td>
	    </tr>
	    <tr>
             <td align="right" class="blackFontLarge">NPI Number:</td>
	        <td align="left">
	            <input type="text" name="npi_number" />
	        </td>
	         
	         <td align="right" class="blackFontLarge">Provider Care #:</td>
	        <td align="left">
	            <input type="text" name="provider_care_number" />
	        </td>
	       
	        
	         <td align="right" class="blackFontLarge">Fax:</td>
	        <td align="left">
	            <input type="text" name="fax" />
	        </td>
	    </tr>
          <tr>
             <td align="right" class="blackFontLarge">Phone:</td>
	        <td align="left">
	            <input type="text" name="phone" />
	        </td>
	          <td align="right" class="blackFontLarge">Ext:</td>
	        <td align="left">
	            <input type="text" name="phone_ext" />
	        </td>
                 <td align="right" class="blackFontLarge">Phone Type:</td>
	        <td align="left">
	            <select name="phone_type">
                      <option value="Office">Office</option>
                      <option value="Cell" >Cell</option>
                       <option value="Home" >Home</option>
                      </select>
	        </td>
	        
	    </tr>
         <tr>
             <td align="right" class="blackFontLarge">Phone:</td>
	        <td align="left">
	            <input type="text" name="phone2" />
	        </td>
	          <td align="right" class="blackFontLarge">Ext:</td>
	        <td align="left">
	            <input type="text" name="phone_ext2" />
	        </td>
                 <td align="right" class="blackFontLarge">Phone Type:</td>
	        <td align="left">
	              <select name="phone_type2">
                      <option value="Office">Office</option>
                      <option value="Cell" >Cell</option>
                       <option value="Home" >Home</option>
                      </select>
	        </td>
	        
	    </tr>
          <tr>
            <td align="right" class="blackFontLarge">Web Address:</td>
	        <td align="left">
	            <input type="text" name="web_site" size="50"/>
	        </td>
	         
	       <td colspan="4"></td>
	        
	    </tr>
	    <tr>
	        <td colspan="7" align="center"><input type="button" onclick="submitForSave();" value="Submit" class="submit" /></td>
	    </tr> 
          
	    <% end if %>
    </table>
</form>
</body>
</html>
