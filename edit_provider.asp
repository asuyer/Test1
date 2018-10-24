<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>iCentrix Corp. Electronic Medical Records</title>
    <!--#include file="upload.asp" --> 
    <%
    cur_month = Month(Date())
    cur_year = Year(Date())
    cur_day = Day(Date())
    
    'Resize image using ASP/VBS   
'2007 www.motobit.com    
Function ResizeImage(FileName, OutFileName, OutFormat, Width, Height)   
  Dim Chs, chConstants   
  'Create an OWC chart object   
  Set Chs = CreateObject("OWC11.ChartSpace")   
     
  Set chConstants = Chs.Constants   
     
  'Set background of the chart   
  Chs.Interior.SetTextured FileName, chConstants.chStretchPlot, , chConstants.chAllFaces   
  Chs.border.color = -3   
  
  'Do something with border   
  'Chs.border.color = &H0000FF   
  'Chs.border.Weight = 3   
  
  'export the picture to a file   
  Chs.ExportPicture OutFileName, OutFormat, Width, Height   
     
  'or return it as a binary data for BinaryWrite   
  'ResizeImage = Chs.GetPicture(OutFormat, Width, Height)   
  ResizeImage = 1
End Function

    dim MyUploader
    Set MyUploader = New FileUploader 
    MyUploader.Upload()

           thecount= MyUploader.Count() -1   
               
           'FIND LOCATION THAT MATCH THIS ID
            Set SQLStmtF = Server.CreateObject("ADODB.Command")
  	        Set rsF = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtF.CommandText = "exec get_provider_info " & Request.QueryString("prid")
  	        SQLStmtF.CommandType = 1
  	        Set SQLStmtF.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rsF.Open SQLStmtF
  	        
  	        cur_id = rsF("Provider_ID")
  	        cur_last = rsF("last_name")
  	        cur_first = rsF("first_name")
  	        cur_name = rsF("Provider_Name")
  	        cur_specialty = rsF("Specialty")
  	        cur_active = rsF("is_active")
  	        cur_address = rsF("Address")
  	        cur_city = rsF("City")
  	        cur_state = rsF("State")
  	        cur_zip = rsF("Zip")  	
  	        cur_phone = rsF("phone")
            cur_phone_ext = rsF("phone_ext")
            cur_phone_type = rsF("phone_type")
            cur_phone2 = rsF("phone2")
            cur_phone_ext2 = rsF("phone_ext2")
            cur_phone_type2 = rsF("phone_type2")
  	        cur_fax = rsF("fax")
  	        cur_email = rsF("email")
            cur_concat_name = rsF("concat_name")  
  	        cur_npi_number = rsF("np_number")  
            cur_provider_care_number = rsF("provider_care_number")  
            cur_web_site = rsF("web_site")
        
        
       ' response.write cur_phone_type  
  	                    
        if MyUploader.Item("page") = "edit_provider" THEN
               
            'CHECK IF THIS FIRST,LAST,DOB EXISTS IF SO REJECT OR NOT BASED ON COMPANY CHOICE
            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt2.CommandText = "exec check_provider_dup_for_update " & Request.QueryString("prid") & ",'" & MyUploader.Item("last_name") & "','" &  MyUploader.Item("first_name") & "'"
  	        SQLStmt2.CommandType = 1
  	        Set SQLStmt2.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rs2.Open SQLStmt2
  	        
  	        if rs2("provider_exists") = "Yes" THEN
  	        
  	            'CHOICE DEPENDS ON COMPANY
  	            Response.write "provider exists in the system already sorry"
  	        
  	        else
  	        
  	            if MyUploader.Item("is_active") <> "" THEN
  	                cur_active = 1
  	            else
  	                cur_active = 0
  	            end if
 	            
  	            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	            Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmt2.CommandText = "exec edit_provider " & Request.QueryString("prid") & ",'" & Replace(MyUploader.Item("last_name"),"'","''") & "','" &  Replace(MyUploader.Item("first_name"),"'","''") & "','" & MyUploader.Item("specialty") & "'," & cur_active & ",'" & MyUploader.Item("address") & "','" & MyUploader.Item("city") & "','" & MyUploader.Item("state") & "','" & MyUploader.Item("zip") & "','" & MyUploader.Item("phone") & "','" & MyUploader.Item("phone_ext") & "','" & MyUploader.Item("phone_type") & "','" & MyUploader.Item("phone2") & "','" & MyUploader.Item("phone_ext2") & "','" & MyUploader.Item("phone_type2") & "','"  & MyUploader.Item("fax") & "','" & MyUploader.Item("email") & "','" & Replace(MyUploader.Item("concat_name"),"'","''") & "','" & MyUploader.Item("npi_number") & "','" & MyUploader.Item("provider_care_number") & "','" & MyUploader.Item("web_site") & "'"
  	            SQLStmt2.CommandType = 1
  	            Set SQLStmt2.ActiveConnection = conn
  	            'response.write "SQL = " & SQLStmt2.CommandText
  	            rs2.Open SQLStmt2
  	            
  	            ' do file last
                Dim File
                For each File in MyUploader.Files.Items
                    if (File.FileName <> "") Then
                        'SAVE THE FILE TO THE TEMP DIR USING THE USER NAME, AND TIMESTAMP
                        Orig_name = File.FileName
                        'response.write "size = " & File.Size
                        New_name = Session("user_name") & "_" & Year(Date()) & Month(Date()) & Day(Date()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now()) & "_" & Orig_name
                        'response.write "new name = " & New_name
                        File.FileName = New_name
                        File.SaveToDisk form_root_path & "web_root\temp_docs"
                    end if
                Next
 
                if (New_name <> "") THEN

                    old_temp_file = form_doc_path & New_name
                               
                    const bytesToKb = 1024
                    set objFSO = createobject("Scripting.FileSystemObject")
                    set objFile = objFSO.GetFile(old_temp_file)
                    'response.write "File Size: " & cint(objFile.Size / bytesToKb) & "Kb"
                    
                    if cint(objFile.Size / bytesToKb) > 3000 THEN
                        response.Write "<font color='red'><b>Please upload an image 3 mb or smaller in size for the client image.</b></font>"
                        image_error = 1
                    else
                        image_error = 0
                    
                        if MyUploader.Item("picture_date") <> "" THEN
                            picture_date = MyUploader.Item("picture_date")
                        else
                            picture_date = cur_year & "-" & cur_month & "-" & cur_day
                        end if
                    
                        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
                        Set rs2 = Server.CreateObject ("ADODB.Recordset")

                        SQLStmt2.CommandText = "update providers_master set [picture] = (SELECT * FROM OPENROWSET(BULK '" & form_doc_path & New_name & "', SINGLE_BLOB)AS x ), [orig_picture] = (SELECT * FROM OPENROWSET(BULK '" & form_doc_path & New_name & "', SINGLE_BLOB)AS x ) WHERE [provider_id]="& Request.QueryString("prid")
                        SQLStmt2.CommandType = 1
                        Set SQLStmt2.ActiveConnection = conn
                        'response.write "SQL = " & SQLStmt2.CommandText
                        rs2.Open SQLStmt2
                    end if                    

                    'DELETE TEMP FILE
                    Set fs=Server.CreateObject("Scripting.FileSystemObject")
                    if fs.FileExists(form_doc_path & New_name) then
                         fs.DeleteFile(form_doc_path & New_name)
                    end if
                    set fs=nothing
                      
  	            end if
  	              	            
  	            close_window = 1
  	        end if
            
        end if
    %>
    
    <script type="text/javascript">
        if('<%=close_window%>' == '1')
        {
            window.opener.location.reload();
            window.location.href = "manage_providers.asp?sf=" + '<%=cur_last%>';
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
                document.providerForm.page.value="edit_provider";
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
/*	function setValues()
	{

	    document.providerForm.specialty.value = '<%=cur_specialty%>';
	} */
    </script>
<link href="includes/styles.css" rel="stylesheet" type="text/css" />
</head>
<body >
<form action="edit_provider.asp?prid=<%=Request.QueryString("prid")%>" name="providerForm" method="post" enctype="multipart/form-data">
<input type="hidden" name="page" value="search_providers" />
<!--#include file="includes/header_client.asp" -->
    <table cellpadding="4" cellspacing="0" border="0" width="969" align="center" id="box"
	    <tr>
	        <td colspan="6"><h1>Edit Provider:</h1></td>
	    </tr>
	    <tr>
	        <td colspan="6" valign="middle"><img src="images/back_arrow.gif" width="17" height="16" border="0" alt="Back" title="Back" />
	        <% if Request.QueryString("fromModule") = "1" THEN %>
	        <a href="#" onclick="window.close();">Back to Provider Module</a>
	        <% else %>
	        <a href="manage_providers.asp?sf=<%=cur_last%>">Back to Providers Search</a>
	        <% end if %>
	        </td>
	    </tr>
	    <tr>
	        <td align="right" class="blackFontLarge">Provider Description:</td>
	        <td align="left">
	            <input type="text" name="concat_name" size="50" value="<%=cur_concat_name%>">
	        </td>  
	    
	        <td align="right" class="blackFontLarge">First Name:</td>
	        <td align="left">
	            <input type="text" name="first_name" value="<%=cur_first%>">
	        </td>
	        
	        <td align="right" class="blackFontLarge">Last Name:</td>
	        <td align="left">
	            <input type="text" name="last_name" value="<%=cur_last%>">
	        </td>	   
	        
	    </tr>
	    <tr>
	        <td align="right" class="blackFontLarge">Specialty:</td>
	        <td align="left">
	            <select name="specialty" id="specialty">
                    <option value="">Select a Specialty</option>
	                <%
	                Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                SQLStmt2.CommandText = "select code_name, short_desc from code_map where code_type = (select code_type from code_def where form_sid = 'Specialist_Type') order by short_desc"  
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                'response.write "SQL = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
  	                Do Until rs2.EOF


                      
                        code_name = rs2("code_name")
                          short_desc = rs2("short_desc")


                    
	                %>
	                    <option value="<%=code_name%>" <%if UCASE(code_name) = UCASE(cur_specialty) then response.write "selected" %>><%=UCASE(short_desc)%></option>
	                <%
	                rs2.MoveNext
                    Loop
	                %>
                 
	            </select>	            
	        </td>
	        
	        <td align="right" class="blackFontLarge">Is Active:</td>
	        <td align="left">
	            <input type="checkbox" name="is_active" <%if cur_active = "1" THEN%>checked <%end if %> />
	        </td>
	        
	        <td align="right" class="blackFontLarge">Address:</td>
	        <td align="left">
	            <input type="text" name="address" value="<%=cur_address%>">
	        </td>
	    </tr>
	    <tr>
	        <td align="right" class="blackFontLarge">City:</td>
	        <td align="left">
	            <input type="text" name="city" value="<%=cur_city%>">
	        </td>
	        <td align="right" class="blackFontLarge">State:</td>
	        <td align="left">
                <select name="state">
	                <option value="AL" <%if cur_state="AL" then response.write "selected" end if%>>AL</option>
	                <option value="AK" <%if cur_state="AK" then response.write "selected" end if%>>AK</option>
	                <option value="AR" <%if cur_state="AR" then response.write "selected" end if%>>AR</option>	
	                <option value="AZ" <%if cur_state="AZ" then response.write "selected" end if%>>AZ</option>
	                <option value="CA" <%if cur_state="CA" then response.write "selected" end if%>>CA</option>
	                <option value="CO" <%if cur_state="CO" then response.write "selected" end if%>>CO</option>
	                <option value="CT" <%if cur_state="CT" then response.write "selected" end if%>>CT</option>
	                <option value="DC" <%if cur_state="DC" then response.write "selected" end if%>>DC</option>
	                <option value="DE" <%if cur_state="DE" then response.write "selected" end if%>>DE</option>
	                <option value="FL" <%if cur_state="FL" then response.write "selected" end if%>>FL</option>
	                <option value="GA" <%if cur_state="GA" then response.write "selected" end if%>>GA</option>
	                <option value="HI" <%if cur_state="HI" then response.write "selected" end if%>>HI</option>
	                <option value="IA" <%if cur_state="IA" then response.write "selected" end if%>>IA</option>	
	                <option value="ID" <%if cur_state="ID" then response.write "selected" end if%>>ID</option>
	                <option value="IL" <%if cur_state="IL" then response.write "selected" end if%>>IL</option>
	                <option value="IN" <%if cur_state="IN" then response.write "selected" end if%>>IN</option>
	                <option value="KS" <%if cur_state="KS" then response.write "selected" end if%>>KS</option>
	                <option value="KY" <%if cur_state="KY" then response.write "selected" end if%>>KY</option>
	                <option value="LA" <%if cur_state="LA" then response.write "selected" end if%>>LA</option>
	                <option value="MA" <%if cur_state="MA" then response.write "selected" end if%>>MA</option>
	                <option value="MD" <%if cur_state="MD" then response.write "selected" end if%>>MD</option>
	                <option value="ME" <%if cur_state="ME" then response.write "selected" end if%>>ME</option>
	                <option value="MI" <%if cur_state="MI" then response.write "selected" end if%>>MI</option>
	                <option value="MN" <%if cur_state="MN" then response.write "selected" end if%>>MN</option>
	                <option value="MO" <%if cur_state="MO" then response.write "selected" end if%>>MO</option>	
	                <option value="MS" <%if cur_state="MS" then response.write "selected" end if%>>MS</option>
	                <option value="MT" <%if cur_state="MT" then response.write "selected" end if%>>MT</option>
	                <option value="NC" <%if cur_state="NC" then response.write "selected" end if%>>NC</option>	
	                <option value="NE" <%if cur_state="NE" then response.write "selected" end if%>>NE</option>
	                <option value="NH" <%if cur_state="NH" then response.write "selected" end if%>>NH</option>
	                <option value="NJ" <%if cur_state="NJ" then response.write "selected" end if%>>NJ</option>
	                <option value="NM" <%if cur_state="NM" then response.write "selected" end if%>>NM</option>			
	                <option value="NV" <%if cur_state="NV" then response.write "selected" end if%>>NV</option>
	                <option value="NY" <%if cur_state="NY" then response.write "selected" end if%>>NY</option>
	                <option value="ND" <%if cur_state="ND" then response.write "selected" end if%>>ND</option>
	                <option value="OH" <%if cur_state="OH" then response.write "selected" end if%>>OH</option>
	                <option value="OK" <%if cur_state="OK" then response.write "selected" end if%>>OK</option>
	                <option value="OR" <%if cur_state="OR" then response.write "selected" end if%>>OR</option>
	                <option value="PA" <%if cur_state="PA" then response.write "selected" end if%>>PA</option>
	                <option value="RI" <%if cur_state="RI" then response.write "selected" end if%>>RI</option>
	                <option value="SC" <%if cur_state="SC" then response.write "selected" end if%>>SC</option>
	                <option value="SD" <%if cur_state="SD" then response.write "selected" end if%>>SD</option>
	                <option value="TN" <%if cur_state="TN" then response.write "selected" end if%>>TN</option>
	                <option value="TX" <%if cur_state="TY" then response.write "selected" end if%>>TX</option>
	                <option value="UT" <%if cur_state="UT" then response.write "selected" end if%>>UT</option>
	                <option value="VT" <%if cur_state="VT" then response.write "selected" end if%>>VT</option>
	                <option value="VA" <%if cur_state="VA" then response.write "selected" end if%>>VA</option>
	                <option value="WA" <%if cur_state="WA" then response.write "selected" end if%>>WA</option>
	                <option value="WI" <%if cur_state="WI" then response.write "selected" end if%>>WI</option>	
	                <option value="WV" <%if cur_state="WV" then response.write "selected" end if%>>WV</option>
	                <option value="WY" <%if cur_state="WA" then response.write "selected" end if%>>WY</option>
              </select>
	          
	        </td>
	        <td align="right" class="blackFontLarge">Zip:</td>
	        <td align="left">
	            <input type="text" name="zip" value="<%=cur_zip%>">
	        </td>              
	    </tr>

        
         <tr>	        
	        <td align="right" class="blackFontLarge">Phone:</td>
	        <td align="left">
	            <input type="text" name="phone" value="<%=cur_phone%>">
	        </td>
	        <td align="right" class="blackFontLarge">Ext:</td>
	        <td align="left">
	            <input type="text" name="phone_ext" value="<%=cur_phone_ext%>">
	        </td>
	        
	        <td align="right" class="blackFontLarge">Phone Type:</td>
	        <td align="left">
	            <select name="phone_type">
                      <option value="Office" <%if cur_phone_type="Office" then response.write "selected" end if%>>Office</option>
                      <option value="Cell" <%if cur_phone_type="Cell" then response.write "selected" end if%>>Cell</option>
                       <option value="Home" <%if cur_phone_type="Home" then response.write "selected" end if%>>Home</option>
                      </select>
	        </td>
	    </tr>

         <tr>	        
	        <td align="right" class="blackFontLarge">Phone:</td>
	        <td align="left">
	            <input type="text" name="phone2" value="<%=cur_phone2%>">
	        </td>
	        <td align="right" class="blackFontLarge">Ext:</td>
	        <td align="left">
	            <input type="text" name="phone_ext2" value="<%=cur_phone_ext2%>">
	        </td>
	        
	        <td align="right" class="blackFontLarge">Phone Type:</td>
	        <td align="left">
	            <select name="phone_type2">
                      <option value="Office" <%if cur_phone_type2="Office" then response.write "selected" end if%>>Office</option>
                      <option value="Cell" <%if cur_phone_type2="Cell" then response.write "selected" end if%>>Cell</option>
                       <option value="Home" <%if cur_phone_type2="Home" then response.write "selected" end if%>>Home</option>
                      </select>
	        </td>
	    </tr>



	    <tr>	        
	        
	        <td align="right" class="blackFontLarge">Fax:</td>
	        <td align="left">
	            <input type="text" name="fax" value="<%=cur_fax%>">
	        </td>
	        
	        <td align="right" class="blackFontLarge">Web Address:</td>
	        <td align="left">
	            <input type="text" name="web_site" value="<%=cur_web_site%>" size="50">
	        </td>

            <td align="right" class="blackFontLarge">NPI Number:</td>
	        <td align="left">
	            <input type="text" name="npi_number" value="<%=cur_npi_number%>">
	        </td>

	    </tr>
	    <tr>     	        
	        
	        <td align="right" class="blackFontLarge" bgcolor="#e0e0aa">Upload picture file</td>
	        <td align="left" class="blackFontLarge"><input type="file" name="target_file" size="" /></td>
	           <td align="right" class="blackFontLarge">Provider Care #:</td>
	        <td align="left">
	            <input type="text" name="provider_care_number" value="<%=cur_provider_care_number%>">
	        </td>
	        <td align="right" class="blackFontLarge">Email:</td>
	        <td align="left">
	            <input type="text" name="email" value="<%=cur_email%>" size="50">
	        </td>
	    </tr>
	    <tr>
	        <td colspan="6" align="center"><input type="button" onclick="submitForSave();" value="Save Changes" class="submit" /></td>
	    </tr>   
    </table>
</form>
</body>
</html>
