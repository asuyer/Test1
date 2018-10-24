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

  	        SQLStmtF.CommandText = "exec get_location_info " & Request.QueryString("lid")
  	        SQLStmtF.CommandType = 1
  	        Set SQLStmtF.ActiveConnection = conn
  	     '   response.write "SQL = " & SQLStmtF.CommandText
  	        rsF.Open SQLStmtF
  	        
  	        cur_id = rsF("Location_ID")
  	        cur_name = rsF("Location_Name")
  	        cur_address = rsF("Address")
  	        cur_city = rsF("City")
  	        cur_state = rsF("State")
  	        cur_zip = rsF("Zip")  
           if rsF("Is_Active") then
            cur_is_active = 1
           else
             cur_is_active = 0	
           end if
                
        cur_phone = rsF("phone")  
        cur_phone_ext = rsF("phone_ext")  
      
  	                    
        if MyUploader.Item("page") = "edit_location" THEN
               
            'CHECK IF THIS FIRST,LAST,DOB EXISTS IF SO REJECT OR NOT BASED ON COMPANY CHOICE
            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt2.CommandText = "exec check_location_dup_for_update " & Request.QueryString("lid") & ",'" & MyUploader.Item("Location_name") & "'" 
  	        SQLStmt2.CommandType = 1
  	        Set SQLStmt2.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rs2.Open SQLStmt2
  	        
  	        if rs2("location_exists") = "Yes" THEN
  	        
  	            'CHOICE DEPENDS ON COMPANY
  	            Response.write "location exists in the system already sorry"
  	        
  	        else
    
               if MyUploader.Item("is_active") = "on" then
                cur_is_active = 1
                else
               cur_is_active = 0
                end if

 	            
  	            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	            Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmt2.CommandText = "exec edit_location " & Request.QueryString("lid") & ",'" & MyUploader.Item("location_name") & "','" & MyUploader.Item("address") & "','" & MyUploader.Item("city") & "','" & MyUploader.Item("state") & "','" & MyUploader.Item("zip") & "'," & cur_is_active & ",'" & MyUploader.Item("phone") & "','" &  MyUploader.Item("phone_ext") & "'"
  	            SQLStmt2.CommandType = 1
  	            Set SQLStmt2.ActiveConnection = conn
  	            response.write "SQL = " & SQLStmt2.CommandText
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

                        SQLStmt2.CommandText = "update location_master set [picture] = (SELECT * FROM OPENROWSET(BULK '" & form_doc_path & New_name & "', SINGLE_BLOB)AS x ), [orig_picture] = (SELECT * FROM OPENROWSET(BULK '" & form_doc_path & New_name & "', SINGLE_BLOB)AS x ) WHERE [location_id]="& Request.QueryString("lid")
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
            window.location.href = "manage_locations.asp?sf=" + '<%=cur_name%>' + "&is_active=<%=cur_is_active %>";
        }
        
        function submitForSave()
        {
            document.locationForm.page.value="edit_location";
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
</head>
<body>
<form action="edit_location.asp?lid=<%=Request.QueryString("lid")%>" name="locationForm" method="post" enctype="multipart/form-data">
<input type="hidden" name="page" value="search_locations" />
<!--#include file="includes/header_client.asp" -->
    <table cellpadding="4" cellspacing="0" border="0" width="969" align="center" id="box"
	    <tr>
	        <td colspan="6"><h1>Edit Location:</h1></td>
	    </tr>
	    <tr>
	        <td colspan="6" valign="middle"><img src="images/back_arrow.gif" width="17" height="16" border="0" alt="Back" title="Back" />
	        <% if Request.QueryString("fromModule") = "1" THEN %>
	        <a href="#" onclick="window.close();">Back to Location Module</a>
	        <% else %>
	        <a href="manage_locations.asp?sf=<%=cur_last%>&is_active=<%=cur_is_active %>">Back to Location Search</a>
	        <% end if %>
	        </td>
	    </tr>
	    <tr>
	        <td align="right" class="blackFontLarge">Location ID:</td>
	        <td align="left">
	            <input type="text" name="location" value="<%=cur_id%>" disabled="disabled" />
	        </td>
	    
	        <td align="right" class="blackFontLarge">Location Name:</td>
	        <td align="left">
	            <input type="text" name="location_name" value="<%=cur_name%>">
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
	            <input type="text" name="state" value="<%=cur_state%>">
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
	        <td align="right" class="blackFontLarge">Phone Ext:</td>
	        <td align="left">
	            <input type="text" name="phone_ext" value="<%=cur_phone_ext%>">
	        </td>
	        <td align="right" class="blackFontLarge"></td>
	        <td align="left">
	            
	        </td>              
	    </tr>
	    <tr>
	        <td align="right" class="blackFontLarge" bgcolor="#e0e0aa">Upload picture file</td>
	        <td align="left" class="blackFontLarge" colspan="2"><input type="file" name="target_file" size="" /></td>

            <td colspan="4" class="blackFontLarge" >
         
	       
	          <div style="padding-left:216px"> &nbsp;Is Active: <input type="checkbox" name="is_active" <%if cur_is_active = "1" THEN%>checked <%end if %> /></div>
	        </td>    
	        
	    </tr>
	    <tr>
	        <td colspan="6" align="center"><input type="button" onclick="submitForSave();" value="Save Changes" class="submit" /></td>
	    </tr>   
    </table>
</form>
</body>
</html>
