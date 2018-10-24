<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>iCentrix Corp. Electronic Medical Records</title>
    <link href="includes/styles.css" rel="stylesheet" type="text/css" />
<%copy_form_id = request.QueryString("uid") %>
<script type="text/javascript" src="js/tw-sack.js"></script>

<script>
    function filterStaff2()
    {
        if(frmMain2.form_program.value != '')
        {
            frmMain2.program_default.disabled = false;
        }
        else
        {
            frmMain2.program_default.disabled = true;
        }
        
        frmMain2.program_default.checked = false;
    
        Ajax('staff2','scripts/ShowStaffInProgram.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain2.form_program.value+'&usid=<%=Request.querystring("usid")%>');
    }
    function checkForm2()
    {
         if(frmMain2.form_program.value == '')
         {
            alert("Please select a Program, and Staff before creating a new form.");
                           
            return false;
         }
         
         if(frmMain2.form_staff.value == '')
         {
            alert("Please select a Staff before creating a new form.");
  
            return false;
         }
    
        var defaultProg = "";
        
        if(frmMain2.program_default.checked)
        {
            defaultProg = frmMain2.form_program.value; 
        }
            
        var chosenParentValue = "";
        var parentFormRequired = 0;
        var is_for_fix = get_radio_value();
    
        if(document.getElementById("parent_form_id"))
        {
             parentFormRequired = 1;
                
            chosenParentValue = getSelectedRadioValue(document.frmMain2.parent_form_id);
        }
        
        if('<%=Request.QueryString("lfid")%>' != '')
        {

   
            popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=Request.querystring("uid")%>&ft=<%=Request.QueryString("ft") %>&cid=' + '<%=Request.QueryString("cid")%>' + '&pid=' + frmMain2.form_program.value + '&sid='+ frmMain2.form_staff.value+'&dp='+defaultProg+'&lfid=<%=Request.QueryString("lfid") %>&if=' + is_for_fix ,800,1000,'<%=Request.QueryString("ft") %>_copy_' + '<%=Request.QueryString("cid")%>');
        }
        else
        {        
            if(chosenParentValue == "")
            {
                if(parentFormRequired == 1)
                {
                    alert("Please choose an associated form to create this form for");
                    return false;
                }
                else
                {



               <%if Request.QueryString("ft") = "AFCMPOC" then %>


                 popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=<%=Request.QueryString("ft") %>&cid=' + '<%=Request.QueryString("cid")%>' + '&pid=' + frmMain2.form_program.value + '&sid='+ frmMain2.form_staff.value +'&dp=&lfid=' + '<%=Request.querystring("uid") %>' + '&if=Fix',800,1000,'Request.QueryString("ft")>_copy_' + '<%=Request.QueryString("cid")%>');


                   
                <%else %>      
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=Request.querystring("uid")%>&ft=<%=Request.QueryString("ft") %>&cid=' + '<%=Request.QueryString("cid")%>' + '&pid=' + frmMain2.form_program.value + '&sid='+ frmMain2.form_staff.value+'&dp='+defaultProg + '&if=' + is_for_fix ,800,1000,'<%=Request.QueryString("ft") %>_copy_' + '<%=Request.QueryString("cid")%>');
    
    
                 <%end if %>       


                }            
            }
            else
            {
    
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=Request.querystring("uid")%>&ft=<%=Request.QueryString("ft") %>&cid=' + '<%=Request.QueryString("cid")%>' + '&pid=' + frmMain2.form_program.value + '&sid='+ frmMain2.form_staff.value+'&dp='+defaultProg+'&lfid='+chosenParentValue + '&if=' + is_for_fix ,800,1000,'<%=Request.QueryString("ft") %>_copy_' + '<%=Request.QueryString("cid")%>');
            }
        }        
    }
    
    function get_radio_value()
    {
        var rad_val = "";
        
       for (var i=0; i < document.frmMain2.is_for_fix.length; i++)
       {
            if (document.frmMain2.is_for_fix[i].checked)
            {
                rad_val = document.frmMain2.is_for_fix[i].value;
            }
       }
       
       return rad_val
    }

</script>
</head>
<body>
<form name="frmMain2" action="choose_parent_for_copy.asp" method="post">
  	      <div class="pod_sub_title" style="margin-top: 4px; margin-left: 2px;">Is this a fix for the current form?:&nbsp;&nbsp;&nbsp;&nbsp;
  	        <input type="radio" name="is_for_fix" value="Fix" / <%if Request.QueryString("ft")="AFCMPOC" then response.write "checked" end if %>>Yes&nbsp;&nbsp;&nbsp;&nbsp;
              <%if Request.QueryString("ft")<>"AFCMPOC" then %>
                  <input type="radio" name="is_for_fix" value="Copy" />No&nbsp;&nbsp;(If no is selected a new form will be created)
            <%end if %>
                </td>
		  </div><br /><br /><b>HINT: Use fix only when correcting information on an existing incorrect form.</b>
		  <br /><br />    
  	      <div class="pod_sub_title" style="margin-top: 4px; margin-left: 2px;">Please select the Associated Form from the list below:</div>
	      <br />
              <table id="childPrePop" border=0 align="center">
              
              <tr>
		        <td>&nbsp;</td>
		        <td>
		            <select name="form_program" onchange="filterStaff2();" <% if Request.QueryString("cid") = "" THEN%>disabled<%end if%>>
                    <option value="">Select a Program</option>
                        <%
                            if Request.QueryString("cid") <> "" THEN
                            
	                            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                            Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                            SQLStmt2.CommandText = "exec get_programs_for_staff_and_client '" & Session("user_name") & "'," & Request.QueryString("cid")  
  	                            SQLStmt2.CommandType = 1
  	                            Set SQLStmt2.ActiveConnection = conn
  	                            'response.write "SQL = " & SQLStmt2.CommandText
  	                            rs2.Open SQLStmt2
  	                            Do Until rs2.EOF
	                    %>
	                            <option value="<%=rs2("Program_ID")%>"><%Response.write rs2("Program_Name")%></option>
	                    <%
	                            rs2.MoveNext
                                Loop
	                        end if
	                    %>
                    </select>
                </td>
			    <td>
			        <div id="staff2">
                        <select name="form_staff" disabled>
                            <option value="">Choose a Program</option>
                        </select>
                    </div>
	            </td>
	            <td colspan="3" valign="middle"></td>
	        </tr>
	        <tr>
		        <td>&nbsp;</td>
		        <td valign="middle"><input type="checkbox" class="radio" name="program_default" disabled/> set as my default program</td>
		        <td class="blackFontSmall">&nbsp;</td>
		        <td colspan="4" valign="middle"></td>
		    </tr>
		    <%
		        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
    	        Set rs2 = Server.CreateObject ("ADODB.Recordset")    
  	            SQLStmt2.CommandText = "exec get_possible_parent_forms_for_type '" & Request.QueryString("ft") & "'," & Request.QueryString("cid") & "," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & "," & update_forms_require_finalized_parent & "," & addendum_forms_require_finalized_parent & "," & review_revision_forms_require_finalized_parent & "," & notes_require_finalized_iap
  	            SQLStmt2.CommandType = 1
  	            Set SQLStmt2.ActiveConnection = conn
  	            'response.write "SQL = " & SQLStmt2.CommandText
  	            rs2.Open SQLStmt2
  	            
                if rs2.EOF or rs2.BOF or Request.QueryString("lfid") <> "" THEN
                    'NO PARENTS NEEDED
                else
		    %>
	        <tr>
		      <th valign="top" align="left" class="blackFontSmall"><b>Form Name:</b></th>
		      <th valign="top" align="left" class="blackFontSmall"><b>Status:</b></th>
		      <th valign="top" align="left" class="blackFontSmall"><b>Created On:</b></th>
		      <th valign="top" align="left" class="blackFontSmall"><b>Created By:</b></th>
		      <th valign="top" align="left" class="blackFontSmall"><b>Program:</b></th>
		      <th valign="top" align="left" class="blackFontSmall"><b>Staff:</b></th>
		      <th valign="top" align="center" class="pod_sub_title">Select:</th>
	        </tr>
	        <%
  	                rsArray = rs2.GetRows() 
                    nr = UBound(rsArray, 2) + 1
  	                cur_rec_count = nr
      	            
  	                rs2.MoveFirst
      	            
  	                Do Until rs2.EOF    
      	             	                
  	                    cur_parent_id = rs2("Unique_Form_ID")
  	                    cur_root_info = rs2("root_form_info")
  	                    cur_date_date = rs2("Update_Date")
	        %>
	        <tr>
		      <td align="left" class="displayFontMed"><%=cur_root_info %>&nbsp;<%=rs2("Form_Name")%></td>
		      <td align="left" class="displayFontMed"><%=rs2("Status")%>&nbsp;<%=cur_date_date %></td>
		      <td align="left" class="displayFontMed"><%=rs2("Create_Date")%></td>
		      <td align="left" class="displayFontMed"><%=rs2("Create_User_Name")%></td>
		      <td align="left" class="displayFontMed"><%=rs2("Main_Program_Name")%></td>
		      <td align="left" class="displayFontMed"><%=rs2("Main_Staff_Name")%></td>
		      <td align="center"><input type="radio" <%if cur_rec_count = 1 THEN%>checked<%end if%> class="radio" id="parent_form_id" name="parent_form_id" value="<%=cur_parent_id%>" /></td>
		    </tr>
	        <%
	                    rs2.MoveNext
                    Loop
                end if
	        %>
	        <tr>
	        <td colspan="7"><a href="#" onclick="checkForm2();" id="add_new_form_button"><img src="images/button_add_new.jpg" width="28" height="16" border="0" class="textmiddle" alt="Create New Form" title="Create New Form" /></a></td>
	        </tr>
	        </table>
    </form>
</body>
</html>