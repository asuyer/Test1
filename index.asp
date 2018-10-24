<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<%
'PAGE MAIN SECURITY CHECK
if Session("active_staff_can_access_client_module") <> "1" THEN
    response.Redirect "login.asp"
end if 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

     <link rel="stylesheet" href="includes/jquery-ui-1.12.1/jquery-ui.min.css" type="text/css" />
<script src="includes/jquery-ui-1.12.1/external/jquery/jquery.js" type="text/javascript"></script>
<script src="includes/jquery-ui-1.12.1/jquery-ui.js" type="text/javascript"></script>
   
  
<link rel="stylesheet" href="includes/thickbox.css" type="text/css" media="screen" />
    <script type="text/javascript" src="includes/thickbox.js"></script>
<script type="text/javascript" src="includes/nav.js"></script>
<script type="text/javascript" src="js/tw-sack.js"></script>
<%
'*****CALC THESE AFTER*****'
cur_month = Month(Date())
cur_year = Year(Date())




if Request.QueryString("action") = "schedule" then


    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
Set rs2 = Server.CreateObject ("ADODB.Recordset")

SQLStmt2.CommandText = "exec insert_next_doctor_visit " & Request.QueryString("cid") & "," & Request.QueryString("uid") & ",'" & Request.QueryString("date") & "','" & Request.QueryString("time") & "'"
SQLStmt2.CommandType = 1
Set SQLStmt2.ActiveConnection = conn
SQLStmt2.CommandTimeout = 45 'Timeout per Command
    ' response.write SQLStmt2.CommandText
rs2.Open SQLStmt2

    

   response.redirect "index.asp?cid=" & Request.QueryString("cid")
end if


hcp_upload = "-1"

if Request.QueryString("hcp_upload") <> "" then
hcp_upload = Request.QueryString("hcp_upload")
end if

    if Request.QueryString("cid") <> "" THEN
       Set SQLStmt3 = Server.CreateObject("ADODB.Command")
    Set rs3 = Server.CreateObject ("ADODB.Recordset")
         
    SQLStmt3.CommandText = "exec get_hcp_inprocess_encounter_forms_count  " & Request.QueryString("cid")
    SQLStmt3.CommandType = 1
    Set SQLStmt3.ActiveConnection = conn
                    'response.write "SQL = " & SQLStmt3.CommandText
    rs3.Open SQLStmt3

    hcef_count = rs3("hcef_count")
end if 




cur_call_log_id = ""

Set SQLStmt2 = Server.CreateObject("ADODB.Command")
Set rs2 = Server.CreateObject ("ADODB.Recordset")

SQLStmt2.CommandText = "select top 1 unique_form_id from forms_master where form_type = 'CALL_LOG'"
SQLStmt2.CommandType = 1
Set SQLStmt2.ActiveConnection = conn
SQLStmt2.CommandTimeout = 45 'Timeout per Command
rs2.Open SQLStmt2

Do until rs2.EOF
    cur_call_log_id = rs2("unique_form_id")
rs2.MoveNext
Loop

    if Request.QueryString("cid") <> "" THEN

        if Request.QueryString("fn") = "" THEN
            cur_client_prog = "-1"
        else
            cur_client_prog = Request.QueryString("fn")
        end if
        
        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	    Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	    SQLStmt2.CommandText = "get_client_info " & Request.QueryString("cid") & "," & cur_client_prog
  	    SQLStmt2.CommandType = 1
  	    Set SQLStmt2.ActiveConnection = conn
  	    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	   ' response.write "SQL = " & SQLStmt2.CommandText
  	    rs2.Open SQLStmt2
  	
  	    client_name = rs2("Last_Name") & ", " & rs2("First_Name")
  	    client_dob = rs2("DOB")
  	    client_ssn = rs2("SSN")
  	    client_id = rs2("client_id")
  	    client_reg = rs2("Registration_Date")
  	    client_st = rs2("Street")
  	    client_city = rs2("City")
  	    client_state = rs2("State")
  	    client_zip = rs2("Zip")
  	    client_external_info = rs2("external_info")
  	    client_consents_count = rs2("client_consents_count")
  	    client_correspondence_count = rs2("client_correspondence_count")
  	    client_court_docs_count = rs2("client_court_docs_count")
  	    client_reg_packets_count = rs2("client_reg_packets_count")
  	    client_prior_paper_count = rs2("client_prior_paper_count")
  	    client_medical_count = rs2("client_medical_count")
  	    client_misc_id_docs_count = rs2("client_misc_id_docs_count")
  	    client_misc_fin_docs_count = rs2("client_misc_fin_docs_count")
  	    client_media_releases_count = rs2("client_media_releases_count")
  	    client_satisfaction_surveys_count = rs2("client_satisfaction_surveys_count")
  	    client_drills_assessments_count = rs2("client_drills_assessments_count")
  	    client_ot_pt_spl_count = rs2("client_ot_pt_spl_count")
  	    client_masshealth_critical_incidents_count = rs2("client_masshealth_critical_incidents_count")  	    
  	    client_afc_letter_count = rs2("client_afc_letter_count")
  	    client_entitlements_count = rs2("client_entitlements_count")
  	    client_isp_count = rs2("client_isp_count")
        client_hcsis_id = rs2("hcsis_id")
  	    next_isp_date = rs2("next_isp_date")
  	    is_dnr = rs2("is_dnr")
        isp_assessments_status = rs2("isp_assessments_status")
        dhsp_date = rs2("dhsp_date")
        intake_program_id= rs2("intake_program_id")
  	    
  	    
  	    if use_cmhc_id_for_recnum = 1 THEN
            rec_num_id = client_external_info
        else
            rec_num_id = client_id
        end if
  	    
  	    sdate = client_dob
  	    
  	    temp_result = DateDiff("d",sdate,Date)
  	    
  	    temp_age = temp_result / 365
  	    
  	    temp_age2 = Round(temp_age)

	    if temp_age2 > temp_age then 
	        temp_age2 = temp_age2 - 1
        end if
        
        client_age = temp_age2
        total_form_count = rs2("Total_Form_Count")
        total_encounter_count = rs2("Total_Encounter_Count")
        in_proc_count = rs2("In_Proc_Count")
        finalized_count = rs2("Finalized_Count")
        active_goals_count = rs2("Active_Goals_Count")
        active_obj_count = rs2("Active_Objectives_Count")
        active_needs_count = rs2("Active_Needs_Count")
        last_update_date = rs2("Last_Update")
        last_encounter_date = rs2("Last_Encounter")
        last_update_user = rs2("Last_Update_User")
        last_encounter_user = rs2("Last_Encounter_User")
        last_update_form = rs2("Last_Update_Form")
        last_encounter_form = rs2("Last_Encounter_Form")
        last_signer = rs2("Last_Signer")
        last_sig_date = rs2("Last_Sig_Date")
        last_sig_form = rs2("Last_Sig_Form")
        max_pif_date_diff = rs2("max_pif_date_diff")
        cur_ep_num = rs2("cur_ep_number")
        cur_ep_id = rs2("cur_ep_id")
        cur_ep_start = rs2("cur_ep_start")
        client_has_picture = rs2("has_picture")	
        client_meds_count = rs2("client_meds_count")	
        client_diags_count = rs2("client_diags_count")       
  	else
  	
  	    client_name = "&nbsp;"
  	    client_dob = "&nbsp;"
  	    client_ssn = "&nbsp;"
  	    client_id = "&nbsp;"
  	    client_reg = "&nbsp;"
  	    client_st = "&nbsp;"
  	    client_city = "&nbsp;"
  	    client_state = "&nbsp;"
  	    client_zip = "&nbsp;"
  	    client_age = "&nbsp;"
  	    cur_ep_id = -1000
  	end if
%>

	<title>MSDP | Electronic Healthcare Forms</title>
<meta name="keywords" content="" /> 
<meta name="description" content="" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Style-Type" content="text/css" />
<meta http-equiv="Cache-Control" content="no-cache" />
<meta http-equiv="Pragma" content="no-cache" />
<link href="includes/nav.css" rel="stylesheet" type="text/css" />

<!--[if IE 6]>
<link rel="stylesheet" type="text/css" href="includes/nav_ie6.css" />
<![endif]-->
<!--[if IE 7]>
<link rel="stylesheet" type="text/css" href="includes/nav_ie6.css" />
<![endif]-->
<link href="includes/styles.css" rel="stylesheet" type="text/css" />

<%
    Set SQLStmtS = Server.CreateObject("ADODB.Command")
  	Set rsS = Server.CreateObject ("ADODB.Recordset")
    SQLStmtS.CommandText = "select *, (select Group_Name from Program_Group_Master where group_id = (select top 1 group_id from program_group_assign where program_id = s.Default_Program)) as Default_Program_Group from staff_master s where user_name = '" & Session("user_name") & "'"  
    SQLStmtS.CommandType = 1
    Set SQLStmtS.ActiveConnection = conn
    SQLStmtS.CommandTimeout = 45 'Timeout per Command
    'response.write "SQL = " & SQLStmt2.CommandText
    rsS.Open SQLStmtS
    
    user_staff_id = rsS("staff_id")
    cur_last_login = rsS("last_login")
    user_default_prog = rsS("Default_Program")
    user_default_prog_group = rsS("Default_Program_Group")

    client_count = 0
    total_client_forms = 0

    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	Set rs2 = Server.CreateObject ("ADODB.Recordset")
    SQLStmt2.CommandText = "exec get_staff_roles '" & Session("user_name") & "'"  
    SQLStmt2.CommandType = 1
    Set SQLStmt2.ActiveConnection = conn
    SQLStmt2.CommandTimeout = 45 'Timeout per Command
    'response.write "SQL = " & SQLStmt2.CommandText
    rs2.Open SQLStmt2
    
    cur_staff_role = rs2("Role_ID")
    cur_staff_role_desc = rs2("Role_Description")
    create_client_auth = rs2("Create_Client_Auth")
    manage_staff_auth = rs2("Manage_Staff_Auth")
    manage_programs_auth = rs2("Manage_Programs_Auth")
    manage_roles_auth = rs2("Manage_Roles_Auth")
    
      'GET CLIENTS FOR THIS STAFF MEMBER    
	Set SQLStmtClientsForStaff = Server.CreateObject("ADODB.Command")
  	Set rsClientsForStaff = Server.CreateObject ("ADODB.Recordset")

  	SQLStmtClientsForStaff.CommandText = "exec get_clients_for_staff '" & Session("user_name") & "','" & Request.QueryString("pcode") & "','" & Request.QueryString("lCode") & "'"   
  	SQLStmtClientsForStaff.CommandType = 1
  	Set SQLStmtClientsForStaff.ActiveConnection = conn
  	SQLStmtClientsForStaff.CommandTimeout = 45 'Timeout per Command
    'if Session("user_name") = "pwcard" THEN
    '    response.write "sql = " & SQLStmtClientsForStaff.CommandText
    'end if 
  	rsClientsForStaff.Open SQLStmtClientsForStaff
  	Do Until rsClientsForStaff.EOF
	        client_count = client_count + 1
	       ' total_client_forms = total_client_forms + rsClientsForStaff("client_form_count")
	rsClientsForStaff.MoveNext
    Loop
    
    total_alerts_for_header = 0
    
    client_id_for_alerts = -1
    
    if Request.QueryString("cid") <> "" THEN
        client_id_for_alerts = Request.QueryString("cid")
    end if
    
    'GET ALERTS FOR THIS CLIENT/STAFF
    Set SQLStmtAlertsCount = Server.CreateObject("ADODB.Command")
  	Set rsAlerts = Server.CreateObject ("ADODB.Recordset")
    SQLStmtAlertsCount.CommandText = "exec get_alerts " & user_staff_id & "," & client_id_for_alerts & ", " & cur_ep_id
  	SQLStmtAlertsCount.CommandType = 1
  	Set SQLStmtAlertsCount.ActiveConnection = conn
  	SQLStmtAlertsCount.CommandTimeout = 45 'Timeout per Command
  	'if Session("user_name") = "pwcard" THEN
    '  	response.write "SQL = " & SQLStmtAlertsCount.CommandText
  	'end if
  	rsAlerts.Open SQLStmtAlertsCount
    Do Until rsAlerts.EOF
	        total_alerts_for_header = total_alerts_for_header + 1
	rsAlerts.MoveNext
    Loop 
  	  	
%>
<script type="text/javascript">
    function confirmFileDelete(fname, fid) {



        $.ajax({url: "scripts/get_child_form_names.asp?lfid=" + fid, success: function(child_forms){

       
        if (confirm("Are you sure you want to delete " + fname + child_forms + "?")) {
            var url = "delete_form.asp?cid=" + '<%=request.queryString("cid")%>' + "&fid=" + fid;
     
              $.ajax({url: url, success: function(result){

              location.reload();

                }});

              }

            }});

    }
	
	function clickToClearAlert(alert_id)
	{
	    if(confirm("Are you sure you want to clear this system message?"))
		{
			var url = "click_to_clear_alert.asp?aid="+ alert_id + "&cid=" + '<%=Request.QueryString("cid")%>';
			window.location.href=url;
		}
		else
		{
		    return false;
		}
	}	
	
    function refreshProg(newProg)
    {
        window.location.href = "index.asp?pCode=" + newProg + "&sd=" + clientForm.show_discharged.checked;
    }
    function refreshLocation(newLocation)
    {
        window.location.href = "index.asp?pCode=" + document.clientForm.program.value + "&lCode=" + newLocation + "&sd=" + clientForm.show_discharged.checked;
    }
    
    function refreshClient(newClient)
    {
        Ajax('create_form_div','&nbsp;');
    
        window.location.href= "index.asp?cid="+ newClient + "&pCode=" + document.clientForm.program.value + "&lCode=" + document.clientForm.location.value + "&sd=" + clientForm.show_discharged.checked;
    } 
    
    function filterStaff()
    {
        if(frmMain.form_program.value != '')
        {
            frmMain.program_default.disabled = false;
        }
        else
        {
            frmMain.program_default.disabled = true;
        }
        
        frmMain.program_default.checked = false;
    
        Ajax('staff','scripts/ShowStaffInProgram.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&usid=<%=user_staff_id%>');
    }
    
  function undoFilter(formID, linkedFormID, formType)
    {
        window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=ALL&sd=" + clientForm.show_discharged.checked + "&ufid=" + formID + "&ulfid=" + linkedFormID + "&uft=" + formType + '&show_dashboard_flag=' + '<%=Request.QueryString("show_dashboard_flag")%>' + '&soft=' + formType;
    }


   function undoFilterInner(formID, linkedFormID, formType,innerType)
    {
        window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=ALL&sd=" + clientForm.show_discharged.checked + "&ufid=" + formID + "&ulfid=" + linkedFormID + "&uft=" + innerType + '&show_dashboard_flag=' + '<%=Request.QueryString("show_dashboard_flag")%>' + '&soft=' + formType;
    }

    function showFormType(formType) {
        window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=" + '<%=Request.QueryString("tf")%>' + "&sd=" + clientForm.show_discharged.checked + "&soft=" + formType + '&show_dashboard_flag=' + '<%=Request.QueryString("show_dashboard_flag")%>';
    }

    function checkFormCopyPrevious() {
        if (frmMain.create_form_type.value == '') {
            alert("You must choose a form type in order to copy it from the previous episode");
        }
        else {
            Ajax('possible_parent_forms', 'scripts/ShowPossibleFormsForCopy.asp?cid=<%=Request.QueryString("cid")%>&ft=' + frmMain.create_form_type.value);
        }
    }
    
   function reloadCaseNoteType(formValue,csValue)
    {
        Ajax('possible_parent_forms','scripts/ShowPossibleParentsForForm.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&cft=' + formValue + '&cnd=' + csValue);
     
		var possibleParentButtons = document.getElementsByName("parent_form_id");
	    
	    if (possibleParentButtons.length == 1)
	    {
	        possibleParentButtons[0].checked =true;
	    }
	    else
	    {
	        //alert("found more than 1 button");
	    } 
    }  

    function reloadFormParents(formValue)
    {


        


        Ajax('possible_parent_forms','scripts/ShowPossibleParentsForForm.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&cft=' + formValue);
     
		var possibleParentButtons = document.getElementsByName("parent_form_id");
	    
	    if (possibleParentButtons.length == 1)
	    {
	        possibleParentButtons[0].checked =true;
	    }
	    else
	    {
	        //alert("found more than 1 button");
	    } 


   }  


      function reloadFormParentsCHK(formValue)
    {



     

         if ($("#Prior_DHSP").prop('checked')) {

               Ajax('possible_parent_forms','scripts/ShowPossibleParentsForForm.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&cft=' + $("select[name=create_form_type]").val() + '&pd=1');
         } else {
              Ajax('possible_parent_forms','scripts/ShowPossibleParentsForForm.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&cft=' + $("select[name=create_form_type]").val() + '&pd=0');
        }

    
     
		    var possibleParentButtons = document.getElementsByName("parent_form_id");
	    
	        if (possibleParentButtons.length == 1)
	        {
	            possibleParentButtons[0].checked =true;
	        }
	        else
	        {
	            //alert("found more than 1 button");
	        } 






    }   
 










 function confirmFileDelete(fname, fid) {



        $.ajax({url: "scripts/get_child_form_names.asp?lfid=" + fid, success: function(child_forms){

       
        if (confirm("Are you sure you want to delete " + fname + child_forms + "?")) {
            var url = "delete_form.asp?cid=" + '<%=request.queryString("cid")%>' + "&fid=" + fid;
     
              $.ajax({url: url, success: function(result){

              location.reload();

                }});

              }

            }});

    }


    function resetStaffDefault(staffValue)
    {
    }
    var enableCache = false;
    var jsCache = new Array();
    var AjaxObjects = new Array();

    function ShowContent(divId,ajaxIndex,url)
    {
	    document.getElementById(divId).innerHTML = AjaxObjects[ajaxIndex].response;
	    if(enableCache){
	     jsCache[url] = 	AjaxObjects[ajaxIndex].response;
	    }
	    AjaxObjects[ajaxIndex] = false;
    }

    function Ajax(divId,url)
    {
        if (enableCache && jsCache[url])
	    {
	     document.getElementById(divId).innerHTML = jsCache[url];
	     return;
	    }	
	    var ajaxIndex = AjaxObjects.length;
	    document.getElementById(divId).innerHTML = '<img src=images/movewait.gif width=16 height=16 hspace=10 vspace=10 />';
	    AjaxObjects[ajaxIndex] = new sack();
	    AjaxObjects[ajaxIndex].requestFile = url;
	    AjaxObjects[ajaxIndex].onCompletion = function(){ ShowContent(divId,ajaxIndex,url);

        $( "#start_date" ).datepicker();

         
       };
	    AjaxObjects[ajaxIndex].runAJAX();
    }
    
    function getSelectedRadio(buttonGroup) 
    {
        // returns the array number of the selected radio button or -1 if no button is selected
        if (buttonGroup[0]) 
        { // if the button group is an array (one button is not an array)
            for (var i=0; i<buttonGroup.length; i++) 
            {
                if (buttonGroup[i].checked) 
                {
                    return i
                }
            }
        } 
        else 
        {
            if (buttonGroup.checked) 
            { 
                return 0; 
            } // if the one button is checked, return zero
        }
        // if we get to this point, no radio button is selected
        return -1;
    } // Ends the "getSelectedRadio" function

function getSelectedRadioValue(buttonGroup) 
{
   // returns the value of the selected radio button or "" if no button is selected
   var i = getSelectedRadio(buttonGroup);
   if (i == -1) 
   {
    //alert("no button selected");
      return "";
   } 
   else 
   {
      if (buttonGroup[i]) 
      { // Make sure the button group is an array (not just one button)
         return buttonGroup[i].value;
      } 
      else 
      { // The button group is just the one button, and it is checked
         return buttonGroup.value;
      }
   }
} // Ends the "getSelectedRadioValue" function
       
    function checkForm()
    {

        var pifReq = 0;
        var pifReqAge = 0;
        
        if(clientForm.client_id.value == '')
        {
            alert("Please select a Client, Program, Staff, and Form Type before creating a new form.");
            return false;
        }
         
        if(frmMain.form_program.value == '')
        {
            if(frmMain.create_form_type.value != '')
            {
                alert("Please select a Program, and Staff before creating a new form.");
            }
            else
            {
                alert("Please select a Program, Staff, and Form Type before creating a new form.");
            } 
                    
            return false;
        }
        else
        {
            var selIndex = frmMain.form_program.selectedIndex;

            pifReq = frmMain.form_program.options[selIndex].pif_age_require;
            pifReqAge = frmMain.form_program.options[selIndex].pif_max_age_hours;
        }
         
        if(frmMain.form_staff.value == '')
        {
            if(frmMain.create_form_type.value != '')
            {
                alert("Please select a Staff before creating a new form.");
            }
            else
            {
                alert("Please select a Staff, and Form Type before creating a new form.");
            } 
                    
            return false;
        }
         
        if(frmMain.create_form_type.value == '')
        {
            alert("Please select a Form Type before creating a new form.");
            return false;
        }
         
        if(pifReq == '1')
        {
            if( (frmMain.create_form_type.value == 'CCA' ||
                frmMain.create_form_type.value == 'ACA' ||
                frmMain.create_form_type.value == 'ESPCCA' ||
                frmMain.create_form_type.value == 'ESPACA') &&
                parseInt('<%=max_pif_date_diff%>') > pifReqAge && 
                '<%=max_pif_date_diff%>' != '-1' )
            {
           
                alert("You must create a more recent Personal Information Form prior to creating a new Assessment.");
                return false;
            }
        }
    
        var defaultProg = "";
        
        if(frmMain.program_default.checked)
        {
            defaultProg = frmMain.form_program.value; 
        }


        var chosenQuarterValue = "";
        var chosenMonthValue = "";
        var chosenYearValue = "";
        var chosenScoreCardTypeValue = "";
        var chosenStartDateValue = "";
         var chosenActionValue = "";
		var chosenDateValue = "";
		var chosenParentValue = "";
        var chosenPhysicianValue = "";
        var chosenCaseNoteType = ""
        var chosenNurseVisitNoteMonth = ""
         var chosenCaseNoteLowDate = ""
        var chosenCaseNoteHighDate = ""
        var chosenSpecialityValue = "";
        var chosenMonthYearValue = "";
        var chosenGoalValue = "";
        var chosenSpecialityValue = "";
        var chosenISPGoalValue = "";
        var chosenISPObjValue = "";
        var parentFormRequired = 0;
        var parentGoalRequired = 0
        var parentPhysicianRequired = 0;
        var parentNurseVisitNoteMonthRequired = 0;
        var parentCaseNoteTypeRequired = 0;
        var parentCaseNoteLowDateRequired = 0;
        var parentCaseNoteHighDateRequired = 0;
        var parentSpecialityRequired = 0;
        var parentScoreCardTypeRequired = 0;
        var parentMonthYearRequired = 0;
        var parentSpecialityRequired = 0;
        var parentISPGoalRequired = 0;
        var parentISPObjRequired = 0;
        var parentStartDateRequired = 0;
        var popup_path = "";
        var popup_path1 = "";
        var popup_path2 = "";
        var popup_path3 = "";




         if(document.getElementById("action"))
        {
            
                
           chosenActionValue = document.frmMain.action.value;


     
           if(chosenActionValue=="Continue") {
              //alert(chosenActionValue);

                  popupwindow('http://<%=url_org_name%>/copy_dhsp_into_form.asp?cid=' + clientForm.client_id.value,800,1000,frmMain.create_form_type.value + '' + clientForm.client_id.value);
                  return false;

           }

        }




      if(document.getElementById("start_date"))
        {
            
            parentStartDateRequired = 1 
           chosenStartDateValue = document.frmMain.start_date.value;
           chosenStartDateValue = chosenStartDateValue.replace("/","_");
           chosenStartDateValue = chosenStartDateValue.replace("/","_");
           chosenStartDateValue = chosenStartDateValue.replace("/","_");
           
        }


           if(document.getElementById("quarter"))
        {
            
                
           chosenQuarterValue = document.frmMain.quarter.value;

        }

         if(document.getElementById("month"))
        {
            
                
           chosenMonthValue = document.frmMain.month.value;

        }


      if(document.getElementById("year"))
        {
            
                
           chosenYearValue = document.frmMain.year.value;

        }


        if(document.getElementById("parent_isp_goal"))
        {
            parentISPGoalRequired = 1;
                
            chosenISPGoalValue = getSelectedRadioValue(document.frmMain.parent_isp_goal);

        }

       if(document.getElementById("date_of_service"))
        {
            parentNurseVisitNoteMonthRequired = 1;
                
            chosenNurseVisitNoteMonth = document.frmMain.date_of_service.value;

        }

        if(document.getElementById("parent_isp_obj"))
        {
            parentISPObjRequired = 1;
                
            chosenISPObjValue = getSelectedRadioValue(document.frmMain.parent_isp_obj);

        }

        if(document.getElementById("goal_obj"))
         {
            parentISPGoalRequired = 1;
                
            chosenISPGoalValue = getSelectedRadioValue(document.frmMain.goal_obj);

        }



        if (chosenStartDateValue == "" && parentStartDateRequired == 1) 
        {
            alert("Please choose an associated Goal to create this form for");
            return false;
        }

       if (chosenISPGoalValue == "" && parentISPGoalRequired == 1) 
        {
            alert("Please choose an associated Goal to create this form for");
            return false;
        }

        if (chosenISPObjValue == "" && parentISPObjRequired == 1) 
        {
            alert("Please choose an associated Objective to create this form for");
            return false;
        }

        if (chosenISPObjValue == "" && parentISPObjRequired == 1) 
        {
            alert("Please choose an associated Objective to create this form for");
            return false;
        }

        if (chosenISPGoalValue != "") 
        {
            popup_path = '&ispgid=' + chosenISPGoalValue;
        }

        if (chosenISPObjValue != "") 
        {
            popup_path = '&ispoid=' + chosenISPObjValue;
        }

      
        if (document.getElementById("score_card_type")) {
            parentScoreCardTypeRequired = 1;
            
            chosenScoreCardTypeValue = document.frmMain.score_card_type.value;
        }

         if (document.getElementById("date_month_year")) {
           
            parentMonthYearRequired = 1;
            chosenMonthYearValue = document.frmMain.date_month_year.value;
        }


          if (document.getElementById("speciality_id")) {
            parentSpecialityRequired = 1;
            chosenSpecialityValue = document.frmMain.speciality_id.value;
            chosenSMADateFrom = document.frmMain.sma_date_from.value;
            chosenSMADateTo = document.frmMain.sma_date_to.value;
        }
        if(document.getElementById("parent_form_id"))
        {
            parentFormRequired = 1;
                
            chosenParentValue = getSelectedRadioValue(document.frmMain.parent_form_id);
        }
        
        if (document.getElementById("parent_physician_id")) 
        {
            parentPhysicianRequired = 1;
            chosenPhysicianValue = document.frmMain.parent_physician_id.value;
        }
        if (document.getElementById("parent_goal_name")) {
            parentGoalRequired = 1;
            chosenGoalValue = getSelectedRadioValue(document.frmMain.parent_goal_name);
        }
		if(document.getElementById("date_month_year"))
        {
            parentDateRequired = 1;
                
            chosenDateValue = document.frmMain.date_month_year.value;
        }
        if (chosenScoreCardTypeValue == "" && parentScoreCardTypeRequired == 1) {
            alert("Please choose an associated Score Card Type to create this form for");
            return false;
        }
       if (chosenNurseVisitNoteMonth == "" && parentNurseVisitNoteMonthRequired == 1) {
            alert("Please choose an associated Month to create this form for");
            return false;
        }

        if (chosenPhysicianValue == "" && parentPhysicianRequired == 1) 
        {
            alert("Please choose an associated Physician to create this form for");
            return false;
        }
          if (document.getElementById("parent_casenote_type_id")) {
            parentCaseNoteTypeRequired = 1;
            chosenCaseNoteType = document.frmMain.parent_casenote_type_id.value;
        }
        if (document.getElementById("casenote_low_daterange")) {
            parentCaseNoteLowDateRequired = 1;
            chosenCaseNoteLowDate = document.frmMain.casenote_low_daterange.value;
           // alert(chosenCaseNoteLowDate);
        }
        if (chosenGoalValue == "" && parentGoalRequired == 1) {
            alert("Please choose an associated Goal to create this form for");
            return false;
        }
        if (document.getElementById("casenote_high_daterange")) {
            parentCaseNoteHighDateRequired = 1;
            chosenCaseNoteHighDate = document.frmMain.casenote_high_daterange.value;
        }
        if (document.getElementById("speciality_id")) {
            parentSpecialityRequired = 1;
            chosenSpecialityValue = document.frmMain.speciality_id.value;
            chosenSMADateFrom = document.frmMain.sma_date_from.value;
            chosenSMADateTo = document.frmMain.sma_date_to.value;
        }

        if (chosenSpecialityValue == "" && parentSpecialityRequired == 1) {
            alert("Please choose an associated Speciality to create this form for");
            return false;
        }

       

        if (chosenSpecialityValue != "") {
            popup_path = '&speciality=' + chosenSpecialityValue + '&sdf=' + chosenSMADateFrom + '&sdt=' + chosenSMADateTo;
        }

       if (chosenQuarterValue != "") {
            popup_path = '&qt=' + chosenQuarterValue  + '&mo=' + chosenMonthValue + '&yr=' + chosenYearValue;
        }


        if (chosenGoalValue != "") {
            popup_path = '&lpid=' + chosenGoalValue;
        }

        if (chosenPhysicianValue != "")
           
        {
             if (frmMain.create_form_type.value == "MR_MO") {
             popup_path = '&lpid=' + chosenPhysicianValue 
             } else {
              popup_path = '&lpid=' + chosenPhysicianValue + '&bf=' + document.frmMain.blank_form.checked;
              }
           
        }
		
		if (chosenDateValue != "") 
        {
            popup_path = '&lpid=' + chosenDateValue;

        }
       if (chosenNurseVisitNoteMonth != "") {
            popup_path = '&my=' + chosenNurseVisitNoteMonth;
        }
        if (chosenScoreCardTypeValue != "") {
           
            popup_path = '&sc=' + chosenScoreCardTypeValue;
        }

         if (chosenMonthYearValue != "") {
            popup_path = '&my=' + chosenMonthYearValue;
        }

          if (chosenCaseNoteType == "" && parentCaseNoteTypeRequired == 1) {
            alert("Please choose an associated Case Note Type to create this form for");
            return false;
        }
        if (chosenCaseNoteLowDate == "" && parentCaseNoteLowDateRequired == 1) {
            alert("Please choose an associated Starting Date Range to create this form for");
            return false;
        }
        if (chosenCaseNoteHighDate == "" && parentCaseNoteHighDateRequired == 1) {
            alert("Please choose an associated Starting Date Range to create this form for");
            return false;
        }


         if (chosenStartDateValue != "") {

          
          
          
            popup_path = '&sd=' + chosenStartDateValue;
        }

		
		   if (chosenCaseNoteType != "") {

          
          var res = chosenCaseNoteType.replace(" ", "_"); 
          
            popup_path1 = '&lpid=' + res;
        }


        if (chosenCaseNoteLowDate != "") {

                      
            popup_path2 = '&hd=' + chosenCaseNoteLowDate;
        }

        if (chosenCaseNoteHighDate != "") {
           
            popup_path3 = '&ld=' + chosenCaseNoteHighDate;
        }
       
        if(chosenParentValue == "")
        {
            if(parentFormRequired == 1)
            {
                alert("Please choose an associated form to create this form for");
                return false;
            }
            else
            {
                if (frmMain.create_form_type.value == "BTMS" || frmMain.create_form_type.value == "DMTC" || frmMain.create_form_type.value == "GCFTR" || frmMain.create_form_type.value == "CLOG" || frmMain.create_form_type.value == "CONTRACTS" || frmMain.create_form_type.value == "DHSPS" || frmMain.create_form_type.value == "FDIP" || frmMain.create_form_type.value == "FFSE" || frmMain.create_form_type.value == "FMTP" || frmMain.create_form_type.value == "FTR" || frmMain.create_form_type.value == "HIPPA" || frmMain.create_form_type.value == "IMMUNE" ||frmMain.create_form_type.value == "FTR_CASH" || frmMain.create_form_type.value == "FTR_BANK" || frmMain.create_form_type.value == "GCFTR" || frmMain.create_form_type.value == "MEDLIST" || frmMain.create_form_type.value == "MRC" || frmMain.create_form_type.value == "RTF" || frmMain.create_form_type.value == "SPMU" || frmMain.create_form_type.value == "MS" || frmMain.create_form_type.value == "PAS" || frmMain.create_form_type.value == "MS_V2" || frmMain.create_form_type.value == "MS_PROC" || frmMain.create_form_type.value == "LPOF" || frmMain.create_form_type.value == "MMR" || frmMain.create_form_type.value == "SLMMS")
                {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid='+ frmMain.form_staff.value+'&dp='+defaultProg + popup_path,800,1400,frmMain.create_form_type.value + '' + clientForm.client_id.value);
                }
				 else if (frmMain.create_form_type.value == "CNDR" || frmMain.create_form_type.value == "ECNDR" || frmMain.create_form_type.value == "NDNDR") {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path1 + popup_path2 + popup_path3, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
            else if (frmMain.create_form_type.value == "VIEW_MA") {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path, 800,1300, frmMain.create_form_type.value + '' + clientForm.client_id.value);
               
                }  

           else if (frmMain.create_form_type.value == "IIRSN") {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path2 + popup_path3, 800, 1110, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
            else if (frmMain.create_form_type.value == "ENCMN") {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }

            else if (frmMain.create_form_type.value == "DDS_ISP_MPS") {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&isp=' + '<%=next_isp_date %>' + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
                }

           else if (frmMain.create_form_type.value == "MR_MO" || frmMain.create_form_type.value == "ISPPS" || frmMain.create_form_type.value == "ISPR" || frmMain.create_form_type.value == "ISPPN" || frmMain.create_form_type.value == "ISP_PSS") {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
                 }
         


            else if (frmMain.create_form_type.value == "SC" || frmMain.create_form_type.value == "QNS_ADH") {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            } 
            else if (frmMain.create_form_type.value == "NURSE_SUM") {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            } 

              else if (frmMain.create_form_type.value == "DDS_ISP_WGDT") 
               {
                var obj_name = ""
                var goal_name = ""
                var id_parts;

                id_parts = chosenISPGoalValue.split("_");
                goal_name = id_parts[0];
                obj_name = id_parts[1];
                lfid = id_parts[2];
               

                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + '&lfid=' + lfid + '&goal_name=' + goal_name + '&obj_name=' + obj_name, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
                }  
           else if (frmMain.create_form_type.value == "DCF" || frmMain.create_form_type.value == "PSA_DDS" || frmMain.create_form_type.value == "PSS" || frmMain.create_form_type.value == "DHSWGDC") {
                var main_id_part = ""
                var goal_name = ""
                var id_parts;
                var chosenParentValue = $('input[name=goal_obj]:checked').val();

                //alert("chosenParentvalue = " + chosenParentValue);

                id_parts = chosenParentValue.split("_");
                main_id_part = id_parts[0];
                goal_name = id_parts[1];

              
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + '&lfid=' + main_id_part + '&goal_name=' + goal_name + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }


            else if (frmMain.create_form_type.value == "QNR") {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + popup_path + '&lfid=' + $("#linked_form_id").val(), 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            } 

            else if (frmMain.create_form_type.value == "DHMPN") {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + popup_path, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
                }
                else
                {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid='+ frmMain.form_staff.value+'&dp='+defaultProg,800,1000,frmMain.create_form_type.value + '' + clientForm.client_id.value);
                }
                
            }            
        }
        else
        {
            if (frmMain.create_form_type.value == "CALL_LOG" || frmMain.create_form_type.value == "DDS")
            {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid='+ frmMain.form_staff.value+'&dp='+defaultProg+'&lfid='+chosenParentValue,800,1400,frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
            else if (frmMain.create_form_type.value == "GOAL_TRACK")
            {            
                var main_id_part = ""
                var goal_track_type = ""
                var id_parts;
                
                //alert("chosenParentvalue = " + chosenParentValue);
                
                id_parts = chosenParentValue.split("_");
                main_id_part = id_parts[0];
                goal_track_type = id_parts[1];
                
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid='+ frmMain.form_staff.value+'&dp='+defaultProg+'&lfid='+main_id_part+'&goal_type='+goal_track_type,800,1000,frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
            else if (frmMain.create_form_type.value == "BTMS")
            {            
                var main_id_part = ""
                var month_year_part = ""
                var id_parts;
                
                //alert("chosenParentvalue = " + chosenParentValue);
                
                id_parts = chosenParentValue.split("_");
                main_id_part = id_parts[0];
                month_year_part = id_parts[1];
                
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid='+ frmMain.form_staff.value+'&dp='+defaultProg+'&lfid='+main_id_part+'&month_year_part='+month_year_part,800,1400,frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
           
            else if (frmMain.create_form_type.value == "PSA_DDS" || frmMain.create_form_type.value == "PSS") {
                var main_id_part = ""
                var goal_name = ""
                var id_parts;

                //alert("chosenParentvalue = " + chosenParentValue);

                id_parts = chosenParentValue.split("_");
                main_id_part = id_parts[0];
                obj_name = $('input[name=parent_form_id]:checked').attr("obj");
                //  alert(obj_name);
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + '&lfid=' + main_id_part + '&obj_name=' + obj_name, 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
           else if (frmMain.create_form_type.value == "BEH_TRACK") {
            popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=' + frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid=' + frmMain.form_staff.value + '&dp=' + defaultProg + '&lfid=' + chosenParentValue, 800, 1100, frmMain.create_form_type.value + '' + clientForm.client_id.value);
           }
            else
            {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft='+ frmMain.create_form_type.value + '&cid=' + clientForm.client_id.value + '&pid=' + frmMain.form_program.value + '&sid='+ frmMain.form_staff.value+'&dp='+defaultProg+'&lfid='+chosenParentValue,800,1000,frmMain.create_form_type.value + '' + clientForm.client_id.value);
            }
            
        }        
    }
    
    function initSet()
    {
        bstartloc = parseInt(document.clientForm.client_id.length/2);
		bdone = 0;
		blastloc = -1;
		bcurrentloc = 0;
		bcurrenttop = 0;
		bthesize = document.clientForm.client_id.length;
		matchfound = 0;
		bcurrentbottom = bthesize;
		bdebug =0
    
        document.clientForm.program.value = '<%=Request.QueryString("pCode")%>';
        document.clientForm.location.value = '<%=Request.QueryString("lCode")%>';
        document.clientForm.client_id.value = '<%=Request.QueryString("cid")%>';

        //frmMain.form_program.value = '<%=Request.QueryString("pid")%>';
        if('<%=user_default_prog%>' != '')
        {
            frmMain.form_program.value = '<%=user_default_prog%>';
            frmMain.program_default.disabled = false;
            frmMain.program_default.checked = true;
        }           
        
        Ajax('staff','scripts/ShowStaffInProgram.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&usid=<%=user_staff_id%>');
 
        frmMain.create_form_type.value = '<%=Request.QueryString("cft")%>';
        
        var radioButtons = document.getElementsByName("sort_by_radio");
	    for(var x=0; x < radioButtons.length; x++)
	    {
	        if( radioButtons[x].value == '<%=Request.QueryString("fb")%>')
	        {
	            radioButtons[x].checked = true;
	        }
	    }
	    	    
	    //add_button = document.getElementById("add_new_form_button");
        //add_button.disabled = false;

    }
    function Person_Search_Filter()
		{
			var find = document.clientForm.person_search_filter.value;
			
			if (find != "")
			{
				init_searchb();
				findstring(find,"b","0");

				if (matchfound == 1)
				{
					if (bdebug)   alert("found match = " + document.clientForm.client_id[bcurrentloc].label + " at location " + bcurrentloc + " id = " + document.clientForm.client_id[bcurrentloc].value);
					document.clientForm.client_id[bcurrentloc].selected = true;
					
				} // end match found
			}
		}
		
		function init_searchb()
		{
			bdone = 0;
			bcurrentloc = bstartloc;
			bcurrenttop = 0;
			bcurrentbottom = bthesize;
			matchfound = 0;
			blastloc = -1;
		}
		
		function findstring(ss,which,exact)
		{   
		    while (bdone == 0)
			searchitb(ss,exact);
		}
		
		function searchitb(ss,exact)
		{  
		    if ((bcurrentloc == 0) || (bcurrentloc == bthesize-1) || (bcurrentloc == blastloc))
			{
			    if (compare(ss,document.clientForm.client_id[bcurrentloc].label,exact));
				{
					bdone=1;
					return;
				}
			}     
			else
			{
			    blastloc = bcurrentloc;
				
				var comp_result = compare(ss,document.clientForm.client_id[bcurrentloc].text,exact)
				
				switch (comp_result)
				{
					case -1:
					bcurrenttop = bcurrentloc;
					bcurrentloc = parseInt((bcurrentbottom + bcurrentloc)/2);
					break;

					case 0:
					while ((bcurrentloc > 0) && (compare(ss,document.clientForm.client_id[bcurrentloc].text,exact) == "0"))
					{ 
						bcurrentloc = bcurrentloc - 1; 
					}
					bcurrentloc = bcurrentloc + 1;
					bdone = 1;
					break;
		   
					case 1:
					bcurrentbottom = bcurrentloc;
					bcurrentloc = parseInt((bcurrentloc + bcurrenttop) /2);
					break; 
				 
					default:
					bdone = 1;
					break;

				} // switch
			} // else
		}
		
		function compare(a1,b1,exact)
		{
		    if (bdebug)
			{
				alert("incompare a1 = " + a1);
				alert("          b1 = " + b1);
			}

			var a = a1.toLowerCase();
			var b = b1.toLowerCase();

			if (b < a)
			{
				if (bdebug) document.write("thinks b < a <BR>");
				return -1;
			}

			if (exact != "1")
			{
				if (b.substring(0,a.length) == a.substring(0,a.length))
				{
					if (bdebug) document.write("thinks b == a <BR>");
						matchfound = 1;
				    
					done = 1;
					return 0;
				}
			}
			else
			{
				if (b.substring(0,b.length) == a.substring(0,a.length))
				{
					if (bdebug) document.write("thinks b == a <BR>");
						matchfound = 1;
					done = 1;
					return 0;
				}
			}
			if (b > a)
			{
				if (bdebug) document.write("thinks b > a <BR>");
					return 1;
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

	function filterList(filterName)
	{
	    window.location.href= "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + filterName + "&sb=" + '<%=Request.QueryString("sb")%>' + "&tf=" + filterName + "&sd=" + clientForm.show_discharged.checked;
	}
	
	function filterTimeFrame(filterName) 
	{

	    window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&tf=" + filterName + "&sd=" + clientForm.show_discharged.checked + '&show_dashboard_flag=1' + '&soft=' + '<%=Request.QueryString("soft")%>';
	}

   	function filterTimeFrame2(filterName) 
	{

	    window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&tf=ALL&sd=" + clientForm.show_discharged.checked + '&show_dashboard_flag=1' + '&soft=' + filterName;
	}
	
	function filterNotes(filterName)
	{
	    window.location.href= "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + filterName + "&tf=" + filterName + "&sd=" + clientForm.show_discharged.checked;
	}

  
	function filterEpisodes(filterName) {
	    window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + filterName + "&tf=" + '<%=Request.QueryString("tf")%>' + "&sd=" + clientForm.show_discharged.checked;
	}

	function filterFormProgram(filterName) {
	    window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + filterName + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=" + '<%=Request.QueryString("tf")%>' + "&sd=" + clientForm.show_discharged.checked;
	}

    function filterClients() 
    {
        Ajax('clients', 'scripts/FilterClientList.asp?pid=<%=Request.QueryString("pCode")%>&cur_client=<%=Request.QueryString("cid")%>&sd=' + clientForm.show_discharged.checked);
	}
	
	function sortList(sortName)
	{
	    window.location.href= "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + sortName + "&tf=" + filterName + "&sd=" + clientForm.show_discharged.checked;
	}
	
	function closeEpisode()
	{
	    conf_text = "Are you sure you would like to close out this person's current episode? Click OK to close the episode, or Cancel to leave it open.";
	    
	    if(confirm(conf_text))
	    {
	        window.location.href= "close_out_episode.asp?cid=" + '<%=Request.QueryString("cid")%>';
	        return true;
	    }
	    else
	    {
	        return false;
	    }   
	}

	function decideShowDiv(formID, sigType) {
	    if (sigType == 'Parent Guardian') {
	        showdiv5(formID, sigType);
	    }
	    else if (sigType == 'Person Served') {
	        showdiv4(formID, sigType);
	    }
	    else
	    //(sigType == 'Provider' || sigType == 'MD' || sigType == 'Supervisor' || sigType == 'Additional' || sigType == 'Family Partner' || sigType == 'Family Partner Supervisor' || sigType == 'MDT 1' || sigType == 'MDT 2' || sigType == 'MDT 3' || sigType == 'MDT 4')
	    {
	        showdiv3(formID, sigType);
	    }

	}
	
	function hidediv2() 
	{ 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow2').style.visibility = 'hidden'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow2.visibility = 'hidden'; 
            } 
            else 
            { // IE 4 
                document.all.hideshow2.style.visibility = 'hidden'; 
            } 
        } 
    }
function hidediv2b() 
	{ 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow2b').style.visibility = 'hidden'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow2b.visibility = 'hidden'; 
            } 
            else 
            { // IE 4 
                document.all.hideshow2b.style.visibility = 'hidden'; 
            } 
        } 
    }
	function showdiv2() 
    { 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow2').style.visibility = 'visible'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow2.visibility = 'visible'; 
            } 
            else
            { // IE 4 
                document.all.hideshow2.style.visibility = 'visible'; 
            } 
        } 
    }
function showdiv2b() 
    { 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow2b').style.visibility = 'visible'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow2b.visibility = 'visible'; 
            } 
            else
            { // IE 4 
                document.all.hideshow2b.style.visibility = 'visible'; 
            } 
        } 
    }
    //PROVIDER/SUPERVISOR/MD SIGN SECTION
    function hidediv3() 
    { 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow3').style.visibility = 'hidden'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow3.visibility = 'hidden';   
            } 
            else 
            { // IE 4 
                document.all.hideshow3.style.visibility = 'hidden'; 
            } 
        } 
    }
    function showdiv3(fid,st) 
    { 
        document.signForm.uid.value = fid;
        document.signForm.st.value = st;
        document.signForm.cur_cid.value = '<%=Request.QueryString("cid")%>';
        document.signForm.cur_sb.value = '<%=Request.QueryString("sb")%>';
        document.signForm.cur_fb.value = '<%=Request.QueryString("fb")%>';
         
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow3').style.visibility = 'visible'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow3.visibility = 'visible'; 
            } 
            else 
            { // IE 4 
                document.all.hideshow3.style.visibility = 'visible'; 
            } 
        } 
    }
    
    //CLIENT SIGN SECTION
    function hidediv4() 
    { 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow4').style.visibility = 'hidden'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow4.visibility = 'hidden';   
            } 
            else 
            { // IE 4 
                document.all.hideshow4.style.visibility = 'hidden'; 
            } 
        } 
    }
    function showdiv4(fid,st) 
    { 
        document.signFormClient.uid.value = fid;
        document.signFormClient.st.value = st;
        document.signFormClient.cur_cid.value = '<%=Request.QueryString("cid")%>';
        document.signFormClient.cur_sb.value = '<%=Request.QueryString("sb")%>';
        document.signFormClient.cur_fb.value = '<%=Request.QueryString("fb")%>';
          
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow4').style.visibility = 'visible'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow4.visibility = 'visible'; 
            } 
            else 
            { // IE 4 
                document.all.hideshow4.style.visibility = 'visible'; 
            } 
        } 
    }  
    
    //PARENT SIGN SECTION
    function hidediv5() 
    { 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow5').style.visibility = 'hidden'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow5.visibility = 'hidden';   
            } 
            else 
            { // IE 4 
                document.all.hideshow5.style.visibility = 'hidden'; 
            } 
        } 
    }
    function showdiv5(fid,st) 
    { 
        document.signFormParent.uid.value = fid;
        document.signFormParent.st.value = st;
        document.signFormParent.cur_cid.value = '<%=Request.QueryString("cid")%>';
        document.signFormParent.cur_sb.value = '<%=Request.QueryString("sb")%>';
        document.signFormParent.cur_fb.value = '<%=Request.QueryString("fb")%>';
          
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshow5').style.visibility = 'visible'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshow5.visibility = 'visible'; 
            } 
            else 
            { // IE 4 
                document.all.hideshow5.style.visibility = 'visible'; 
            } 
        } 
    }
function hidedivISP() 
    { 
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshowISP').style.display = 'none'; 
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshowISP.display = 'none';   
            } 
            else 
            { // IE 4 
                document.all.hideshowISP.style.display = 'none'; 
            } 
        } 
    }
    function showdivISP() 
    {       
        if (document.getElementById) 
        { // DOM3 = IE5, NS6 
            document.getElementById('hideshowISP').style.display = 'block';
        } 
        else 
        { 
            if (document.layers) 
            { // Netscape 4 
                document.hideshowISP.display = 'block'; 
            } 
            else 
            { // IE 4 
                document.all.hideshowISP.style.display = 'block';
            } 
        } 
    }
function decideShowISPDiv() {
	    
	        showdivISP();

	}
</script>

<!-- Script by hscripts.com -->
<script language=javascript>
            var text1="Waiting for the form to be Finalized (i.e. signed) ";
            var text2="Person signature not available on this form";                    
            var text3="This form is not locked and can be edited";
            var text4="This form has been locked and is ready for signatures";  
            var text5="Provider signature not available on this form";  
            var text6="Guardian signature not available on this form";  
            var text7="MD signature not available on this form";  
            var text8="Supervisor signature not available on this form";  
            var text9="Waiting for proper user to sign this signature";
            var text10="Open the form with locked data for view only";
            var text11="Open the form with unlocked data for editing";
            var text12="View the history of this form";
            var text13="View all notes on this form";
            var text14="View all attachments on this form";
            var text15="Create a new instance of this form";
//This is the text to be displayed on the tooltip.

if(document.images)
{
  pic1= new Image(); 
  pic1.src='htooltip/bubble_top.gif'; 
  pic2= new Image();
  pic2.src='htooltip/bubble_middle.gif'; 
  pic3= new Image(); 
  pic3.src='htooltip/bubble_bottom.gif'; 
}

function showToolTip(e,text)
{
    if(document.all)e = event;
    var obj = document.getElementById('bubble_tooltip');
    var obj2 = document.getElementById('bubble_tooltip_content');
		
	obj2.innerHTML = text;
    obj.style.display = 'block';
    var st = Math.max(document.body.scrollTop,document.documentElement.scrollTop);
				
    if(navigator.userAgent.toLowerCase().indexOf('safari')>=0)st=0; 
    var leftPos = e.clientX-2;
				
    if(leftPos<0)leftPos = 0;
    obj.style.left = leftPos + 'px';
    obj.style.top = e.clientY-obj.offsetHeight+2+st+ 'px';
}       
        
        function hideToolTip()
        {
                document.getElementById('bubble_tooltip').style.display = 'none';
        }

        function editClient()
        {
	            if("" == document.clientForm.client_id.value)
	            {
	                alert('Please select a client to edit');
	            }
	            else
	            {
	                popupwindow('edit_client.asp?from_home=1&fromModule=1&cid=' + document.clientForm.client_id.value,600,1300,'editClientWindow');
	            }
	    }
</script>
<style type="text/css">
        #bubble_tooltip{
				width:210px;
                position:absolute;
                display: none;
        }
        #bubble_tooltip .bubble_top{
                position:relative;
                background-image: url(htooltip/bubble_top.gif);
                background-repeat:no-repeat;
                height:18px;
                }
        #bubble_tooltip .bubble_middle{
                position:relative;
                background-image: url(htooltip/bubble_middle.gif);
                background-repeat: repeat-y;
                background-position: bottom left;
        }
        #bubble_tooltip .bubble_middle div{
		        padding-left: 12px;
                padding-right: 20px;
                position:relative;
                font-size: 11px;
                font-family: arial, verdana, san-serif;
                text-decoration: none;
                color: blue;	
		        text-align:justify;	
        }
        #bubble_tooltip .bubble_bottom{
                background-image: url(htooltip/bubble_bottom.gif);
                background-repeat:no-repeat;
                height:65px;
                position:relative;
                top: 0px;
        }
    
/* Script by hscripts.com */
#hideshow {
	position: absolute;
	width: 100%;
	height: 100%;
	top: 0;
	left: 0;
}
.popup_block {
	background: #ddd;
	padding: 10px 20px;
	border: 10px solid #fff;
	float: left;
	width: 400px;
	position: fixed;
	top: 20%;
	left: 50%;
	margin: 0 0 0 -250px;
	z-index: 100;
	font: 9pt verdana,arial,helvetica,sans-serif;
}
.popup_block .popup {
	float: left;
	width: 100%;
	background: #fff;
	margin: 10px 0;
	padding: 10px 0;
	border: 1px solid #bbb;
}
.popup_block2 {
	background: #ddd;
	padding: 10px 20px;
	border: 10px solid #fff;
	float: left;
	width: 1000px;
	position: fixed;
	top: 20%;
	left: 30%;
	margin: 0 0 0 -250px;
	z-index: 100;
	font: 9pt verdana,arial,helvetica,sans-serif;
}
.popup_block2 .popup {
	float: left;
	width: 100%;
	background: #fff;
	margin: 10px 0;
	padding: 10px 0;
	border: 1px solid #bbb;
}
.popup h3 {
	font-size: 17px;
	font-family:  verdana, sans-serif;
	margin : 0 0 12px; padding : 7px 0 3px;
	color : #000080;
	border-bottom : 1px solid #000080;
	letter-spacing : -1px;
	text-align: left;
}
.popup p {
	padding: 5px 10px;
	margin: 5px 0;
}
.popup img.cntrl {
	position: absolute;
	right: -20px;
	top: -20px;
}
#fade {
	background: #000;
	position: fixed;
	width: 100%;
	height: 100%;
	filter:alpha(opacity=80);
	opacity: .80;
	-ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=80)"; /*--IE 8 Transparency--*/
	left: 0;
	top: 0;
	z-index: 10;
}
/*--Making IE6 Understand Fixed Positioning--*/


*html #fade {
	position: absolute;
	top:expression(eval(document.compatMode &&
	document.compatMode=='CSS1Compat') ?
	documentElement.scrollTop : document.body.scrollTop);
}

*html .popup_block {
	position: absolute;
	top:expression(eval(document.compatMode &&
	document.compatMode=='CSS1Compat') ?
	documentElement.scrollTop
	+((documentElement.clientHeight-this.clientHeight)/2)
	: document.body.scrollTop
	+((document.body.clientHeight-this.clientHeight)/2));
	
	left:expression(eval(document.compatMode &&
	document.compatMode=='CSS1Compat') ?
	documentElement.scrollLeft 
	+ (document.body.clientWidth /2 ) 
	: document.body.scrollLeft 
	+ (document.body.offsetWidth /2 ));
}

/*--IE 6 PNG Fix--*/

/*img{ behavior: url(iepngfix.htc) }*/

</style>

</head>
<body onload="initSet();">
    <script type="text/javascript">



 $(document).ready(function() {







if ('<%=hcp_upload %>' != "-1")
{

window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>';
popupwindow("upload_hcp_attachment.asp?uid=" + '<%=hcp_upload %>'  + "&cid=" + '<%=Request.QueryString("cid")%>',800,1400,'<%=hcp_upload %>');


};

$("input[name=record_view]").change(function () {

if($("input[name=record_view]").is(':checked')) {
    window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=ALL&sd=" + clientForm.show_discharged.checked + "&ufid=" + '<%=Request.QueryString("ufid")%>' + "&ulfid=" + '<%=Request.QueryString("ulfid")%>' + "&uft=" + '<%=Request.QueryString("uft")%>' + '&show_dashboard_flag=' + '<%=Request.QueryString("show_dashboard_flag")%>' + '&soft=' + '<%=Request.QueryString("soft")%>';
} else {
   window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=&sd=" + clientForm.show_discharged.checked + "&ufid=" + '<%=Request.QueryString("ufid")%>' + "&ulfid=" + '<%=Request.QueryString("ulfid")%>' + "&uft=" + '<%=Request.QueryString("uft")%>' + '&show_dashboard_flag=' + '<%=Request.QueryString("show_dashboard_flag")%>' + '&soft=';

}

});


$(".goToDoctor").click(function () {

$("#recommendations").show();
$("#proceed_button").show();
$("#proceed_span").show();


$("#dialogDiv").dialog({
          
            width: 700,
            title: "APPOINTMENT INFORMATION",
            autoOpen: false,
            modal : true,
           open: function( event, ui ) {
    
         
        $("select[name=form_program3]").change(function() {
      
         $.ajax({url: 'scripts/ShowStaffInProgram2.asp?cid=<%=Request.QueryString("cid")%>&pid='+$("select[name=form_program3]").val()+'&usid=<%=user_staff_id%>', cache: false, success: function(result){

            $("select[name=form_staff3]").html(result);

            }});

        

          });

          $("#proceed_button").click(function() {


         var theCheckboxes = $("input[name=follow_up]"); 
         if (theCheckboxes.filter(":checked").length > 1) {
        $(this).removeAttr("checked");
        alert( "Please selected one folow-up at a time." );

        return false;
        } else {
      //  alert($("input[name=follow_up]:checked").val());
      //  alert($("input[name=follow_up]:checked").attr("unique_form_id"));


           if($("select[name=form_program3]").val() == "") {
                              
             alert("Please select a Responsible Program");

               return false;

            } else if ($("select[name=form_staff3]").val() == "") {
             
             alert("Please select a Responsible Staff");

            return false;

          
            
              
            } else {
               

            if ($("input[name=follow_up]").is(':checked')) {

            
         
              popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=HCEF&cid=' + clientForm.client_id.value + '&pid=' + $("select[name=form_program3]").val() + '&sid=' + $("select[name=form_staff3]").val() + '&lfid=' + $("input[name=follow_up]:checked").attr("unique_form_id") + '&rec=' + $("input[name=follow_up]:checked").val(), 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);

           


              } else {
        
               alert("Please select a Follow-Up option");

                return false;

             }



            }
        }
         });
        
         

          

        },
           close: function(event, ui) { 

    //  $("#dialogDiv").html('');

         },

            buttons : [
               
                {
                    text : "Start New Appointment       ",
                          click : function() {
                       // $(this).dialog('close');
                   
                       

                        if($("select[name=form_program3]").val() == "") {
                              
             alert("Please select a Responsible Program");

               return false;

            } else if ($("select[name=form_staff3]").val() == "") {
             
             alert("Please select a Responsible Staff");

            return false;

          
            
              
            } else {
               

            popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=HCEF&cid=' + clientForm.client_id.value + '&pid=' + $("select[name=form_program3]").val() + '&sid=' + $("select[name=form_staff3]").val() + '&dp=', 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);



            }



                        
                    }
                } ]
        });
 $("#dialogDiv").dialog("open");
  $("#hcef_complete").hide();



   });

  


$("#sched_appt").click(function () {


var js_uid = $(this).attr("unique_form_id");




$("#dialogDiv2").dialog({
          
            width: 400,
            title: "SCHEDULE APPOINTMENT",
            autoOpen: false,
            modal : true,
           open: function( event, ui ) {
    
         

$("#default_focus").hide();
         

//  $('#divTimePicker').timepicker();  
//$('#divDatePicker').datepicker();  



$('#divTimePicker').blur(function(){
    var validTime = $(this).val().match(/^(0?[1-9]|1[012])(:[0-5]\d)[APap][mM]$/);
    if (!validTime) {
        $(this).val('');
        alert("Time must be entered in HH:MM[AM/PM] format (example: 11:15AM)");
   
       return false;
    } else {
        return true;
    }
});



$("#divDatePicker").datepicker({
    onSelect: function() {
        $("#dialogDiv2").focus();
    }
});
         


        },
           close: function(event, ui) { 

    //  $("#dialogDiv").html('');

         },

            buttons : [
               
                {
                    text : "SUBMIT",
                          click : function() {
                   var appt_date = $('#divDatePicker').val() 
                    var comp = appt_date.split('/');
                    var m = parseInt(comp[0], 10);
                    var d = parseInt(comp[1], 10);
                    var y = parseInt(comp[2], 10);
                    var date = new Date(y,m-1,d);
                    if (date.getFullYear() == y && date.getMonth() + 1 == m && date.getDate() == d) {
                       window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + '&action=schedule&uid=' + js_uid + '&date=' + $('#divDatePicker').val() + '&time=' + $('#divTimePicker').val();
                    } else {
                      alert('Please enter date in MM/DD/YYYY format');

                    }

                 
                    
                    }
                } ]
        });

 $("#dialogDiv2").dialog("open");






});


$(".backToDoctor").click(function () {

$("#recommendations").hide();
$("#proceed_button").hide();
$("#proceed_span").hide();

$("#dialogDiv").dialog({
          
            width: 700,
            title: "APPOINTMENT INFORMATION",
            autoOpen: false,
            modal : true,
           open: function( event, ui ) {
    
         
        $("select[name=form_program3]").change(function() {
      
         $.ajax({url: 'scripts/ShowStaffInProgram2.asp?cid=<%=Request.QueryString("cid")%>&pid='+$("select[name=form_program3]").val()+'&usid=<%=user_staff_id%>',cache: false, success: function(result){

            $("select[name=form_staff3]").html(result);

            }});

           $.ajax({url: 'scripts/ShowPossibleParentsForForm.asp?cid=<%=Request.QueryString("cid")%>&pid='+$("select[name=form_program3]").val()+'&cft=HCEF',cache: false, success: function(result){

            $("#hcef_complete").html(result);

            }});

          });
          

        },
           close: function(event, ui) { 

    //  $("#dialogDiv").html('');

         },

           buttons : [
               
                {
                    text : "Finish Appointment",
                          click : function() {
                   if ($("select[name=form_program3]").val() == "") {
                        return false;
                    } else {
                      popupwindow('http://<%=url_org_name%>/upload_hcp_attachment.asp?uid=' + $("select[name=hcp_form]").val() + '&cid=' + clientForm.client_id.value + '&pid=' + $("select[name=form_program3]").val() + '&sid=' + $("select[name=form_staff3]").val() + '&dp=', 800, 1000, frmMain.create_form_type.value + '' + clientForm.client_id.value);
                  }
                    }
                } ]
        });


 $("#dialogDiv").dialog("open");
$("#hcef_complete").show();


   });

    
 $("#form_group").change(function () {

   window.location.href = "index.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + '<%=Request.QueryString("episodes")%>' + "&tf=ALL" +  "&sd=" + clientForm.show_discharged.checked + "&soft=" + $(this).val() + '&show_dashboard_flag=' + '<%=Request.QueryString("show_dashboard_flag")%>';

});
  


    });

	    
</script>


<div id="hideshowISP" style="display:none;"> 
  <div id="fade"></div>
  <div class="popup_block2"> 
    <div class="popup" align="center"> <a href="javascript:hidedivISP()"><img src="icon_close.png" border="0" width="28" height="31" class="cntrl" /></a> 
      
        <%if client_hcsis_id <> "" then %>
         <table width="100%" cellpadding="4" cellspacing="0" border="0" id="box" style="border:1px solid;">
          <tr> 
            <td colspan="9"><h3>HCSIS ISP Assessment Status</h3></td>
          </tr>
          <tr> 
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Name</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Status</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Requested<br />By/Date</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Due Date</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Started<br />By/Date</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Submitted for DDS<br /> ReviewBy/Date</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Approved<br />By/Date</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>DDS Review Started<br />By/Date</b></td>
                <td bgcolor="#e0e0e0" align="left" style="border:1px solid;"><b>Last Updated<br />By/Date</b></td>
           </tr>
          <% 
              is_row_count = 0

            Set SQLStmtISPStatus = Server.CreateObject("ADODB.Command")
  	        Set rsISPStatus = Server.CreateObject("ADODB.Recordset")
  	        SQLStmtISPStatus.CommandText = "exec get_isp_assessment_status_table '" & client_hcsis_id & "'"
  	        SQLStmtISPStatus.CommandType = 1
  	        Set SQLStmtISPStatus.ActiveConnection = conn
  	        SQLStmtISPStatus.CommandTimeout = 45 'Timeout per Command
  	     '  response.write "SQL = " & SQLStmtISPStatus.CommandText
  	        rsISPStatus.Open SQLStmtISPStatus 
              
              Do Until rsISPStatus.EOF

                cur_assess_has_issue = rsISPStatus("isp_assessments_status")
              %>
          <tr <%if is_row_count mod 2 = 0 THEN %>style="background-color:#FFFBD2;" <% else %>style="background-color:#FFFFFF;"<% end if %>> 
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("AssessmentName")%></td>
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("AssessmentStatus")%></td>
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("requested_by")%><br /><%=rsISPStatus("requested_on")%></td>
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("due_date")%></td>
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("started_by")%><br /><%=rsISPStatus("started_on")%></td>
              <% if cur_assess_has_issue = "Red" THEN %>
              <td class="blueFontSmall" align="left" style="border:1px solid #000000; background-color:#FF0000; color:#FFFFFF;"><%=rsISPStatus("submitted_to_dds_by")%><br /><%=rsISPStatus("submitted_to_dds_on")%></td>
              <% elseif cur_assess_has_issue = "Orange" THEN %>
              <td class="blueFontSmall" align="left" style="border:1px solid #000000; background-color:#FF6633; color:#FFFFFF;"><%=rsISPStatus("submitted_to_dds_by")%><br /><%=rsISPStatus("submitted_to_dds_on")%></td>
              <% elseif cur_assess_has_issue = "Yellow" THEN %>
              <td class="blueFontSmall" align="left" style="border:1px solid #000000; background-color:#993300; color:#FFFFFF;"><%=rsISPStatus("submitted_to_dds_by")%><br /><%=rsISPStatus("submitted_to_dds_on")%></td>
              <% else %>
              <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("submitted_to_dds_by")%><br /><%=rsISPStatus("submitted_to_dds_on")%></td>
              <% end if  %>
                
                <!--<td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("submitted_to_dds_by")%><br /><%=rsISPStatus("submitted_to_dds_on")%></td>-->
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("approved_by")%><br /><%=rsISPStatus("approved_on")%></td>
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("dds_review_started_by")%><br /><%=rsISPStatus("dds_review_started_on")%></td>
                <td class="blueFontSmall" align="left" style="border:1px solid #000000;"><%=rsISPStatus("LastUpdatedBy")%><br /><%=rsISPStatus("LastUpdatedOn")%></td>
           </tr>
              <%
                rsISPStatus.MoveNext
                is_row_count = is_row_count + 1
              Loop   
          %>
        </table>
      <%end if %>
     </div>
  </div>
</div>

<div id="hideshow2" style="visibility:hidden;"> 
  <div id="fade"></div>
  <div class="popup_block"> 
    <div class="popup" align="center"> <a href="javascript:hidediv2()"><img src="icon_close.png" border="0" width="28" height="31" class="cntrl" /></a> 
      <img src="images/loginHeader.gif" width="372" height="32" border="0" alt="MSDP: Electronic Healthcare Forms" title="MSDP: Electronic Healthcare Forms" /> 
      <form action="process_uploads.asp" name="uploadForm" method="POST" enctype="multipart/form-data">
	  <h3>Upload Forms</h3>
	  <p align="left">Please select a form to upload:</p>
	  <p align="left">Form:&nbsp;&nbsp; <input type="file" name="target_file" size="40" /></p>
	  <p align="center"><input type="submit" value="Submit" class="submit" /></p>
      </form>
    </div>
  </div>
</div>
    <div id="hideshow2b" style="visibility:hidden;"> 
  <div id="fade"></div>
  <div class="popup_block"> 
    <div class="popup" align="center"> <a href="javascript:hidediv2b()"><img src="icon_close.png" border="0" width="28" height="31" class="cntrl" /></a> 
      <img src="images/loginHeader.gif" width="372" height="32" border="0" alt="MSDP: Electronic Healthcare Forms" title="MSDP: Electronic Healthcare Forms" /> 
      <form action="process_hcsis_uploads.asp" name="uploadForm" method="POST" enctype="multipart/form-data">
	  <h3>Upload Forms</h3>
      <table>
        <tr>
            <td>Please select an extract file to upload:</td>
            <td><input type="file" name="target_file" size="40" /></td>
        </tr>
        <tr> 
            <td align="right" nowrap>Extract File Type:&nbsp;&nbsp;</td>
            <td align="left"> <select name="hcsis_extract_type">
                <option value="HCR">Healthcare Record</option>
                <option value="Incidents">Incidents</option>
                <option value="ISP">ISP</option>
                <option value="MORs">MORs</option>
                <option value="Restraints">Restraints</option>
              </select>
            </td>
        </tr>
        <tr>
            <td colspan="2"><input type="submit" value="Submit" class="submit" /></td>
        </tr>          
      </table>
      </form>
    </div>
  </div>
</div>
<div id="hideshow3" style="visibility:hidden;"> 
  <div id="Div1"></div>
  <div class="popup_block"> 
    <div class="popup" align="center"> <a href="javascript:hidediv3()"><img src="icon_close.png" border="0" width="28" height="31" class="cntrl" /></a> 
      <img src="images/loginHeader.gif" width="372" height="32" border="0" alt="MSDP: Electronic Healthcare Forms" title="MSDP: Electronic Healthcare Forms" /> 
      <form action="sign_form.asp" name="signForm" method="POST">
        <input type="hidden" name="page" value="sign_form_check" />
        <input type="hidden" name="st" value="" />
        <input type="hidden" name="uid" value="" />
        <input type="hidden" name="cur_cid" value="" />
        <input type="hidden" name="cur_sb" value="" />
        <input type="hidden" name="cur_fb" value="" />
        <table width="100%">
          <tr> 
            <td ><h3>Sign Form</h3>
              <br/></td>
          </tr>
          <tr> 
            <td align="left">Please enter your Password to Sign: &nbsp;
              <input type="password" name="pass_it" size="" /></td>
          </tr>
          <tr> 
            <td ><br />
              <input type="submit" value="Submit" class="submit" /></td>
          </tr>
        </table>
      </form>
    </div>
  </div>
</div>
<script>
function updatePhrase()
{ 
    if(signFormClient.client_sig_type.value == 'Pass Phrase')
    {
        signFormClient.client_pass_phrase.disabled = false;
        if(document.getElementById("phrase_hint"))
        {
            signFormClient.phrase_hint.disabled = false;
        }
    }
    else
    {
        signFormClient.client_pass_phrase.disabled = true;
        if(document.getElementById("phrase_hint"))
        {
            signFormClient.phrase_hint.disabled = true;
        }
    }
}
function checkClientSign()
{
    //alert("checking client sig");
    if(signFormClient.client_sig_type.value == 'Pass Phrase' && (signFormClient.client_pass_phrase.value == '' || signFormClient.client_pass_phrase.value == ' '))
    {
        alert("Please enter a pass phrase if choosing 'Pass Phrase' as the signature type.");
        return false;
    }
}


</script>
<div id="hideshow4" style="visibility:hidden;"> 
  <div id="Div2"></div>
  <div class="popup_block"> 
    <div class="popup" align="center"> <a href="javascript:hidediv4()"><img src="icon_close.png" border="0" width="28" height="31" class="cntrl" /></a> 
      <img src="images/loginHeader.gif" width="372" height="32" border="0" alt="MSDP: Electronic Healthcare Forms" title="MSDP: Electronic Healthcare Forms" /> 
      <form action="sign_form_client.asp" name="signFormClient" method="POST">
        <input type="hidden" name="page" value="sign_form_check" />
        <input type="hidden" name="st" value="" />
        <input type="hidden" name="uid" value="" />
        <input type="hidden" name="cur_cid" value="" />
        <input type="hidden" name="cur_sb" value="" />
        <input type="hidden" name="cur_fb" value="" />
        <table width="100%">
          <tr> 
            <td align="center" colspan="2"><h3>Sign Form</h3>
              <br/></td>
          </tr>
          <tr> 
            <td align="right" nowrap>Signature Type:&nbsp;&nbsp;</td>
            <td align="left"> <select name="client_sig_type" onchange="updatePhrase();">
                <option value="Not Appropriate">Not Appropriate</option>
                <option value="Attached">Attached</option>
                <option value="Filed">Filed</option>
                <option value="Pass Phrase">Pass Phrase</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp;</td>
          </tr>
          <%
          if Request.QueryString("cid") <> "" THEN
            'CHECK FOR EXISTING PASS PHRASE
            Set SQLStmtP = Server.CreateObject("ADODB.Command")
  	        Set rsP = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtP.CommandText = "select case when EXISTS(select * from client_master where client_id = " & Request.QueryString("cid") & " and Pass_Phrase != '') THEN 'Yes' ELSE 'No' END as has_phrase"
  	        SQLStmtP.CommandType = 1
  	        Set SQLStmtP.ActiveConnection = conn
  	        SQLStmtP.CommandTimeout = 45 'Timeout per Command
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rsP.Open SQLStmtP  
  	        
  	        if rsP("has_phrase") = "Yes" THEN
  	            Set SQLStmtP2 = Server.CreateObject("ADODB.Command")
  	            Set rsP2 = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmtP2.CommandText = "select Phrase_Hint from client_master where client_id = " & Request.QueryString("cid")
  	            SQLStmtP2.CommandType = 1
  	            Set SQLStmtP2.ActiveConnection = conn
  	            SQLStmtP2.CommandTimeout = 45 'Timeout per Command
  	            'response.write "SQL = " & SQLStmt2.CommandText
  	            rsP2.Open SQLStmtP2
  	            
  	            cur_sign_hint = rsP2("Phrase_Hint")
  	      %>
  	      <tr>
  	        <td align="right">Password Hint:&nbsp;&nbsp;</td>
  	        <td align="left"><%=cur_sign_hint%>?</td>
  	      </tr>
  	      <%
  	        else
  	      %>
  	      <tr>
  	        <td align="right" nowrap>Password Hint:&nbsp;&nbsp;</td>
  	        <td align="left">
  	        <select name="phrase_hint" id="phrase_hint" disabled>
  	            <option value="What is your mother's maiden name">What is your mother's maiden name</option>
  	            <option value="What was your first pets name">What was your first pets name</option>
  	            <option value="What was your highschool mascot">What was your highschool mascot</option>
  	            <option value="What model was your very first car">What model was your very first car</option>
  	        </select></td>
  	      </tr>
  	      <%  
  	        end if        
          %>
          <tr>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td align="right">Pass Phrase:&nbsp;&nbsp;</td>
            <td align="left"> <input type="password" name="client_pass_phrase" id="client_pass_phrase" size="40" disabled/> 
            </td>
          </tr>
          <tr>
            <td colspan="2">&nbsp;&nbsp;</td>
          </tr>
          <%
            if rsP("has_phrase") <> "Yes" THEN
          %>
          <tr>
            <td align="center" colspan="2"><font color="red">*No password hint and pass phrase exists yet for this client in the system. The values entered will be saved for future validation.</font></td>
          </tr>
          <% 
            end if
          end if
          %>
          <tr> 
            <td align="middle" colspan="2"><br />
			  <input type="submit" value="Submit" class="submit" onclick="return checkClientSign();" /></td>
          </tr>
        </table>
      </form>
    </div>
  </div>
</div>
 <div id="dialogDiv" style="display:none;">
        <table>
             <tr><td><b>Responsible Program:</b>&nbsp;&nbsp;&nbsp;&nbsp;</td><td><b>Responsible Staff:</b></td></tr>
            <tr><td>
         <select name="form_program3" <% if Request.QueryString("cid") = "" THEN%>disabled<%end if%>>
                <option value="">Select a Program</option>
                    <%
                    client_active = 0
                    
                        if Request.QueryString("cid") <> "" THEN
                        
	                        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                        SQLStmt2.CommandText = "exec get_programs_for_staff_and_client '" & Session("user_name") & "'," & Request.QueryString("cid")  
  	                        SQLStmt2.CommandType = 1
  	                        Set SQLStmt2.ActiveConnection = conn
  	                        SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                        'response.write "SQL = " & SQLStmt2.CommandText
  	                        rs2.Open SQLStmt2
  	                        Do Until rs2.EOF
  	                        
  	                            client_active = 1
  	                            pif_age_require = rs2("Require_New_PIF_For_CA")
  	                            pif_max_age_hours = rs2("New_PIF_for_CA_max_hours")
	                %>
	                        <option value="<%=rs2("Program_ID")%>" pif_age_require="<%=pif_age_require%>" pif_max_age_hours="<%=pif_max_age_hours%>" ><%Response.write rs2("Program_Name")%></option>
	                <%
	                        rs2.MoveNext
                            Loop
	                    end if
	                %>
            </select>
            </td><td>
              <div id="staff3">
                    <select name="form_staff3">
                        <option value="">Choose a Program</option>
                    </select>
                </div>
                </td></tr>
                <tr><td colspan="2">
                     <div id="hcef_complete"></div>
                <div id="recommendations" style="display:none;padding-top:10px;">
                    <style>
                        .table_boarder {
                            border: 1px solid black;
                           border-collapse: collapse;
                           text-align: left;
                        }
                       
                        .th_color {
                            background-color: #f1f1c1;
                            padding: 5px;
                        }
                        .td_padding {
                          
                            padding: 5px;
                        }
                    </style>
                  
                       <%

                             recommendations_count = 0
                              recommendations_string = ""
                              

                     if Request.QueryString("cid") <> "" then

                                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtV.CommandText = "get_hcp_encounter_form_recommendations " & Request.QueryString("cid")
  	                            SQLStmtV.CommandType = 1
  	                            Set SQLStmtV.ActiveConnection = conn
  	                            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                            rsV.Open SQLStmtV

                            '  recommendations_count = 0
                            '  recommendations_string = ""
                              

                                  
                                  
  	                            Do Until rsV.EOF

                               cme_prescribing_md = rsV("cme_prescribing_md")
                                      cme_prescribing_minus = instr(rsV("cme_prescribing_md"),"_")
                                      cme_prescribing_total = Len(cme_prescribing_md)
                                      cme_prescribing_calc = (cme_prescribing_total - cme_prescribing_minus)
                                      recommendation = rsV("recommendation") 
                                     recommendation_date = rsV("recommendation_date")
                                     unique_form_id = rsV("unique_form_id")

                            
                             if rsV("follow_up_started")  = "0" then

                          

                                    recommendations_count = recommendations_count + 1
        
                                     recommendations_string = recommendations_string & "<tr class=""table_boarder"">"
                                     recommendations_string = recommendations_string & " <td class=""table_boarder td_padding"">" & rsV("appt_date") & "</td>"
                           

                                    
                                      
                                     recommendations_string = recommendations_string & " <td class=""table_boarder td_padding"">" & replace(Right(cme_prescribing_md,cme_prescribing_calc),"_"," ") & "</td>"
                                     recommendations_string = recommendations_string & " <td class=""table_boarder td_padding"">" & recommendation & "</td>"
                                     recommendations_string = recommendations_string & " <td class=""table_boarder td_padding"">" & recommendation_date & "</td>"
                                   '  recommendations_string = recommendations_string & " <td style=""border: 1px solid black;border-collapse: collapse;text-align: center;""><input type=""checkbox"" name=""follow_up"" unique_form_id=""" & rsV("unique_form_id") & """ value='" & Replace(cme_prescribing_md," ","^") & "'></td>"
                                     recommendations_string = recommendations_string & " <td style=""border: 1px solid black;border-collapse: collapse;text-align: center;""><input type=""checkbox"" name=""follow_up"" unique_form_id=""" & rsV("unique_form_id") & """ value='" & Replace(recommendation," ","^") & "'></td>"
                                     recommendations_string = recommendations_string & " </tr>"
                          end if
                              
                                rsV.MoveNext
                                Loop
                              
                             ' response.write(recommendations_string)
                     end if
                               %>


             

                <%if recommendations_count > 0 then %>
                  <table class="table_boarder" >
             
                  <tr class="table_boarder">
                    <th class="table_boarder th_color">DOA</th>
                    <th class="table_boarder th_color">Doctor</th>
                    <th class="table_boarder th_color">Recommendation&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
                   <th class="table_boarder th_color">Due Date</th>
                     <th class="table_boarder th_color">Select</th>
                  </tr>
                     <%response.write(recommendations_string) %>
              
                </table>
                 <%end if %>                

                </div>
                    
             
                </td></tr>
                
            </table>
            <%if recommendations_count > 0 then %>
           <div style="float:right;">
           <input class="ui-button ui-widget ui-corner-all" id="proceed_button" type="submit" value="Proceed With Follow-Up Appointment"><div style="padding-top:10px;"><span id="proceed_span" style="padding-left:180px;"><b>OR</b></span></div></div>
           <%end if %>
    </div>
    <div id="dialogDiv2" style="display:none;">
        <input id="default_focus" type="text"  />
      <div style="padding-left:3px;"><b>Date:</b><input id="divDatePicker" type="text"  /></div><div><b>Time:</b><input id="divTimePicker" type="text"  /></div>
        
     </div>
<div id="hideshow5" style="visibility:hidden;"> 
  <div id="Div3"></div>
  <div class="popup_block"> 
    <div class="popup" align="center"> <a href="javascript:hidediv5()"><img src="icon_close.png" border="0" width="28" height="31" class="cntrl" /></a> 
      <img src="images/loginHeader.gif" width="372" height="32" border="0" alt="MSDP: Electronic Healthcare Forms" title="MSDP: Electronic Healthcare Forms" /> 
      <form action="sign_form_parent.asp" name="signFormParent" method="POST">
        <input type="hidden" name="page" value="sign_form_check" />
        <input type="hidden" name="st" value="" />
        <input type="hidden" name="uid" value="" />
        <input type="hidden" name="cur_cid" value="" />
        <input type="hidden" name="cur_sb" value="" />
        <input type="hidden" name="cur_fb" value="" />
        <table width="100%">
          <tr> 
            <td colspan="2"><h3>Sign Form</h3>
              <br /></td>
          </tr>
          <tr> 
            <td align="right">Signature Type:</td>
            <td align="left"> <select name="parent_sig_type">
                <option value="Not Appropriate">Not Appropriate</option>
                <option value="Attached">Attached</option>
                <option value="Filed">Filed</option>
              </select> </td>
          </tr>
          <tr> 
            <td align="middle" colspan="2"><br />
			  <input type="submit" value="Submit" class="submit" /></td>
          </tr>
        </table>
      </form>
    </div>
  </div>
</div>
<!-- ###### BEGIN HEADER ###### -->
<!--#include file="includes/header_index.asp" -->
<!-- ###### END HEADER ###### -->
   <div id="container">
   <form name="clientForm" action="">
	<table cellpadding="0" cellspacing="0" border="0" width="100%">
	  <tr >	  
	  <td align="left" valign="bottom"><font style="font-weight:bold; font-size:22px; color:#dc7d18;">&nbsp;CLIENTS</font>
	    </td>  
	   <td align="right" colspan="2">  
	   <!--#include file="nav_menu.asp" -->
	
<script language="javascript"> 
<!--
buildsubmenus_horizontal();
//-->
</script>
	</td>
  </tr>
  <tr>
    <td colspan=2 valign="bottom" align="left"><span style="font: 7pt verdana; font-weight: bold;">&nbsp;<!--PERSON SELECTION FILTERS--></span></td>
  </tr>
  <tr valign="bottom">
  	    <td align="left" colspan=2 valign="bottom" ><span style="font: 7pt verdana; font-weight: bold;">&nbsp;PERSON SELECTION FILTERS</span>
                                                       <span style="font: 8pt verdana; font-weight: normal;">&nbsp;Program:</span>&nbsp;
			  <select name="program" onchange="refreshProg(this.value);">
				  <option value="">Any</option>
	                <%
	                Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                SQLStmt2.CommandText = "exec get_programs_for_staff '" & Session("user_name") & "'"  
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.write "SQL = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
  	                Do Until rs2.EOF
	                %>
	                <option value="<%=rs2("Program_ID")%>"><%Response.write rs2("Program_Name")%></option>
	                <%
	                rs2.MoveNext
                    Loop
                    
	                %>
                    
				</select>
&nbsp;&nbsp;
                <span style="font: 8pt verdana; font-weight: normal;">Location:</span>
				<select style="width: 200px" name="location" onchange="refreshLocation(this.value);" >
				  <option value="">Any</option>
	                <%
	                if Request.QueryString("pcode") <> "" THEN
	                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmt2.CommandText = "exec get_locations_for_staff_and_program " & Request.QueryString("pCode") & ",'" & Session("user_name") & "'"
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                    'response.write "SQL = " & SQLStmt2.CommandText
  	                    rs2.Open SQLStmt2
  	                    Do Until rs2.EOF
	                    %>
	                    <option value="<%=rs2("Location_id")%>"><%=rs2("Description")%></option> 
	                    <%
	                    rs2.MoveNext
                        Loop
                    end if
	                    %>
				</select>
			
	    </td>
  </tr>
</table>
   <div id="spacer5px"></div>
	<table width="100%" cellpadding="0" cellspacing="0" border="0" align="center" style="padding-top: 0px;">
	  <tr>
	    <td valign="middle" align="center" style="padding-left: 2px;">
		  <table cellpadding="0" cellspacing="0" width="100" height="115" class="pod">
		    <tr>
			  <td align="center">
			  <% if client_has_picture = "Yes" THEN %>
			    <a href="#" onclick="popupwindow('display_orig_image.asp?cid=<%=Request.QueryString("cid")%>',800,1000,'manageClientsWindow');"><img src="display_image.asp?cid=<%=Request.QueryString("cid")%>" width="95" height="120" border="0" /></a>
			  <% else %>  
			  Person Photo
			  <% end if %>
			  </td>
			</tr>
		  </table>
		</td>
		<td valign="top" align="center">
		  <table width="370" cellpadding="0" cellspacing="0" border="0">
		    <tr>
			  <td class="pod_title" style="padding-left: 0px;" align="left">SELECTED PERSON</td>
			  <td align="right">
			  <% if Session("active_staff_create_client_auth") = "1" THEN %><img src="/images/icon_add_person.png" width="7" height="7" border="0" alt="Add Person" title="Add Person" style="padding-right: 4px;" /><a href="#" onclick="popupwindow('manage_clients.asp?',600,985,'manageClientsWindow');" class="addPerson">ADD NEW PERSON</a><%end if %>
			  </td>
			</tr>
		  </table>
		  <table cellpadding="1" cellspacing="0" width="370" height="102" class="pod_selected_client" border=0>
		    <tr>
			  <td align="left" valign="top">
			    <table width="100%" cellpadding="0" cellspacing="1" border="0">
				  <tr>
					<td valign="top" class="blackFontSmall">Record #: <span class="blueFontSmall"><b><%=rec_num_id%></b></span></td>
					<td valign="middle" class="blackFontSmall" nowrap ><b><i>last name lookup:</i></b> <input type="text" size="4" value="" name="person_search_filter" onblur="refreshClient(document.clientForm.client_id.value);" onkeyup="javascript:Person_Search_Filter();" /></td>
					<td valign="top" class="blackFontSmall" align="right">AGE:<span class="blueFontSmall">(<b><%=client_age%></b>)</span></td>
				  </tr>

                            </table>
			    <table width="100%" cellpadding="0" cellspacing="2" border="0">
				  <tr>
				    <td   class="blackFontSmall"><b>Person Name:</b></td>
					<td valign="bottom" colspan="2">
					<div id="clients">
					<select name="client_id" onchange="refreshClient(this.value);" style="width:245px;">
				  <option value="">Select Person</option>
	                <%
	                'Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                'Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                'SQLStmt2.CommandText = "exec get_clients_for_staff '" & Session("user_name") & "','" & Request.QueryString("pcode") & "','" & Request.QueryString("lCode") & "'"   
  	                'SQLStmt2.CommandType = 1
  	                'Set SQLStmt2.ActiveConnection = conn
  	                'SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.write "SQL = " & SQLStmt2.CommandText
  	                'rs2.Open SQLStmt2
  	                rsClientsForStaff.MoveFirst
                    Do Until rsClientsForStaff.EOF
  	                
	                if rsClientsForStaff("is_active_in_program") = 1 and Request.QueryString("sd") <> "true" THEN
  	                %>
	                <option value="<%=rsClientsForStaff("client_id")%>" <%if rsClientsForStaff("client_has_alerts") = "Yes" THEN %>style="color:red; font-weight:bold;"<%end if %>><%Response.write rsClientsForStaff("Last_Name") & ", " & rsClientsForStaff("First_Name") & " - " & rsClientsForStaff("DOB")%></option>
	                <%
  	                elseif rsClientsForStaff("is_active_in_program") = 0 and Request.QueryString("sd") = "true" THEN
  	                %>
	                <option value="<%=rsClientsForStaff("client_id")%>" <%if rsClientsForStaff("client_has_alerts") = "Yes" THEN %>style="color:red; font-weight:bold;"<%end if %>><%Response.write "* " & rsClientsForStaff("Last_Name") & ", " & rsClientsForStaff("First_Name") & " - " & rsClientsForStaff("DOB")%></option>
	                <%
  	                end if
	                
	                rsClientsForStaff.MoveNext
                    Loop
	                %>
				  </select>
				  </div>
				</td>
			  </tr>
			  <tr>
			    <td valign="middle" nowrap><input type="checkbox" class="radio" name="show_discharged" onclick="filterClients()" <%if Request.QueryString("sd") = "true" THEN %>checked<%end if %> />  Show Discharged </td>
			  </tr>
			</table>
            <div id="spacer6"></div>
            <table width="100%" cellpadding="2" cellspacing="2" border="0">
		      <tr>
				<td valign="top" class="blackFontSmall">Registration Date:&nbsp; <span class="blueFontSmall"><b><%=client_reg%></b></span> 
                    

                  

				</td>
				<td valign="middle" align="right" rowspan="2"> 
				   <% if Session("active_staff_create_client_auth") = "1" and client_id <> "" THEN %>
	                    <span style="float:right;"><input type="button" onclick="editClient();" value="Edit" class="submit" /></span>
 	               <%end if %>
                   
				 </td>
              </tr>
              <tr>
                <td valign="top" class="blackFontSmall">SSN:&nbsp; <span class="blueFontSmall"><%if client_ssn = "&nbsp;" THEN%><%=client_ssn%><%else%>XXX-XX-<b><%=RIGHT(client_ssn,4)%></b><%end if %></span></td>
				<td> </td>
	           </tr>   
             </table>
			           		
			  </td>
			</tr>
		  </table>
		</td>
		<td valign="top" align="center" class="blackFontSmall"> <b>INFORMATION SUMMARY</b>
                   <table cellpadding="0" cellspacing="2" width="488" height="102" class="pod_selected_client" border="0" align="center">
		    <tr>
			  <td align="left" valign="top">
			    <table width="100%" cellpadding="2" cellspacing="0" border="0">

				  <tr>
                                    <td width="120"> <a href="javascript: void(0);" onclick="filterTimeFrame2('CURRENT');"><b>Record View</b></a>
                                    </td>
				 <!--   <td class="displayFontLarge" width="50"><font color="#663366"><b>ALERTS:</b></font></td>
					<td width="20" align="left" class="displayFontLarge"><font color="#663366"><b><%=total_alerts_for_header %></b></font></td>	

				    -->
					<td class="blackFontSmall" width="67">&nbsp;In-Process:</td>
					<td width="45" align="left" class="displayFontLarge" align="left"><a href="javascript: void(0);" onclick="filterTimeFrame2('INPROCESS');"><b><%=in_proc_count %></b></a></td>

					<td class="blackFontSmall" width="57">&nbsp;Finalized:</td>
					<td width="55" align="left" class="displayFontLarge" align="left"><a href="javascript: void(0);" onclick="filterTimeFrame2('FINALIZED');"><b><%=finalized_count %></b></a></td>

                                        <td class="blackFontSmall" width="35">&nbsp;TOTAL:</td>
					<td width="60" align="left" class="displayFontLarge" ><a href="javascript: void(0);" onclick="filterTimeFrame2('ALL');"><b><%=total_form_count %></b></a></td>
			      </tr>
				  <tr>
				    <td colspan=7>

	
	  
   <%     if client_id <> "&nbsp;" and ( INT(total_form_count) > 0 or Request.QueryString("show_dashboard_flag") = "0") and done_filter_section = 0  THEN
	%>
	
          
             
           
	        
	    <span  style="font: 8pt verdana, arial, helvetica, sans-serif; font-weight: bold; color: blue; float:left;padding-top:6px;padding-left:1px;">

             VIEW FORMS:</span>
	    <div style="float:left;padding-top:6px;padding-left:3px;"><select style="font-size:10px; font-weight: bold; color: blue;" name="form_group" id="form_group">
            <option value="">Select Form Type</option>
	<%

	    Set SQLStmtGroups = Server.CreateObject("ADODB.Command")
  	    Set rsGroups = Server.CreateObject ("ADODB.Recordset")
  	    SQLStmtGroups.CommandText = "exec get_client_form_counts " & Request.QueryString("cid") 
  	    SQLStmtGroups.CommandType = 1
  	    Set SQLStmtGroups.ActiveConnection = conn
        'if Session("user_name") = "pwcard" THEN
      	'    response.Write "sql = " & SQLStmtGroups.CommandText
  	    'end if 
        SQLStmtGroups.CommandTimeout = 180 'Timeout per Command
  	    rsGroups.Open SQLStmtGroups
  	    
  	    form_types_count = 0

        done_filter_section = 0
  	    
  	    Do Until rsGroups.EOF
  	    
  	       
  	        
  	            form_types_count = form_types_count + 1
  	            
                  if Request.QueryString("soft") = rsGroups("form_type") then
                  response.write "<option value='" & rsGroups("form_type") & "' selected>" & rsGroups("form_description") & "(" & rsGroups("form_type_count") & ")</option>"
                  else
                   response.write "<option value='" & rsGroups("form_type") & "'>" & rsGroups("form_description") & "(" & rsGroups("form_type_count") & ")</option>"
                  end if

  	          '  if form_types_count mod 70 = 0 THEN
  	            '    response.Write "<br/>"
  	           ' else
  	            '    response.Write " : "
  	           ' end if 
  	            
  	      
  	    
  	        rsGroups.MoveNext
  	    Loop
  	    %>
            </select></div>
                  
            </div>
  	  
  	<%    
	end if
          %>


  
                                    </td>
				  </tr>
				</table>

                              <div><hr></div>
    
                <table width="100%" cellpadding="1" cellspacing="1" border="0">
                                  <tr>				
                                   <td colspan="1" class="blackFontSmall" width="77" valign="top" align="left" style="padding-right:0px; padding-top:2px;"><b>Last Update:</b></td>
					<td class="blueFontSmall" colspan="3"><span class="blueFontSmall" align="left" style="padding-left:0px; padding-top:0px;"><%if last_update_date <> "" THEN %><%=last_update_date%>&nbsp; <%=last_update_form %>&nbsp;<b><%=last_update_user%></b><%end if %></span>
                                   </td>   

				  </tr>
			
				    
				  <tr>		
                      <% 
                          colspan="colspan='2'"
                          

                     
                          
                          if next_isp_date <> "" THEN 
                          
                          
                          %>		
                                   <td class="blackFontSmall" width="77" valign="top" align="left" style="padding-right:0px; padding-top:2px;"><b>ISP Date:</b></td>
					<td class="blueFontSmall" >
                        <a href="javascript:void();" style="color:Blue; font-weight:bold; font-size:14px;" onclick="decideShowISPDiv();"><span class="blueFontSmall" align="left" style="padding-left:0px; padding-top:0px;"><b><%=next_isp_date%></b></span></a>
                    </td>   
                        <%   
                            colspan=""
                       end if 

                         if dhsp_date <> "" then
                          
                           %>
                        <td <%=colspan%> class="blackFontSmall" width="77" valign="top" align="left" style="padding-right:0px; padding-top:2px;"><b>DHSP Date:</b></td>    
                      <td <%=colspan%>  class="blueFontSmall">
                        <a href="javascript:void();" style="color:Blue; font-weight:bold; font-size:14px;" onclick="decideShowISPDiv();"><span class="blueFontSmall" align="left" style="padding-left:0px; padding-top:0px;"><b><%=dhsp_date%></b></span></a>
                    </td> 
                        <%  end if  %>  
				  </tr>
                    <% if next_isp_date <> "" THEN
                        if cur_staff_role_desc = "Administrator" THEN %>
                  <tr>
                      <% if isp_assessments_status = "Red" THEN %>
                      <td colspan="4" class="blackFontSmall" width="90" valign="top" align="left" style="padding-right:0px; padding-top:4px; color:#FF0033;"><b>HCSIS ISP Needs Immediate Attention</b>
                        <a href="javascript:void();" style="color:#3201a5;" onclick="decideShowISPDiv();"><b>[Click to see Details]</b></a>
                    </td> 
                      <%  elseif isp_assessments_status = "Orange" THEN %>
                      <td colspan="4" class="blackFontSmall" width="90" valign="top" align="left" style="padding-right:0px; padding-top:4px; color:#FF6633;"><b>HCSIS ISP staredwhy, needs to be finished</b>
                        <a href="javascript:void();" style="color:#3201a5;" onclick="decideShowISPDiv();"><b>[Click to see Details]</b></a>
                    </td> 
                      <% elseif isp_assessments_status = "Yellow" THEN %>
                      <td colspan="4" class="blackFontSmall" width="90" valign="top" align="left" style="padding-right:0px; padding-top:4px; color:#993300;"><b>HCSIS ISP up to date, but finished late</b>
                        <a href="javascript:void();" style="color:#3201a5;" onclick="decideShowISPDiv();"><b>[Click to see Details]</b></a>
                    </td> 
                      <% elseif isp_assessments_status <> "None" THEN %>
                      <td colspan="4" class="blackFontSmall" width="90" valign="top" align="left" style="padding-right:0px; padding-top:4px; color:#005166;"><b>HCSIS ISP up to date</b>
                        <a href="javascript:void();" style="color:#3201a5;" onclick="decideShowISPDiv();"><b>[Click to see Details]</b></a>
                    </td> 
                      <% end if %>                      
                  </tr>  
                <%  end if 
                   end if        %>

			   </table>

			  </td>
			</tr>
		  </table>
		</td>
	</table>
	</form>
	<form name="frmMain" action="index.asp" method="post">
	<%
	if total_alerts_for_header > 0 THEN
	    rsAlerts.MoveFirst
  	    
  	    if NOT rsAlerts.EOF THEN
  	        alert_counter = 1
  	%>
	<div class="pod_show_alert" align="center">
	    <table cellpadding="0" cellspacing="3" border="0" width="100%">
	            <tr>
	               <td align="left" colspan="2" bgcolor="#663366"><div class="displayFontLarge" style="margin-left: 2px; margin-right:10px;"><font color="#ffffff">ALERTS</font></div></td>      
	               <td bgcolor="#663366"><font color="#ffffff"><b>DUE DATE:</b></font></td>
	               <td bgcolor="#663366"><font color="#ffffff"><b>CREATED BY:</b></font></td>
	            </tr>
	            <%
	            Do Until rsAlerts.EOF
	                cur_alert_id = rsAlerts("Alert_ID")
	                cur_alert_type = rsAlerts("Alert_Type")
	                cur_target_id = rsAlerts("Alert_Target_Form_ID")
	                cur_trigger_id = rsAlerts("Alert_Trigger_Form_ID")
	                cur_due_within = rsAlerts("due_within_message")
	                cur_log_id = rsAlerts("Call_Log_Trigger_Row_ID")
	                cur_notice_message_p1 = rsAlerts("Notice_Message_P1")
	                cur_notice_message_p2 = rsAlerts("Notice_Message_P2")
	                cur_notice_message_p1_color = rsAlerts("Notice_Message_P1_Color")
	                cur_notice_message_p2_color = rsAlerts("Notice_Message_P2_Color")
	                cur_action_message_p1 = rsAlerts("Action_Message_P1")
	                cur_action_message_p2 = rsAlerts("Action_Message_P2")
	                cur_action_message_p1_color = rsAlerts("Action_Message_P1_Color")
	                cur_action_message_p2_color = rsAlerts("Action_Message_P2_Color")
	                cur_click_to_clear = rsAlerts("click_to_clear") 
	                cur_real_due_date = rsAlerts("real_due_date") 
	                
	                cur_concat_notice_message = ""
	                cur_concat_action_message = ""
	                
	                cur_day_count_color = "green"
	                
	                if cur_due_within <> "" THEN
	                    if cur_due_within <= 7 and cur_due_within >= 0 THEN
                            cur_day_count_color = "orange"
                        elseif cur_due_within < 0 THEN
                            cur_day_count_color = "red"
                        end if 
                    end if
	                
	                if cur_notice_message_p1 <> "" or cur_notice_message_p2 <> "" THEN
    	                cur_concat_notice_message = "<font color='" & cur_notice_message_p1_color & "'> " & cur_notice_message_p1 & " </font><font color='" & cur_notice_message_p2_color & "'> " & cur_notice_message_p2 & " </font>"
	                end if
	                
	                if cur_action_message_p1 <> "" or cur_action_message_p2 <> "" THEN                    
	                    cur_concat_action_message = "&nbsp;<font color='" & cur_action_mesage_p1_color & "'> " & cur_action_message_p1 & " </font><font color='" & cur_action_message_p2_color & "'> " & cur_action_message_p2 & " </font>"
	                end if
	            
	            if cur_concat_notice_message <> "" or cur_concat_action_message <> "" THEN
	            %>
	            <tr>
	                <td align="center"><b><font color="#006666"><%=alert_counter %>.</b></font></td>
	                <td align="left"><b><font color="#006666">
	                <% if cur_target_id <> "-1" THEN %>
	                    <a href="javascript: void(0);" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_target_id%>',800,1000,'<%=cur_target_id%>');"><%=cur_concat_notice_message%><% if cur_concat_action_message <> "" THEN response.Write cur_concat_action_message end if %><% if cur_due_within <> "" THEN Response.Write " within <font color='" & cur_day_count_color & "'>" & cur_due_within & " days</font>" end if %></a>
	                <% else %>
	                    <%=cur_concat_notice_message%><% if cur_concat_action_message <> "" THEN response.Write cur_concat_action_message end if %><%if cur_due_within <> "" THEN Response.Write " within <font color='" & cur_day_count_color & "'>" & cur_due_within & " days</font>" end if%>
	                <% end if %>
	                </b></font></td>
	                <td><b><font color="#006666"><%=cur_real_due_date%></b></font></td>
	                <td><b><font color="#006666"><%=rsAlerts("Creator_Name")%></b></font></td>
	            </tr>
	            <%
	            elseif cur_click_to_clear <> "0" THEN
	            %>
	            <tr>
	                <td align="center"><b><font color="#006666"><%=alert_counter %>.</b></font></td>
	                <td align="left"><b><font color="#006666"><a href="javascript: void(0);"><%=rsAlerts("Alert_Message")%><%if cur_due_within <> "" THEN Response.Write " within <font color='" & cur_day_count_color & "'>" & cur_due_within & " days</font>" end if%></a></b></font>&nbsp;&nbsp;
	                <a href="javascript: void();" onclick="return clickToClearAlert('<%=cur_alert_id%>');">
	                <img src="images/clear_btn.gif" border="0" alt="clear message"  />
	                </a>
	                </td>
	                <td><b><font color="#006666"><%=cur_real_due_date%></b></font></td>
	                <td><b><font color="#006666"><%=rsAlerts("Creator_Name")%></b></font></td>
	            </tr>
	            <%
	            elseif cur_target_id <> "-1" THEN
	            %>
	            <tr>
	                <td align="center"><b><font color="#006666"><%=alert_counter %>.</b></font></td>
	                <td align="left"><b><font color="#006666"><a href="javascript: void(0);" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_target_id%>',800,1400,'<%=cur_target_id%>');"><%=rsAlerts("Alert_Message")%><%if cur_due_within <> "" THEN Response.Write " within <font color='" & cur_day_count_color & "'>" & cur_due_within & " days</font>" end if%></a></b></font></td>
	                <td><b><font color="#006666"><%=cur_real_due_date%></b></font></td>
	                <td><b><font color="#006666"><%=rsAlerts("Creator_Name")%></b></font></td>
	            </tr>
	          <%
	            else	               
	                %>
	                <tr>
	                    <td align="center"><b><font color="#006666"><%=alert_counter %>.</b></font></td>
	                    <td align="left"><b><font color="#006666"><%=rsAlerts("Alert_Message")%><%if cur_due_within <> "" THEN Response.Write " within <font color='" & cur_day_count_color & "'>" & cur_due_within & " days</font>" end if%></b></font></td>
	                    <%if cur_real_due_date = "01/01/1900" then %>
                        <td><b><font color="#006666"></b></font></td>
                        <% else %>
                          <td><b><font color="#006666"><%=cur_real_due_date%></b></font></td>
                        <%end if %>
	                    <td><b><font color="#006666"><%=rsAlerts("Creator_Name")%></b></font></td>
	                </tr>
	                <% 	               
	            end if
	                
	            alert_counter = alert_counter + 1
	            rsAlerts.MoveNext
	            Loop
	            %>
	    </table>
	</div>
	<%
	    end if
	    
	    rsAlerts.Close
	    Set rsAlerts = Nothing
	end if
	%>
	<div class="pod_add_new_form" align="center">
	  <table cellpadding="0" cellspacing="3" border=0 width="100%">
	    <tr>
	      <td align="center" colspan="3">
	        <div align="left" style="float:left;"><input type="checkbox"  class="radio" name="program_default" disabled/> set as my default program</div>
	        <div class="displayFontLarge" style="margin-left: 2px; margin-right:10px; display:inline;" align="center" >ADD NEW FORM</div>
	      </td>                  
	    </tr>
	    <tr>
	      <td class="blackFontSmall" align="left"><b>Responsible Program:</b></td>
		  <td class="blackFontSmall" align="left"><b>Responsible Staff:</b></td>
		  <td class="blackFontSmall" align="left"><b>New Form to Add:</b></td>
	    </tr>
		<tr>
		  <td>
		    <select name="form_program" onchange="filterStaff();" <% if Request.QueryString("cid") = "" THEN%>disabled<%end if%>>
                <option value="">Select a Program</option>
                    <%
                    client_active = 0
                    
                        if Request.QueryString("cid") <> "" THEN
                        
	                        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	                        SQLStmt2.CommandText = "exec get_programs_for_staff_and_client '" & Session("user_name") & "'," & Request.QueryString("cid")  
  	                        SQLStmt2.CommandType = 1
  	                        Set SQLStmt2.ActiveConnection = conn
  	                        SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                        'response.write "SQL = " & SQLStmt2.CommandText
  	                        rs2.Open SQLStmt2
  	                        Do Until rs2.EOF
  	                        
  	                            client_active = 1
  	                            pif_age_require = rs2("Require_New_PIF_For_CA")
  	                            pif_max_age_hours = rs2("New_PIF_for_CA_max_hours")
	                %>
	                        <option value="<%=rs2("Program_ID")%>" pif_age_require="<%=pif_age_require%>" pif_max_age_hours="<%=pif_max_age_hours%>" ><%Response.write rs2("Program_Name")%></option>
	                <%
	                        rs2.MoveNext
                            Loop
	                    end if
	                %>
            </select>
          </td>
			  <td>
			    <div id="staff">
                    <select name="form_staff" disabled>
                        <option value="">Choose a Program</option>
                    </select>
                </div>
	      </td>
		  <td align="left">
		    <select name="create_form_type" onchange="reloadFormParents(this.value);" <% if Request.QueryString("cid") = "" THEN%>disabled<%end if%>>
		        <option value="">Select Form Type to Create</option>
			    <%
			        form_family_level = "0"
			    
			        if Request.QueryString("cid") <> "" THEN
			    
                        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
    	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
        
  	                     SQLStmt2.CommandText = "exec get_creatable_forms_for_user '" & Session("user_name") & "'," & Request.QueryString("cid")
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
                        'if Session("user_name") = "pwcard" THEN
     	                '    response.write "SQL = " & SQLStmt2.CommandText
  	                    'end if 
                        rs2.Open SQLStmt2
  	                    Do Until rs2.EOF
      	                
  	                        cur_required_forms = rs2("Required_Forms")
  	                        cur_form_family = rs2("Form_Family")
      	                                      	        
                  	         if rs2("requirements_met") = "Yes" or cur_required_forms = " " THEN
    	                    
	                            if (cur_form_family = "1" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "1"
	                            %>
	                            <optgroup label="Intake Forms" id="INTAKE">
	                            <%  
	                            elseif (cur_form_family = "2" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "2"
	                            %>
	                            <optgroup label="Emergency Information" id="EMERGENCY" style="color: #FF0000;">
	                            <%
	                            elseif (cur_form_family = "3" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "3"
	                            %>
	                            <optgroup label="Clinical Data Tracking" id="CDT" style="color: #0000FF;">
	                            <%
	                            elseif (cur_form_family = "4" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "4"
	                            %>
	                            <optgroup label="Communication" id="Optgroup2" style="color: #009900;">
	                            <%
	                            elseif (cur_form_family = "5" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "5"
	                            %>
	                            <optgroup label="AFC Forms" id="AFC">
	                            <%  
	                            elseif (cur_form_family = "6" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "6"
	                            %>
	                            <optgroup label="Assessments" id="ASSESS">
	                            <%  
	                            elseif (cur_form_family = "7" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "7"
	                            %>
	                            <optgroup label="Clinical Department Forms" id="CLIN">
	                            <%  
	                            elseif (cur_form_family = "8" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "8"
	                            %>
	                            <optgroup label="Day Services Forms" id="DAY">
	                            <%  
	                            elseif (cur_form_family = "9" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "9"
	                            %>
	                            <optgroup label="DDS ISP Forms" id="DDS">
	                            <%  
	                            elseif (cur_form_family = "x10" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x10"
	                            %>
                                <optgroup label="Family Support Forms" id="FS">
	                            <%  
	                            elseif (cur_form_family = "x11" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x11"
	                            %>
	                            <optgroup label="Universal Forms" id="HMEA">
	                            <%  
	                            elseif (cur_form_family = "x12" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x12"
	                            %>
	                            <optgroup label="Individual Support" id="ISI">
	                            <%  
	                            elseif (cur_form_family = "x13" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x13"
	                            %>
	                            <optgroup label="Medical Forms" id="MED">
	                            <%  
	                            elseif (cur_form_family = "x14" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x14"
	                            %>
	                            <optgroup label="Residential Forms" id="RES">
	                            <%  
	                            elseif (cur_form_family = "x15" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x15"
	                            %>
	                            <optgroup label="Employment" id="EMPLOYMENT">
	                            <%  
	                            elseif (cur_form_family = "x16" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x16"
	                            %>
	                            <optgroup label="HCSIS Forms" id="HCSIS">
	                            <%  
	                          
	                            elseif (cur_form_family = "x17" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x17"
	                            %>
                                      <optgroup label="Therapy Forms" id="THERAPY">
	                            <%  
	                          
	                            elseif (cur_form_family = "x18" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x18"
	                            %>
                               <optgroup label="MRC Forms" id="MRC">
	                            <%  
	                          
	                            elseif (cur_form_family = "x19" and cur_form_family <> form_family_level) THEN
	                                form_family_level = "x19"
	                            %>
	                            <optgroup label="Other Forms" id="OTHER">
	                            <%  
	                            end if                       
	                  %>                              
	                 
	                    <option value="<%=rs2("Form_Type")%>"><%Response.write rs2("Form_Description")%></option>
	                  <%
    	                        
	                        end if
    	                   
	                    rs2.MoveNext
                        Loop
                    end if
	           %>
			</select>
		  </td>
		  <td><div name="create_form_div" id="create_form_div"><a href="#" onclick="checkForm();" id="add_new_form_button"><img src="images/button_add_new.jpg" width="28" height="16" border="0" class="textmiddle" alt="Create New Form" title="Create New Form" /></a></div></td>
	    </tr>
	    
		<%if client_active = 0 and Request.QueryString("cid") <> "" THEN %>
		<tr>
		   <td colspan="5">
		        <font color="red"><b>Selected client is currently inactive in all programs</b></font>
		   </td>
	    </tr>
	    <%end if %>
	  </table>
	  <div id="possible_parent_forms">
	  </div>  
	</div>
	<%
	qs_pid = Request.QueryString("fn")
	
	if qs_pid = "" THEN
	    qs_pid = "1000"
	end if   	
	%>
        
	<table cellpadding="0" cellspacing="0" border="0" width="966">
	  <tr>
	    <td align=center>
		  <table id="filters_old" border="0">
	        <tr>
			<td  width="100%" class="blackFontSmall" align="center" valign="middle">
		        <% if Request.QueryString("cid") <> "" THEN %>
		      <div align="center" class="pod_add_new_form"> 
		      <table cellpadding="0" cellspacing="3" border=0 width="100%">
	            <tr>
	              <td align="center" colspan="3">
                  
<div style="float:left;padding-left:60px;"><a onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?ft=EFS_V2&cid=<%=client_id%>&pid=<%=intake_program_id %>&sid=<%=Session("Staff_ID") %>&dp=',800,960,'<%=cur_form_id%>');" href="javascript: void(0);"><span style="float:left;"><img src="images/EFS2.jpg" border="0" height="25px" width="25px" /></span><span style="font-size:12pt;font-weight: bold;padding-top:2px;float:left;">EFS</span></a></div>

	              <% if Session("active_staff_can_access_client_meds") = "1" THEN %>
	                 <div style="float:left;padding-left:100px;"><a href="javascript: void(0);" onclick="popupwindow('view_client_meds.asp?cid=<%=Request.QueryString("cid")%>',800,1000,'getClientMeds_<%=Request.QueryString("cid") %>');">
                                <div style="float:left;"><img src="images/medicationIcon-25.png"  border="0" /></div><div style="float:left;padding-top:5px;padding-left:3px;"><b>MEDICATIONS</b>:</a>(<b><%=client_meds_count%></b>)</div>
                                &nbsp;&nbsp;&nbsp;&nbsp;</div>
                    <% end if %>
                    
                    <% if Session("active_staff_can_access_client_diags") = "1" THEN %>
                        <div style="float:left;"><a href="javascript: void(0);" onclick="popupwindow('view_client_diags.asp?cid=<%=Request.QueryString("cid")%>',800,1000,'getClientDiags_<%=Request.QueryString("cid") %>');">
                                <div style="float:left;"><img src="images/diagnosisICon-25.png" border="0" /></div><div style="float:left;padding-top:5px;padding-left:3px;"><b>DIAGNOSIS</b>:</a>(<b><%=client_diags_count%></b>)</div>
                              &nbsp;&nbsp;&nbsp;&nbsp;</div>
                    <% end if %>
                    
                    <% if Session("active_staff_can_access_client_schedule") = "1" and 1=2 THEN %>
                        <div style="float:left;"><a href="javascript: void(0);" onclick="popupwindow('view_client_diags.asp?cid=<%=Request.QueryString("cid")%>',800,1000,'getClientDiags_<%=Request.QueryString("cid") %>');">
                                <div style="float:left;"><img src="images/calendar-2b25.png" border="0" /></div><div style="float:left;padding-top:5px;padding-left:3px;"><b>SCHEDULE</b></a></div>
                              &nbsp;&nbsp;&nbsp;&nbsp;</div>
                    <% end if %>
                    
                    <% if Session("active_staff_can_access_client_dr_visits") = "1" THEN %>
                     
                         
                          <div style="float:left;"><img src="images/doctorVisitIcon-25.png" border="0" /></div><div style="float:left;padding-top:5px;padding-left:3px;"><b><a class="goToDoctor" onclick="return false;" href="#">NEW DOCTOR VISIT\FOLLOW-UP(<%=recommendations_count %>)</a></b></div> &nbsp;&nbsp; <b>
                          <div style="float:left;">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/VisitResults_25.png" border="0" /></div><div style="float:left;padding-top:5px;padding-left:3px;"><b>
                             <%if hcef_count = 0 then  %>
                             <a class="backToDoctor" onclick="return false;" href="#" disabled>VISIT RESULTS:(<%=hcef_count %>)</a>
                           <%else %>
                               <a class="backToDoctor" onclick="return false;" href="#">VISIT RESULTS:(<%=hcef_count %>)</a>
                          <%end if %>
                       </b></div></div>
                    <% end if %>
	              </td>                  
	            </tr>
	         </table>
		    </div>
		        <% else %>
		        &nbsp;
		        <% end if  %>
		    </td>
	     
	        </tr>
		  </table>
	    </td>
	</tr>

	</table>
	
	</form>
	
	<table id="forms" border="0">
	  
   <% 
       
     
       
           if client_id <> "&nbsp;" and ( INT(total_form_count) > 0 or Request.QueryString("show_dashboard_flag") = "0") and done_filter_section = 0  THEN
	%>

  	<%    
	end if
          %>

	  <!--PUT ASP FOR PARENT FORM LOOP HERE-->
	  <%
	  if client_id <> "&nbsp;" THEN
            
            
         cur_form_name = ""
        last_cur_type = ""
        
        cur_filt_by = ""
        
        if Request.QueryString("fb") = "" THEN
            cur_filt_by = ""
        else
            cur_filt_by = Request.QueryString("fb")
        end if
        
        form_group_filter = Request.QueryString("fn")
        
        if form_group_filter = "" THEN
            form_group_filter = 1000
        end if
        
       

       

        ep_filt_val = Request.QueryString("episodes")
	                
	                if ep_filt_val = "" THEN
	                    ep_filt_val = cur_ep_id
	                end if

          form_group_counter = 0

           tf = ""
           cur_soft = ""

          if Request.QueryString("tf") <> "" then
            tf = Request.QueryString("tf")
          end if 

         if Request.QueryString("soft") <> "" and  Request.QueryString("soft") <> "ALL"  THEN
            cur_soft = Request.QueryString("soft")
         end if 


          if Request.QueryString("soft") = "CURRENT" THEN
            cur_soft = ""
            tf = "CURRENT"
           end if 

           if Request.QueryString("soft") = "INPROCESS" THEN
            cur_soft = ""
            tf = "INPROCESS"
           end if 


           if Request.QueryString("soft") = "FINALIZED" THEN
            cur_soft = ""
            tf = "FINALIZED"
           end if 

         if cur_soft="DHSP" and Request.QueryString("ufid")="" then
            tf = ""
          end if
        
        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	     SQLStmt2.CommandText = "exec get_form_list_for_client2 " & client_id & ",'" & Session("user_name") & "'," & form_group_filter & ",-1" & ",'" & tf & "','" & Request.QueryString("ufid") & "','" & cur_soft & "'"
  	    SQLStmt2.CommandType = 1
  	    Set SQLStmt2.ActiveConnection = conn
  	    SQLStmt2.CommandTimeout = 45 'Timeout per Command
        if session("user_name")="elarochelle" then
         ' response.write "SQL 1 = " & SQLStmt2.CommandText
        end if
  	    rs2.Open SQLStmt2
  	    Do Until rs2.EOF
  	       
  	    blob_val = ""
  	    cur_create_date = ""
  	    has_children = "No"
  	    cur_copied_form_id = ""
  	    cur_copied_form_type = ""
  	    cur_copied_form_create_date = ""
  	    user_secondary_header = 0
  	    
  	    if rs2("Unique_Form_ID") THEN
  	         
  	        Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	        Set rs3 = Server.CreateObject ("ADODB.Recordset")
    
  	        SQLStmt3.CommandText = "exec get_form_info_without_content " & rs2("Unique_Form_ID")
  	        SQLStmt3.CommandType = 1
  	        Set SQLStmt3.ActiveConnection = conn
  	        SQLStmt3.CommandTimeout = 45 'Timeout per Command
  	        'response.write "SQL = " & SQLStmt3.CommandText
  	        rs3.Open SQLStmt3

            cur_copied_form_id = rs3("Parent_Form_ID")
            blob_val = rs3("File_Content")
            cur_create_date = rs3("Create_Date")
            has_children = rs3("has_children")
            has_dhsp_pages = rs3("has_dhsp_pages")
            lastest_dhsp = rs3("lastest_dhsp")                
            cur_copied_form_type = rs3("copied_form_type")
            cur_copied_form_create_date = rs3("copied_form_create_date")  
            
	    end if
	    
	    if blob_val <> "" THEN
	            
	        if cur_form_name <> rs2("FormName") THEN
	            new_form_type = 1
	        else
	            new_form_type = 0
	        end if 
    	    
	        cur_form_id = rs2("Unique_Form_ID")
	        cur_form_name = rs2("FormName")        
	        cur_type = rs2("Form_Type")
    	    
	        use_header = 1
	        use_secondary_header = 0
    	                	    
            if cur_type = last_cur_type and last_cur_type <> "" and has_children <> "Yes" THEN
	            use_header = 0
	            use_secondary_header = 0
	        end if
    	    
	        if cur_type = last_cur_type and last_cur_type <> "" and has_children = "Yes" THEN
	            use_secondary_header = 1
	        end if
	        
	        if new_form_type = 1 THEN
	            use_header = 1
	        end if
    	    
	        cur_parent_type = rs2("Parent_Form_Type")
	        cur_req_forms = rs2("Required_Forms")
	        cur_num_pages = rs2("Num_Pages")
	        cur_est_comp_time = rs2("Est_Completion_Time")
	        cur_linked_form_id = rs2("Linked_Form_ID")
	        cur_access = rs2("Access_Level")
	        cur_status = rs2("Status")
	        cur_create_user = rs2("Create_User")
	        cur_last_update = rs2("Update_Date")
	        cur_has_history = rs2("Has_History")
	        cur_has_attachments = rs2("Has_Attachments")
	        cur_has_notes = rs2("Has_Notes")
	        cur_has_billing_notes = rs2("Has_Billing_Notes")
	       
	        cur_create_user_name = rs2("Create_User_Name")
	        cur_main_prog = rs2("Main_Program")
	        cur_main_prog_id2 = rs2("Main_Program_id")
	        cur_main_staff = rs2("Main_Staff")
	        current_doa = rs2("form_dos_date")
	        cur_hide_show_count = rs2("hide_show_count")
	        'cur_external_system_id = rs2("external_system_id")   
           	    
	        cur_req_met = 0
	        cur_req_msg = ""


            if current_doa <> "" and cur_status = "Finalized" and cur_type = "HCEF" THEN
                      cur_last_update = current_doa
                      cur_status = "DOA"
            end if


           if current_doa <> "" and cur_status = "Finalized" and (cur_type = "DIPN" or cur_type = "AFC_CASE_NOTE_V2" or cur_type = "CASE_NOTE_V4" ) THEN
                      cur_last_update = current_doa
                      cur_status = "DOS"
           end if

           if current_doa <> "" and cur_status = "Finalized" and (cur_type = "DHSP" or cur_type="DHMPN") THEN
                      cur_last_update = current_doa
                      cur_status = "DOP"
           end if
    	    
    	    cur_form_hover = "This form has " & cur_num_pages & " page(s)" ' and was created on " & cur_create_date & "."
    	    
	        if cur_req_msg <> "" THEN
	            cur_req_msg = "This form requires the completion of " & cur_req_msg & " before it can be started"
	        end if
    	    
	        if use_header = 1 and use_secondary_header <> 1 THEN
	      %>
    	  
	      <tr>
	        <td class="formRow" colspan="7" valign="bottom" <% if cur_type = "PAPER" THEN %>style="background-color:#FFFBD2;"<% end if %>>
	        
	       <%

                 if cur_type <> "PAPER" and cur_hide_show_count > 1 and (Request.QueryString("tf")="ALL" and cur_type="DHSP")  and Request.QueryString("ufid")<>(cur_form_id) and (Request.QueryString("ulfid") <> cur_linked_form_id or Request.QueryString("uft") <> cur_type) THEN 
                
                if tf="CURRENT" or cur_type=Request.QueryString("soft") or tf="INPROCESS" then
                %>
	            <a href="javascript:void();" onclick="undoFilter('<%=cur_form_id%>','<%=cur_linked_form_id %>','<%=cur_type%>');" class="filterOff" style="text-decoration:none;" alt="Showing Filtered Forms"><font style="font-size:14px;" >+</font></a>
               <% else  %>
                 <a href="javascript:void();" onclick="undoFilter('<%=cur_form_id%>','<%=cur_linked_form_id %>','');" class="filterOff" style="text-decoration:none;" alt="Showing Filtered Forms"><font style="font-size:14px;" >+</font></a>

	        <% end if
                
              end if  %>
	        
	        <% if cur_type = "CBFSSN" THEN %>
	            <%if Request.QueryString("fn") = "all_cbfssn" THEN %>
	                <%if cur_CBFSSN_total_count > 7 THEN %>
	                <a href="#" onclick="filterNotes('less_cbfssn');" class="filterOn" alt="Showing All CBFS Service Notes">[see less]</a>
	                <% end if %>
	            <%else %>
	                <%if cur_CBFSSN_showing_count < cur_CBFSSN_total_count THEN %>
	                <a href="#" onclick="filterNotes('all_cbfssn');" class="filterOff" alt="Showing Last <%=cur_CBFSSN_showing_count%> CBFS Service Notes">[see more]</a>
	                <% end if %>
	            <%end if %>
	            
	        <% elseif cur_type = "CBFSSNWE" THEN %>
	            <%if Request.QueryString("fn") = "all_cbfssnwe" THEN %>
	                <a href="#" onclick="filterNotes('less_cbfssnwe');" class="filterOn" alt="Showing All CBFS Weekly Service Notes">[see less]</a>
	            <%else %>
            	    <% if cur_CBFSSNWE_showing_count < cur_CBFSSNWE_total_count THEN %> 
	                <a href="#" onclick="filterNotes('all_cbfssnwe');" class="filterOff" alt="Showing Most Recent CBFS Weekly Service Note Per Week">[see more]</a>
	                <% end if %>
	            <%end if %>
	        <% elseif cur_type = "ESPLBN" THEN %>
	            <%if Request.QueryString("fn") = "all_log" THEN %>
	                <%if cur_esp_log_bill_total_count > 7 THEN %>
	                <a href="#" onclick="filterNotes('less_log');" class="filterOn" alt="Showing All ESP Log/Billing Notes">[see less]</a>
	                <% end if %>	                
	            <%else %>
            	    <% if cur_esp_log_bill_showing_count < cur_esp_log_bill_total_count THEN %> 
	                <a href="#" onclick="filterNotes('all_log');" class="filterOff" alt="Showing Last <%=cur_esp_log_bill_showing_count%> ESP Log/Billing Notes">[see more]</a>
	                <% end if %>
	            <%end if %>
	            
	        <%elseif cur_type <> "PAPER" THEN %>
	              <%if cur_type = "GCFTR" or cur_type = "CONTACT2" THEN
                 
                   Set SQLStmtComp = Server.CreateObject("ADODB.Command")
  	            Set rsComp = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmtComp.CommandText = "select Staff_ID from Staff_Master where user_name = '" & Session("user_name") & "'"
  	            
  	            SQLStmtComp.CommandType = 1
  	            Set SQLStmtComp.ActiveConnection = conn
  	            SQLStmtComp.CommandTimeout = 45 'Timeout per Command
  	            rsComp.Open SQLStmtComp

                if NOT rsComp.EOF THEN
  	                Staff_ID = rsComp("Staff_ID")
  	             
  	            end if

                 if cur_type = "CONTACT2" then
                  %>
                 <img src="images/form_iconSmall.gif" border="0"> 
	            <a onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=cur_form_id%>&ft=<%=cur_type%>&cid=<%=client_id%>&pid=<%=cur_main_prog_id2 %>&sid=<%=Staff_ID %>&dp=&if=Fix',800,960,'<%=cur_form_id%>');" href="#"><img title="Update Form" alt="" src="images/plus_sign.gif" border="0"></a>
                 <% else %> 
                  <img src="images/form_iconSmall.gif" border="0"> 
	            <a onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=cur_form_id%>&ft=<%=cur_type%>&cid=<%=client_id%>&pid=<%=cur_main_prog_id2 %>&sid=<%=Staff_ID %>&dp=&if=Fix',800,1400,'<%=cur_form_id%>');" href="#"><img title="Update Form" alt="" src="images/plus_sign.gif" border="0"></a>
                    <% end if
                     
                     else %> 
                 <img src="images/form_iconSmall.gif" border="0">
                 <%end if %>
	        <%end if %>
	        
	        <%if cur_access = "V" or cur_status = "Group Lock" THEN%>
	            <font color="#777777">
                     <%if cur_type = "DHSP" then %>
	                     <%=cur_form_name%>&nbsp;-&nbsp; [plan date: <%=cur_last_update%>]
                        <%else %>
                        <%=cur_form_name%>
                         <%end if %>


	            </font>
	        <%else%>
	           <%if cur_type = "PAPER" THEN%>
	             <font color="#333300">
                     
                       <%if cur_type = "DHSP" then %>
	                     <%=cur_form_name%>&nbsp;-&nbsp; [plan date: <%=cur_last_update%>]
                        <%else %>
                        <%=cur_form_name%>
                         <%end if %>



	             </font>	 
	             <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&cid=<%=client_id%>',800,1000,'getAttachmentsWindow');">&nbsp;&nbsp;<%if cur_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg" border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg" border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>            
	           <%else%>
                <%if cur_type = "DHSP" then %>
	             <%=cur_form_name%>&nbsp;-&nbsp; [plan date: <%=cur_last_update%>]
                <%else %>
                <%=cur_form_name%>
                 <%end if %>
	           <%end if%>    
	        <%end if%>
	        
	        </td>
   	      </tr>
   	         	      
	      <%
	      end if 
    	  
	      '_______________________2nd header choice
	      if use_secondary_header = 1 THEN
	      %>
    	  
	      <tr>
	        <td class="formRow" colspan="7" valign="bottom" <% if cur_type = "PAPER" THEN %>style="background-color:#FFFBD2;"<% end if %>>
	            
            
             <%if cur_type <> "PAPER" and cur_hide_show_count > 1 and Request.QueryString("ufid") <> cur_form_id THEN %>
	        <a href="javascript:void();" onclick="undoFilter('<%=cur_form_id%>','<%=cur_linked_form_id %>','<%=cur_type%>');" class="filterOff" style="text-decoration:none;" alt="Showing Filtered Forms"><font style="font-size:14px;" >+</font></a>
	        <% end if  %>
            <img src="images/form_iconSmall.gif" border="0">  
	        <%if cur_access = "V" or cur_status = "Group Lock" THEN%>
	            <font color="#777777">
                    
                     <%if cur_type = "DHSP" then %>
	                     <%=cur_form_name%>&nbsp; <%=cur_last_update%>
                        <%else %>
                        <%=cur_form_name%>
                         <%end if %>



	            </font>
	        <%else%>
	           <%if cur_type = "PAPER" THEN%>
	             <font color="#333300">
                     
                        <%if cur_type = "DHSP" then %>
	                     <%=cur_form_name%>&nbsp; <%=cur_last_update%>
                        <%else %>
                        <%=cur_form_name%>
                         <%end if %>



	             </font>
	             <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&cid=<%=client_id%>',800,1000,'getAttachmentsWindow');">&nbsp;&nbsp;<%if cur_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg" border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg" border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>
	           <%else%>
	             <%if cur_type = "DHSP" then %>
	             <%=cur_form_name%>&nbsp; <%=cur_last_update%>
                <%else %>
                <%=cur_form_name%>
                 <%end if %>
	           <%end if%>    
	        <%end if%>
	        </td>
	      </tr>
	      <%
	      end if
    	  
	      if blob_val <> "" THEN
	      %>
	      
	      <tr>
	        <% if cur_type <> "PAPER" THEN %>
	        <td class="buttons">
               
            <%if cur_staff_role_desc = "Administrator" or (cur_status = "In-Process" and (cur_access = "E" or cur_access = "L")) THEN 
            %>
                <% if (cur_create_user = Session("user_name") or cur_staff_role_desc = "Administrator") and cur_type <> "PAPER" THEN %>
                <a href="#" onclick="confirmFileDelete('<%=cur_form_name%>',<%=cur_form_id%>);"><img title="Delete in-process form" src="images/delete_file.gif" border="0"/></a>
                <%end if %>
            <%end if %>
    
            <!-- NEW FORM COPY PROCESS FOR IAPs-->
            <%
            if cur_type <> "PAPER" and cur_type <> "INTAKE" and cur_type <> "DHSP"  THEN %>
                <a href="choose_parent_for_copy.asp?width=700&height=300&uid=<%=cur_form_id%>&cid=<%=client_id%>&usid=<%=user_staff_id %>&ft=<%=cur_type %>" title="FAQs" class="thickbox"><img src="images/MSDPCopyIcon.jpg" title="Make a copy of this form" border="0"></a>
             <%
            elseif  cur_type = "DHSP" and CDate(cur_create_date) > CDate("06/25/2018") and (cur_status="DOS" or cur_status = "DOP" or cur_status="Finalized") and has_dhsp_pages="Yes" and lastest_dhsp=cur_form_id then %>
                 <a href="javascript:void(0);" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_form_id%>&if=Fix',800,1000,'<%=current_type %>_copy_' + '<%=Request.QueryString("cid")%>');"><img src="images/plus_sign.gif"  border="0" title="Edit Form" alt="Edit Form" /></a>&nbsp;
            <% elseif  cur_type = "DHSP" and CDate(cur_create_date) > CDate("06/25/2018") and (cur_status="DOS" or cur_status = "DOP" or cur_status="Finalized") and has_dhsp_pages="No" and lastest_dhsp=cur_form_id then %>

                 <a href="javascript:void(0);" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=cur_form_id%>&ft=<%=cur_type%>&cid=<%=client_id%>&pid=<%=cur_main_prog_id2 %>&sid=<%=user_staff_id %>&dp=&if=Fix',800,1000,'<%=cur_form_id%>');"><img src="images/plus_sign.gif"  border="0" title="Edit Form" alt="Edit Form" /></a>&nbsp;

            <%end if %> 

            <%if cur_type <> "PAPER" THEN %>
                <a href="#" onclick="popupwindow('view_history.asp?uid=<%=cur_form_id%>',800,1000,'getHistoryWindow');"><img src="images/button_history.jpg"  border="0" alt="History" title="History" /></a>
                <a href="#" onclick="popupwindow('view_notes.asp?uid=<%=cur_form_id%>',800,1000,'getNotesWindow');">
                <%if cur_has_notes = "Yes" THEN %>
                    <%if cur_has_billing_notes = "Yes" THEN %>
                    <img src="images/button_notes_clip_dollar.jpg"  border="0" alt="Notes" title="Notes" />
                    <%else %>
                    <img src="images/button_notes_clip.jpg"  border="0" alt="Notes" title="Notes" />
                    <%end if %>
                <%else %>
                    <%if cur_has_billing_notes = "Yes" THEN %>
                    <img src="images/button_notes_dollar.jpg"  border="0" alt="Notes" title="Notes" />
                    <%else %>
                    <img src="images/button_notes.jpg" border="0" alt="Notes" title="Notes" />
                    <%end if %>
                <%end if %>
                </a>                
            
                <%if (cur_status = "Finalized" or cur_access = "V" or cur_status = "Group Lock" or cur_status = "DOA" or cur_status = "DOS" or cur_status = "DOP") THEN%>
                    <% if cur_type = "BTMS" or cur_type = "CONTRACTS" or cur_type = "DMTC" or cur_type = "DHSPS" or cur_type = "FDIP" or cur_type = "FFSE" or cur_type = "GCFTR" or cur_type = "FMTP" or cur_type = "FTR" or cur_type = "FTR_CASH" or cur_type = "FTR_BANK"  or cur_type = "GCFTR" or cur_type = "HIPPA" or cur_type = "IMMUNE" or cur_type ="MEDLIST" or cur_type ="MMR" or cur_type ="MRC" or cur_type ="RTF" or cur_type ="SPMU" or cur_type ="CLOG" or cur_type ="MS" or cur_type ="PAS" or cur_type ="MS_V2" or cur_type ="MS_PROC" or cur_type ="FIRE_GROUP" or cur_type ="MMR" or cur_type ="DDS" or cur_type ="SLMMS" THEN%>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                    <% else %>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_form_id%>',800,1000,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                    <% end if %>
                <%elseif cur_status = "In-Process" and (cur_access = "E" or cur_access = "L") THEN%>
                    <% if cur_type = "BTMS" or cur_type = "CONTRACTS" or cur_type = "DMTC" or cur_type = "DHSPS" or cur_type = "FDIP" or cur_type = "FTR_CASH" or cur_type = "FTR_BANK"  or cur_type = "GCFTR" or cur_type = "FFSE" or cur_type = "FMTP" or cur_type = "GCFTR" or cur_type = "FTR" or cur_type = "HIPPA" or cur_type = "CLOG" or cur_type = "IMMUNE" or cur_type ="MMR" or cur_type ="MEDLIST" or cur_type ="MRC" or cur_type ="RTF" or cur_type ="SPMU" or cur_type ="MS" or cur_type ="PAS" or cur_type ="MS_V2" or cur_type ="MS_PROC" or cur_type ="MMR" or cur_type ="FIRE_GROUP" or cur_type ="DDS" or cur_type ="SLMMS" THEN%>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                    <% else %>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=cur_form_id%>&from=dash',800,1000,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                    <% end if %>
                <%end if %>
            <% end if %>
 
            
            <% if cur_type <> "PAPER" THEN %>
            <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&cid=<%=client_id%>',800,1000,'getAttachmentsWindow');"><%if cur_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg"  border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg"  border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>&nbsp; 
         
            
            <% end if  %>
        </td>
    <% end if  %>
    
    <%if cur_type = "PAPER" THEN %>
            
              <td colspan="7" class="blackFontSmall" align="center" style="background-color:#FFFBD2;"> 
              <%
             Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	         Set rs3 = Server.CreateObject ("ADODB.Recordset")
                        
  	         SQLStmt3.CommandText = "exec get_doc_type_counts_for_client " & Request.QueryString("cid")
  	         SQLStmt3.CommandType = 1
  	         Set SQLStmt3.ActiveConnection = conn
  	         'response.write "SQL = " & SQLStmt3.CommandText
  	         rs3.Open SQLStmt3
                	            
	         Do Until rs3.EOF
               	    
	            cur_doc_category = rs3("category")
	            cur_doc_desc = rs3("desc")
	            cur_doc_count = rs3("doc_count")
                 if cur_doc_category <> "Medications" then
	            %>
	                <span style="white-space: nowrap;"><% if cur_doc_count <> 0 THEN %><a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&doc_type=<%=cur_doc_category %>&doc_subtype=All&cid=<%=client_id%>',800,1000,'getAttachmentsWindow');"><b><%=cur_doc_desc %></b></a><% else %><%=cur_doc_desc %> <% end if  %> (<b><%=cur_doc_count%></b>)</span>&nbsp;
	            <%
                end if

	            rs3.MoveNext
                	        
	         Loop %>              
</td>
            
    <% else %>
		    <td class="status">
		    <%if has_the_needs = "1" THEN%><font color="#993300"><%=cur_status%></font>
		    <%elseif cur_status = "Finalized" or cur_status = "DOA" or cur_status = "DOS" or cur_status = "DOP" or cur_status = "Group Lock" THEN%><span class="finalized"><%=cur_status%></span>
		    <%else%><%=cur_status%>
		    <%end if%>
		    <br />
		    <%if cur_copied_form_id <> "" and cur_copied_form_id <> "0" THEN %>
		    <div onmouseover="showToolTip(event,'Copied From <%=cur_copied_form_type%> with create date of <br / > <%=cur_copied_form_create_date%> <br/><br /> Created on <%=cur_create_date%> <br /> By <%=REPLACE(cur_create_user_name, "'", "\")%> <br /> Program: <%=cur_main_prog %> <br /> Staff: <%=REPLACE(cur_main_staff,"'", "\'") %>')" onmouseout="hideToolTip()">
		    <%else%>
		    <div onmouseover="showToolTip(event,'Created on <%=cur_create_date%> <br /> By <%=REPLACE(cur_create_user_name, "'", "\'")%> <br /> Program: <%=cur_main_prog %> <br /> Staff: <%=REPLACE(cur_main_staff, "'", "\'") %>')" onmouseout="hideToolTip()">
		    <%end if %>
		    <%if has_the_needs = "1" THEN%><font color="#993300"><%=cur_last_update%></font>
		    <%elseif cur_status = "Finalized" or cur_status = "DOA" or cur_status = "DOS" or cur_status = "DOP" or cur_status = "Group Lock" THEN%><span class="finalized"><%=cur_last_update%></span>
		    <%else%><%=cur_last_update%>
		    <%end if%></div>
		    </td>
		    
		    <td class="current_sigs" colspan="5">
		    <%
		    Set SQLStmtSI = Server.CreateObject("ADODB.Command")
  	        Set rsSI = Server.CreateObject ("ADODB.Recordset")
    
  	        SQLStmtSI.CommandText = "exec get_signatures_completed_info_for_staff_and_form " & rs2("Unique_Form_ID") & ",'" & Session("user_name") & "'"
  	        SQLStmtSI.CommandType = 1
  	        Set SQLStmtSI.ActiveConnection = conn
  	        SQLStmtSI.CommandTimeout = 45 'Timeout per Command
  	       if Session("user_name") = "elarochelle" THEN
               ' response.write "SQL1 = " & SQLStmtSI.CommandText
            end if 

                  if Session("user_name")="elarochelle" then
                          '  response.write "SQL1 = " & SQLStmtSI.CommandText
                        end if

  	        rsSI.Open SQLStmtSI
		    
		    cur_sig_count = 0
		    
		    Do Until rsSI.EOF
		    
		        datalock_temp = rsSI("Completed_Hash")
	            cur_sig_req_type = rsSI("Required_Type")
	            cur_sig_comp_type = rsSI("Completed_Type")
	            cur_sig_comp_signer = rsSI("Completed_Signer")
	            cur_sig_comp_sig_date = rsSI("Completed_Signature_Date") 
	            cur_sig_comp_signer_name = rsSI("Completed_Signer_Name") 
	            cur_sig_auth_type = rsSI("Authorized_Type")
	            cur_additional_signer_req_name = rsSI("additional_signer_req_name")
	     	    cur_additional_is_this_user = rsSI("additional_is_this_user")
	     	    
	     	    if cur_sig_comp_type = "Parent Guardian" THEN
            	    signed_message = "Signature " 
            	elseif cur_sig_comp_type = "Person Served" THEN
            	    signed_message = ""
            	else
            	    signed_message = "Signed By "
            	end if
	     	        
		        'GET ALL POSSIBLE/COMPLETED SIGS AND SPIT THEM OUT HERE
		        if cur_sig_req_type <> ""  THEN 

                

		            
		            if cur_sig_count = 0 THEN
		            %>
		                &nbsp;&nbsp;
		            <%
		            end if 
		            
		            if cur_sig_count > 0 THEN		            
	     		    %>
	     		    &nbsp;|&nbsp;
	     		    <%
	     		    end if
		        
		            if cur_sig_comp_signer <> "" THEN
		                'REQUIRED AND FILLED OUT BRANCH
		                %>
		                <span style="color:Black; font-weight:bold; font-size:14px;" onmouseover="showToolTip(event,'<%=signed_message %> <%=cur_sig_comp_signer_name %> on  <%=cur_sig_comp_sig_date %>')" onmouseout="hideToolTip()"><%if cur_sig_comp_type = "Additional" THEN %><%=cur_sig_comp_signer_name %><%else  %><%=cur_sig_comp_type %><%end if %></span>
		                <% 
		            else
		              
		                if cur_sig_auth_type <> "" and datalock_temp <> "" and (cur_sig_req_type <> "Additional" or cur_additional_is_this_user = "Yes") THEN%>
		                		                
		                <a href="javascript:void();" style="color:Blue; font-weight:bold; font-size:14px;" onclick="decideShowDiv('<%=cur_form_id%>','<%=cur_sig_comp_type %>');"><%if cur_sig_comp_type = "Additional" THEN %><%=cur_additional_signer_req_name %><%else  %><%=cur_sig_comp_type %><%end if %></a>
		                <% else %>	                    
		                    <span style="color:Gray; font-weight:bold; font-size:14px;" onmouseover="showToolTip(event,'Unsigned. Current user is NOT able to sign.')" onmouseout="hideToolTip()"><%if cur_sig_comp_type = "Additional" THEN %><%=cur_additional_signer_req_name %><%else  %><%=cur_sig_comp_type %><%end if %></span>		                    
		                <%end if %>
		            <%end if	
		            
		            cur_sig_count = cur_sig_count + 1			        
		        end if 		        
		            
		        rsSI.MoveNext
		    Loop        
		    %>&nbsp;&nbsp;
		    </td>
		    		    
		    <% end if %>
	      </tr>
	      <%end if    


	if cur_type = "PAPER" THEN%>

	   <tr>
	           <td border="0" width="100%" colspan="7" style="font-size: 7pt; font-family: verdana; font-weight: bold;" valign="bottom">
                   
                   
                    <div style="float:left;">
	          Documentation Program Filter:&nbsp;&nbsp;<select name="notes_filter" onchange="filterFormProgram(this.value);">
	                <option value="1000">Show All Programs</option>
	                <%
	                if Request.QueryString("cid") <> "" THEN
	                    Set SQLStmt2a = Server.CreateObject("ADODB.Command")
  	                    Set rs2a = Server.CreateObject ("ADODB.Recordset")

  	                    SQLStmt2a.CommandText = "exec [get_programs_for_staff_and_client] '" & Session("user_name") & "'," & Request.QueryString("cid")  
  	                    SQLStmt2a.CommandType = 1
  	                    Set SQLStmt2a.ActiveConnection = conn
  	                    SQLStmt2a.CommandTimeout = 45 'Timeout per Command
  	                    'response.write "SQL = " & SQLStmt2.CommandText
  	                    rs2a.Open SQLStmt2a
  	                    Do Until rs2a.EOF
      	                
  	                    cur_pid = rs2a("Program_ID")
	                    %>
	                    <option value="<%=cur_pid%>" <% if Int(cur_pid) = Int(qs_pid) THEN %>selected<% end if %> ><%Response.write rs2a("Program_Name")%></option>
	                    <%
	                    rs2a.MoveNext
                        Loop
                    end if
	                %>
	            </select> 
                        </div>

                    <!-- <div style="float:left;padding-left:20px;"><div style="float:left;padding-top:2px;"><input type="checkbox" name="record_view" value="1" <%if Request.QueryString("tf")="ALL" then response.write "checked" end if %>></div><label for="record_view" style="padding-left:3px;padding-bottom:12px;">Record View</label></div>-->
	          
	           <span style="float:right"><b>Signature Panel Key:</b> <font color="grey"><b>Grey=Not-Available</b></font>, <font color="blue"><b>Blue=Available</b></font>, <b>Black</b>=Signed&nbsp;</span>
	        </td>
	   </tr>
	        
	  <tr>
	    <td class="tabBarLeft"><img src="images/FormsList5.gif" width="349" height="18" border="0" alt="Form Name" title="Form Name" /></td>
		<td class="tabBarCenter"><img src="images/Status5.gif" width="72" height="18" border="0" alt="Status" title="Status" /></td>
		<td class="tabBarCenter" colspan="5"><img src="images/SignaturePanel5.gif" width="543" height="18" border="0" alt="Signatures" title="Signatures" /></td>
	  </tr>  
   <%end if
	      
    	  	
	  	    if cur_form_id <> "" and has_children = "Yes" THEN 
	        '---------------START INNER FORMS
	          inner_form_name = ""
	        cur_type = ""

             tf = Request.QueryString("tf")

            if cur_soft="DHSP" then
               tf = ""
            end if

           if Request.QueryString("utf")="DCF" or Request.QueryString("utf")="DHSWGDC" then
               tf = "All"
            end if
            Set SQLStmt30 = Server.CreateObject("ADODB.Command")
  	        Set rs30 = Server.CreateObject ("ADODB.Recordset")

  	           SQLStmt30.CommandText = "exec get_form_list_for_client2 " & client_id & ",'" & Session("user_name") & "'," & form_group_filter & "," & cur_form_id & ",'" & tf & "','" & Request.QueryString("ufid") & "','" & cur_soft & "'"
  	        SQLStmt30.CommandType = 1
  	        Set SQLStmt30.ActiveConnection = conn
  	        SQLStmt30.CommandTimeout = 45 'Timeout per Command
  	        if session("user_name")="elarochelle" then
              'response.write "SQL 1 = " & SQLStmt30.CommandText
            end if
  	        rs30.Open SQLStmt30
  	        Do Until rs30.EOF
      	          	        
  	        inner_blob_val = ""
  	        inner_create_date = ""
  	        inner_copied_form_id = ""
  	        inner_has_children = ""
  	        inner_copied_form_type = ""
  	        inner_copied_form_create_date = ""

  	        if rs30("Unique_Form_ID") THEN
      	         
  	            Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	            Set rs3 = Server.CreateObject ("ADODB.Recordset")
        
  	            SQLStmt3.CommandText = "exec get_form_info_without_content " & rs30("Unique_Form_ID")
  	            SQLStmt3.CommandType = 1
  	            Set SQLStmt3.ActiveConnection = conn
  	            SQLStmt3.CommandTimeout = 45 'Timeout per Command
  	            'if cur_form_id = "12854" THEN
      	        '    response.write "SQL = " & SQLStmt3.CommandText
  	            'end if 
  	            rs3.Open SQLStmt3

                inner_copied_form_id = rs3("Parent_Form_ID")
                inner_blob_val = rs3("File_Content")
                inner_create_date = rs3("Create_Date")  
                inner_copied_form_type = rs3("copied_form_type")
                inner_has_children = rs3("has_children")
                inner_copied_form_create_date = rs3("copied_form_create_date") 
                inner_main_program = rs3("Main_Program")
                
	        end if
    	    
    	    if inner_blob_val <> "" THEN
    	            
	            if inner_form_name <> rs30("FormName") THEN
	                inner_new_form_type = 1
	            else
	                inner_new_form_type = 0
	            end if 
        	    
	            inner_form_id = rs30("Unique_Form_ID")
	            inner_form_name = rs30("FormName")        
	            inner_type = rs30("Form_Type")
	            inner_parent_type = rs30("Parent_Form_Type")
	            inner_num_pages = rs30("Num_Pages")
	            inner_est_comp_time = rs30("Est_Completion_Time")
	            inner_linked_form_id = rs30("Linked_Form_ID")
	            inner_access = rs30("Access_Level")
	            inner_status = rs30("Status")
	            inner_create_user = rs30("Create_User")
	            inner_last_update = rs30("Update_Date")
	            inner_has_history = rs30("Has_History")
	            inner_has_attachments = rs30("Has_Attachments")
	            inner_has_notes = rs30("Has_Notes")
	            inner_has_billing_notes = rs30("Has_Billing_Notes")
	            
        	    inner_create_user_name = rs30("Create_User_Name")
	            inner_main_prog = rs30("Main_Program")
	            inner_main_staff = rs30("Main_Staff")
	            
	            inner_cur_dos = rs30("form_dos_date")
    	       	        	
	        	inner_hide_show_count = rs30("hide_show_count")
	           ' inner_external_system_id = rs30("external_system_id")     	
               
	            inner_display_status = ""
	            inner_display_date = ""
	        
	            'DETERMINE DISPLAY STATUS/DATE HERE
                if inner_status = "In-Process" and inner_type <> "HCEF" THEN
                    inner_display_status = inner_status
                    inner_display_date = inner_last_update
                else
                    if inner_cur_doa <> "" THEN
                        if inner_type = "BTMS" THEN
                            inner_display_status = "DOR"
                        else
                            inner_display_status = "DOA"
                        end if
                        inner_display_date = inner_cur_doa
                    elseif inner_cur_ntuc <> "" THEN
                        inner_display_status = "NTUC"
                        inner_display_date = inner_cur_ntuc                
                    elseif inner_cur_dos <> "" THEN
                        inner_display_status = "DOS"
                        inner_display_date = inner_cur_dos
                    end if
                end if

                 if inner_type = "HCEF" THEN
                     inner_display_status = "DOA"
                elseif inner_type = "DCF" or inner_type = "DHSWGDC" THEN
                     inner_display_status = "WOS"
               elseif inner_type = "DHMPN"  THEN
                     inner_display_status = "DOP"
                end if
                
                if inner_display_status = "" THEN
                    inner_display_status = "Finalized"
                    inner_display_date = inner_last_update
                end if	            
	                   	    
	            inner_form_hover = "This form has " & inner_num_pages & " page(s)"
	            
       
            if inner_type <> "QNR" then
       
               	    
	            if inner_new_form_type = 1 THEN
	            %>
	          <tr><td colspan="7" class="childDivider"></td></tr>
	          <tr>
	             <%if inner_type = "DCF" or inner_type = "DHSWGDC" or inner_type = "DHMPN" or inner_type = "ISSF" or inner_type = "MEDREV" or inner_type = "HCEF"  then %> 
	            <td colspan="7" class="formRowChild" valign="bottom"><div style="float:left;margin-left:90px;">
	            <%else %>
	             <td class="formRowChild" valign="bottom">
	            <%end if
	              if cur_type <> "PAPER" and cur_type <> "QNR" and tf<>"ALL" and inner_hide_show_count > 1 and ((Request.QueryString("ulfid") <> cstr(inner_linked_form_id)) or Request.QueryString("uft") <> inner_type) THEN %>
	        <a href="javascript:void();" onclick="undoFilterInner('<%=inner_form_id%>','<%=inner_linked_form_id %>','<%=inner_parent_type%>','<%=inner_type%>');" class="filterOff" style="text-decoration:none;" alt="Showing Filtered Forms"><font style="font-size:14px;" >+</font></a>
	        <% end if  %>
	            
	            <% if inner_type = "CBFSSN" THEN %>
	                <%if Request.QueryString("fn") = "all_cbfssn" THEN %>
	                    <%if inner_CBFSSN_total_count > 7 THEN %>
	                    <a href="#" onclick="filterNotes('less_cbfssn');" class="filterOn" alt="Showing All CBFS Service Notes">[see less]</a>
	                    <% end if %>
	                <%else %>
	                    <% if inner_CBFSSN_showing_count < inner_CBFSSN_total_count THEN %>
	                    <a href="#" onclick="filterNotes('all_cbfssn');" class="filterOff" alt="Showing Last <%=inner_CBFSSN_showing_count %> CBFS Service Notes">[see more]</a>
	                    <% end if %>
	                <%end if %>
    	            
	            <% elseif inner_type = "CBFSSNWE" THEN %>
	                <%if Request.QueryString("fn") = "all_cbfssnwe" THEN %>
	                    <a href="#" onclick="filterNotes('less_cbfssnwe');" class="filterOn" alt="Showing All CBFS Weekly Service Notes">[see less]</a>
	                <%else %>
	                    <% if inner_CBFSSNWE_showing_count < inner_CBFSSNWE_total_count THEN %>
	                    <a href="#" onclick="filterNotes('all_cbfssnwe');" class="filterOff" alt="Showing Most Recent CBFS Weekly Service Note Per Week">[see more]</a>
	                    <% end if %>
	                <%end if %>
	            <% elseif inner_type = "ESPLBN" THEN %>
	                <%if Request.QueryString("fn") = "all_log" THEN %>
	                    <%if inner_esp_log_bill_total_count > 7 THEN %>
	                    <a href="#" onclick="filterNotes('less_log');" class="filterOn" alt="Showing All ESP Log/Billing Notes">[see less]</a>
	                    <% end if %>
	                <%else %>
	                    <% if inner_esp_log_bill_showing_count < inner_esp_log_bill_total_count THEN %>
	                    <a href="#" onclick="filterNotes('all_log');" class="filterOff" alt="Showing Last <%=inner_esp_log_bill_showing_count %> ESP Log/Billing Notes">[see more]</a>
	                    <% end if %>
	                <%end if %>
	            <%else  %>
	
	               <img src="images/form_iconSmall.gif" border="0">
	              
	          <% end if %>
	            
	            <%if inner_access = "V" THEN%>
	                <font color="#777777"><%=inner_form_name%></font>
	            <%else%>
	                <%=inner_form_name%>
	            <%end if

                       if inner_type <> "DCF" and inner_type <> "DHSWGDC" and inner_type <> "DHMPN" and inner_type <> "ISSF" and inner_type <> "MEDREV" and inner_type <> "HCEF"   then%>
	            </div>
	            
	            <%end if %>

                </td>

                    <%if inner_type <> "DCF" and inner_type <> "DHSWGDC" and inner_type <> "DHMPN" and inner_type <> "ISSF" and inner_type <> "MEDREV" and inner_type <> "HCEF"  then %> 

		        <td  class="NoneStatus"><%if inner_blob_val = "" or inner_linked_form_id <> cur_form_id THEN %><b>NONE</b><%else %>&nbsp;<%end if %></td>

		        <td class="client">&nbsp;</td>
		        <td class="provider">&nbsp;</td>
		        <td class="guardian">&nbsp;</td>
		        <td class="md">&nbsp;</td>
		        <td class="supervisor">&nbsp;</td>
                   <%end if %>
	          </tr>
	          <%
	          inner_new_form_type = 0
	          end if 
        	  
	          if inner_blob_val <> "" and inner_linked_form_id = cur_form_id THEN
	          %>
	          <tr>
	            <td class="buttons">
                   

        <%if cur_staff_role_desc = "Administrator" or (inner_status = "In-Process" and (inner_access = "E" or inner_access = "L")) THEN 

       %>
            <% if inner_create_user = Session("user_name") or cur_staff_role_desc = "Administrator" THEN %>
            <a href="#" onclick="confirmFileDelete('<%=inner_form_name%>',<%=inner_form_id%>);"><img alt="Delete in-process form" src="images/delete_file.gif" border="0"/></a>
            <%end if %>
        <%end if %>

        <a href="choose_parent_for_copy.asp?width=700&height=300&uid=<%=inner_form_id%>&cid=<%=client_id%>&usid=<%=user_staff_id %>&ft=<%=inner_type %>&lfid=<%=inner_linked_form_id%>" title="FAQs" class="thickbox"><img src="images/MSDPCopyIcon.jpg" title="Make a copy of this form" border="0"></a>

        <a href="#" onclick="popupwindow('view_history.asp?uid=<%=inner_form_id%>',800,1000,'getHistoryWindow');"><img src="images/button_history.jpg"  border="0" alt="History" title="History" /></a>
        <a href="#" onclick="popupwindow('view_notes.asp?uid=<%=inner_form_id%>',800,1000,'getNotesWindow');">
        <%if inner_has_notes = "Yes" THEN %>
            <%if inner_has_billing_notes = "Yes" THEN %>
                <img src="images/button_notes_clip_dollar.jpg"  border="0" alt="Notes" title="Notes" />
            <%else %>
                <img src="images/button_notes_clip.jpg"  border="0" alt="Notes" title="Notes" />
            <%end if %>
        <%else %>
            <%if inner_has_billing_notes = "Yes" THEN %>
                <img src="images/button_notes_dollar.jpg"  border="0" alt="Notes" title="Notes" />
            <%else %>
                <img src="images/button_notes.jpg" border="0" alt="Notes" title="Notes" />
            <%end if %>
        <%end if %>
        </a>
        <%   if inner_type <> "DCF" and inner_type <> "DHSWGDC" then
            
            if inner_status = "Finalized" or inner_access = "V" THEN%>
            <% if inner_type ="DDS" or inner_type = "BTMS" THEN%>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                      <%elseif inner_type ="BEH_TRACK" THEN%>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1100,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                <% else %>
            <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1000,'<%=inner_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                <% end if %>
        <%elseif inner_status = "In-Process" and (inner_access = "E" or inner_access = "L") THEN%>
            <% if inner_type ="DDS" or inner_type = "BTMS" THEN%>
                <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                      <%elseif inner_type ="BEH_TRACK" THEN%>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1100,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                <% else %>
            <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1000,'<%=inner_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
            <% end if %>
        <%end if
            
            end if %>
        <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=inner_form_id%>&cid=<%=client_id%>',800,1000,'getAttachmentsWindow');"><%if inner_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg"  border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg"  border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>&nbsp; 
             <%if (inner_type = "DCF" or inner_type = "DHSWGDC") and inner_status = "Finalized" then 
          
                   Set SQLStmtDHMPN = Server.CreateObject("ADODB.Command")
                    Set rsDHMPN = Server.CreateObject ("ADODB.Recordset")                
                    SQLStmtDHMPN.CommandText = "select fxd.xml_data_section.value('(//sunday)[1]','datetime') as end_week_date from Form_XML_Data fxd where fxd.unique_form_id=" & inner_form_id
                    SQLStmtDHMPN.CommandType = 1 
                    Set SQLStmtDHMPN.ActiveConnection = conn
                ' response.Write "sql 1= " & SQLStmtDHMPN.CommandText
                    SQLStmtDHMPN.CommandTimeout = 45 'Timeout per Command
                    rsDHMPN.Open SQLStmtDHMPN
                    
                 '   end_week_date = ""
                    
                    Do Until rsDHMPN.EOF 
                      
                      end_week_date = rsDHMPN("end_week_date")
                    
                       rsDHMPN.MoveNext 
  	                 Loop
                    
                  ' d = CDATE(end_week_date)
                  
                    
                  '  response.Write(FormatDateTime(end_week_date,2))
                    
                   
                 
                         if CDate(FormatDateTime(end_week_date,2)) >= CDate(FormatDateTime(date(),2)) then
          
                       %> 
                        <a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy&uid=<%=inner_form_id%>&ft=<%=inner_type %>&cid=' + '<%=Request.QueryString("cid")%>' + '&pid=' + '<%=inner_main_program %>' + '&sid='+ '<%=Session("staff_id") %>' +'&dp=&lfid=<%=inner_linked_form_id %>&if=Fix',800,1000,'<%=inner_type %>_copy_' + '<%=Request.QueryString("cid")%>');"><img src="images/button_AddData.jpg"  border="0" title="Add Data to Form" alt="Add Data" /></a>&nbsp;
        
                      <%
                      else%>
                      <a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1000,'<%=inner_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a>
                      <%end if%>


          <%elseif (inner_type = "DCF" or inner_type = "DHSWGDC") and inner_status = "In-Process" then %>
                       <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?uid=<%=inner_form_id%>',800,1000,'<%=inner_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
        
         <% end if %>
              
       
        </td>

		        <td class="status">
		        <%if inner_has_the_needs = "1" THEN%><font color="#993300"><%=inner_display_status%></font>
		        <%elseif inner_status = "Finalized" THEN%><span class="finalized"><%=inner_display_status%></span>
		        <%else%><%=inner_display_status%>
		        <%end if%><br />
		        
		        <%if inner_copied_form_id <> "" and inner_copied_form_id <> "0" THEN %>
		    <div onmouseover="showToolTip(event,'Copied From <%=inner_copied_form_type%> with create date of <br / > <%=inner_copied_form_create_date%> <br/><br /> Created on <%=inner_create_date%> <br /> By <%=REPLACE(inner_create_user_name, "'", "\")%> <br /> Program: <%=inner_main_prog %> <br /> Staff: <%=REPLACE(inner_main_staff,"'", "\'") %>')" onmouseout="hideToolTip()">
		    <%else%>
		        <div onmouseover="showToolTip(event,'Created on <%=inner_create_date%> <br /> By <%=REPLACE(inner_create_user_name, "'", "\'")%> <br /> Program: <%=inner_main_prog %> <br /> Staff: <%=REPLACE(inner_main_staff, "'", "\'") %>')" onmouseout="hideToolTip()">
		    <%end if %>
		        <%if inner_has_the_needs = "1" THEN%><font color="#993300"><%=inner_display_date%></font>
		        <%elseif inner_status = "Finalized" THEN%><span class="finalized"><%=inner_display_date%></span>
		        <%else%><%=inner_display_date%>
		        <%end if%></div>
		        </td>
		        
		        <td class="current_sigs" colspan="5">
		        <%
		        Set SQLStmtSI = Server.CreateObject("ADODB.Command")
  	            Set rsSI = Server.CreateObject ("ADODB.Recordset")
        
  	            SQLStmtSI.CommandText = "exec get_signatures_completed_info_for_staff_and_form " & rs30("Unique_Form_ID") & ",'" & Session("user_name") & "'"
  	            SQLStmtSI.CommandType = 1
  	            Set SQLStmtSI.ActiveConnection = conn
  	            SQLStmtSI.CommandTimeout = 45 'Timeout per Command
  	     '     if Session("user_name") = "pwcard" THEN
         '        response.write "SQL2 = " & SQLStmtSI.CommandText
         '   end if 
  	            rsSI.Open SQLStmtSI
    		    
		        cur_sig_count = 0
		    
		    Do Until rsSI.EOF
		    
		        datalock_temp = rsSI("Completed_Hash")
	            cur_sig_req_type = rsSI("Required_Type")
	            cur_sig_comp_type = rsSI("Completed_Type")
	            cur_sig_comp_signer = rsSI("Completed_Signer")
	            cur_sig_comp_sig_date = rsSI("Completed_Signature_Date") 
	            cur_sig_comp_signer_name = rsSI("Completed_Signer_Name") 
	            cur_sig_auth_type = rsSI("Authorized_Type")
	            cur_additional_signer_req_name = rsSI("additional_signer_req_name")
	            cur_additional_is_this_user = rsSI("additional_is_this_user")
	     	  	
	     	  	if cur_sig_comp_type = "Parent Guardian" THEN
            	    signed_message = "Signature " 
            	elseif cur_sig_comp_type = "Person Served" THEN
            	    signed_message = ""
            	else
            	    signed_message = "Signed By "
            	end if
	     	  		    
		        'GET ALL POSSIBLE/COMPLETED SIGS AND SPIT THEM OUT HERE
		        if cur_sig_req_type <> "" and datalock_temp <> "" THEN 
		            
		            if cur_sig_count = 0 THEN
		            %>
		                &nbsp;&nbsp;
		            <%
		            end if
		            
		            if cur_sig_count > 0 THEN		            
	     		    %>
	     		    &nbsp;|&nbsp;
	     		    <%
	     		    end if
		        
		            if cur_sig_comp_signer <> "" THEN
		                'REQUIRED AND FILLED OUT BRANCH
		                %>
		                <span style="color:Black; font-weight:bold; font-size:14px;" onmouseover="showToolTip(event,'<%=signed_message %> <%=cur_sig_comp_signer_name %> on  <%=cur_sig_comp_sig_date %>')" onmouseout="hideToolTip()"><%if cur_sig_comp_type = "Additional" THEN %><%=cur_sig_comp_signer_name %><%else  %><%=cur_sig_comp_type %><%end if %></span>
		                <% 
		            else
		                'REQUIRED AND NOT FILLED OUT BRANCH
		                if cur_sig_auth_type <> "" and datalock_temp <> "" and (cur_sig_req_type <> "Additional" or cur_additional_is_this_user = "Yes") THEN%>
		                <a href="javascript:void();" style="color:Blue; font-weight:bold; font-size:14px;" onclick="decideShowDiv('<%=inner_form_id%>','<%=cur_sig_comp_type %>');"><%if cur_sig_comp_type = "Additional" THEN %><%=cur_additional_signer_req_name %><%else  %><%=cur_sig_comp_type %><%end if %></a>
		                <% else %>                    
		                    <span style="color:Gray; font-weight:bold; font-size:14px;" onmouseover="showToolTip(event,'Unsigned. Current user is NOT able to sign.')" onmouseout="hideToolTip()"><%if cur_sig_comp_type = "Additional" THEN %><%=cur_additional_signer_req_name %><%else  %><%=cur_sig_comp_type %><%end if %></span>		                    
		                <%end if %>
		            <%end if	
		            
		            cur_sig_count = cur_sig_count + 1			        
		        end if 		        
		            
		        rsSI.MoveNext
		    Loop        
		    %>&nbsp;&nbsp;
		        </td>    		
	          </tr>
	          <% end if 
	          end if
             end if
	            
	        rs30.MoveNext
            Loop        
            
	        '---------------END INNER FORMS
	        end if
	        last_cur_type = cur_type
	    end if

        form_group_counter = form_group_counter + 1
	  rs2.MoveNext
      Loop
      end if
      
      conn.Close()
      Set conn = Nothing
      %>           
	</table>
</div>
<p>&nbsp;</p>

<table id="bubble_tooltip" border="0" cellpadding="0" cellspacing="0">
        <tr class="bubble_top"><td></td></tr>
        <tr class="bubble_middle"><td><div id="bubble_tooltip_content"></div></td></tr>
        <tr class="bubble_bottom"><td></td></tr>
</table>
</body>
</html>
