<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
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
<script type="text/javascript" src="includes/scrollsave.js"></script>



<style type="text/css" >
    .ui-dialog-titlebar {
  background-color: #BECEE0;
  background-image: none;
  color: #000;

}

    .ui-dialog .ui-dialog-title {
      width: 100%;
      text-align: center;
}

    </style>
<link rel="stylesheet" href="includes/thickbox.css" type="text/css" media="screen" />
<%




     Set SQLStmt2 = Server.CreateObject("ADODB.Command")
    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	SQLStmt2.CommandText =  "exec get_staff_info_by_username '" & Session("user_name") & "'"
  	SQLStmt2.CommandType = 1
  	Set SQLStmt2.ActiveConnection = conn
  	SQLStmt2.CommandTimeout = 45 'Timeout per Command
  '	response.write "SQL = " & SQLStmt2.CommandText
  	rs2.Open SQLStmt2
  	
    cur_staff_id = rs2("staff_id")
  	emr_staff = rs2("emr_staff")
    agency_program_id = rs2("agency_program_id")
    agency_contact = rs2("agency_contact")

      if cur_staff_id=7 or cur_staff_id=9 or cur_staff_id=1335 or cur_staff_id=1336 then
    agency_program_id =-1
    end if



'*****CALC THESE AFTER*****'
cur_month = Month(Date())
cur_year = Year(Date())
   
    if Request.QueryString("sid") <> "" THEN
    
        staff_id = Request.QueryString("sid")
    
        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	    Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	    SQLStmt2.CommandText = "get_staff_info " & Request.QueryString("sid") 
  	    SQLStmt2.CommandType = 1
  	    Set SQLStmt2.ActiveConnection = conn
  	    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	  '  response.write "SQL = " & SQLStmt2.CommandText
  	    rs2.Open SQLStmt2  	
  	    
  	    choosen_staff_id = rs2("staff_id")
  	    choosen_staff_name = rs2("Staff_Name")
  	    choosen_staff_creds = rs2("Position")
  	    choosen_staff_role_name = rs2("Role_Description")
  	    choosen_staff_username = rs2("user_name")
  	    choosen_staff_external = rs2("external_info")
  	    choosen_map_cert = rs2("map_cert")
  	    staff_has_picture = rs2("has_picture")
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
    
    'GET CLIENT COUNT FOR THIS STAFF MEMBER    
	Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	SQLStmt2.CommandText = "exec [get_staff_header_info] '" & Session("user_name") & "'"  
  	SQLStmt2.CommandType = 1
  	Set SQLStmt2.ActiveConnection = conn
  	SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	rs2.Open SQLStmt2
  	Do Until rs2.EOF
	        client_count = client_count + 1
	        total_client_forms = total_client_forms + rs2("client_form_count")
	rs2.MoveNext
    Loop
    
    total_alerts_for_header = 0
    total_calls_for_header = 0
    
    client_id_for_alerts = -1
    
    if Request.QueryString("cid") <> "" THEN
        client_id_for_alerts = Request.QueryString("cid")
    end if
    
%>
<script type="text/javascript">
    function confirmFileDelete(fname, fid)
	{
	    if(confirm("Are you sure you want to delete this " + fname + "?"))
		{
			var url = "delete_staff_form.asp?sid="+'<%=request.queryString("sid")%>'+"&fid="+ fid;
			 window.location.href = url;
		}
	}
	
	function clickToClearAlert(alert_id, call_log_id)
	{
	    if(confirm("Are you sure you want to clear this system message?"))
		{
			var url = "click_to_clear_alert.asp?aid="+ alert_id + "&call_id=" + call_log_id + "&cid=" + '<%=Request.QueryString("cid")%>';
			navigate(url);
			return true;
		}
		else
		{
		    return false;
		}
	}
	
    function refreshProg(newProg)
    {    
        window.location.href = "index_staff.asp?pCode=" + newProg;
    }
    function refreshLocation(newLocation)
    {    
        window.location.href = "index_staff.asp?pCode=" + document.clientForm.program.value + "&lCode=" + newLocation;
    }
    
    function refreshClient(newClient)
    {
        Ajax('create_form_div','&nbsp;'); 
        
        window.location.href = "index_staff.asp?cid=" + newClient + "&pCode=" + document.clientForm.program.value + "&lCode=" + document.clientForm.location.value + "&sd=" + clientForm.show_discharged.checked;
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
        
        filterCreatableForms();     
                          
    }
    
    function filterStaffFormSD() 
    {
        

         window.location.href = "index_staff.asp?sid=" + frmMain.form_staff.value + "&sd=" + frmMain.show_discharged.checked + "&csid=" + '<%=user_staff_id %>';

      
    }
    
    function filterStaffForms(val)
    {    
        window.location.href = "index_staff.asp?sid=" + val + "&sd=" + frmMain.show_discharged.checked + "&csid=" + '<%=user_staff_id %>';
      
    }
    
    function filterCreatableForms()
    {
        if(frmMain.form_program.value != '')
        {
            Ajax('create_form_div','scripts/ShowCreatableForms.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&cage=<%=client_age %>'); 
        }
        else if(frmMain.service_appointment)
        {
            Ajax('create_form_div','scripts/ShowCreatableForms.asp?cid=<%=Request.QueryString("cid")%>&service_id='+frmMain.service_appointment.value+'&cage=<%=client_age %>'); 
        }
    }

    function checkFormCopyPrevious() {
        if (frmMain.create_form_type.value == '') {
            alert("You must choose a form type in order to copy it from the previous episode");
        }
        else {
            Ajax('possible_parent_forms', 'scripts/ShowPossibleFormsForCopy.asp?cid=<%=Request.QueryString("cid")%>&ft=' + frmMain.create_form_type.value);
        }
    }
    
    function reloadFormParents(formValue)
    {
        Ajax('possible_parent_forms','scripts/ShowPossibleParentsForStaffForm.asp?sid=<%=Request.QueryString("sid")%>&sft=' + formValue);
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

    function updateServiceTypeColor(serviceType) 
    {
        if (serviceType == 'Scheduled') 
        {
            var thestyle= eval ('document.all.sched_td_1.style');
            thestyle.backgroundColor = '#99CCCC';
            
            thestyle= eval ('document.all.sched_td_2.style');
            thestyle.backgroundColor = '#99CCCC';
            
            thestyle= eval ('document.all.unsched_td_1.style');
            thestyle.backgroundColor='';
            
            thestyle= eval ('document.all.unsched_td_2.style');
            thestyle.backgroundColor='';
            
            //turn off and clear the inputs for unscheduled section here
            document.frmMain.form_program.value = "";
            document.frmMain.form_program.disabled = true;
            if(document.getElementById("form_staff"))
            {
                document.frmMain.form_staff.value = "";
                document.frmMain.form_staff.disabled = true;
            }
            document.frmMain.program_default.checked = false;
            document.frmMain.program_default.disabled = true;
            
            //turn on scheduled section here
            if(document.getElementById("service_appointment"))
            {
                document.frmMain.service_appointment.disabled = false;
            }
        }
        else 
        {
            var thestyle= eval ('document.all.sched_td_1.style');
            thestyle.backgroundColor='';
            
            thestyle= eval ('document.all.sched_td_2.style');
            thestyle.backgroundColor='';
            
            thestyle= eval ('document.all.unsched_td_1.style');
            thestyle.backgroundColor = '#B0E0E6';
            
            thestyle= eval ('document.all.unsched_td_2.style');
            thestyle.backgroundColor = '#B0E0E6';
            
            //turn off the inputs for unscheduled section here
            if(document.getElementById("service_appointment"))
            {
                document.frmMain.service_appointment.value = "";
                document.frmMain.service_appointment.disabled = true;
            }
        
            //turn on scheduled section here
            document.frmMain.form_program.disabled = false;
            if(document.getElementById("form_staff"))
            {
                document.frmMain.form_staff.disabled = false;
            }
            document.frmMain.program_default.disabled = false;
            
            if(document.frmMain.form_program.value != "")
            {
                Ajax('create_form_div','scripts/ShowCreatableForms.asp?cid=<%=Request.QueryString("cid")%>&pid='+frmMain.form_program.value+'&cage=<%=client_age %>'); 
            }
        }        
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
        if(enableCache && jsCache[url]){
	     document.getElementById(divId).innerHTML = jsCache[url];
	     return;
	    }	
	    
	    var ajaxIndex = AjaxObjects.length;
	    document.getElementById(divId).innerHTML = '<img src=images/movewait.gif width=16 height=16 hspace=10 vspace=10 />';
	    AjaxObjects[ajaxIndex] = new sack();
	    AjaxObjects[ajaxIndex].requestFile = url;
	    AjaxObjects[ajaxIndex].onCompletion = function(){ ShowContent(divId,ajaxIndex,url); };
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
                 
         if(frmMain.form_staff.value == '' && getSelectedRadioValue(document.frmMain.service_type) == "Unscheduled")
         {
            alert("Please select a Staff before creating a new form.");
                               
            return false;
         }
         
         if(frmMain.create_form_type.value == '')
         {
            alert("Please select a Form Type before creating a new form.");
            return false;
         }
                     
        var chosenParentValue = "";
        var parentFormRequired = 0;
        var chosenSearchValue = "";
        var chosenOpenDateValue = "";
        var chosenClosedDateValue = "";
        var chosenStatusValue = "";
     


         if(document.getElementById("tags"))
        {        
           
                
            chosenSearchValue = document.frmMain.tags.value;
        }

      if(document.getElementById("open_date"))
        {        
           
                
            chosenOpenDateValue = document.frmMain.open_date.value;
        }

      if(document.getElementById("status"))
        {        
           
                
            chosenStatusValue = document.frmMain.status.value;
        }

        if(document.getElementById("closed_date"))
        {        
           
                
            chosenClosedDateValue = document.frmMain.closed_date.value;
        }
    
        if(document.getElementById("parent_form_id"))
        {        
            parentFormRequired = 1;
                
            chosenParentValue = getSelectedRadioValue(document.frmMain.parent_form_id);
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
                if (frmMain.create_form_type.value == "CALL_LOG" || frmMain.create_form_type.value == "GIAP" || frmMain.create_form_type.value == "KHTP" || frmMain.create_form_type.value == "LRA" || frmMain.create_form_type.value == "DMTC" || frmMain.create_form_type.value == "AF")
                {
                    popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft='+ frmMain.create_form_type.value + '&sid=' + frmMain.form_staff.value,800,1360,frmMain.create_form_type.value + '' + frmMain.form_staff.value);
                } else if (frmMain.create_form_type.value == "STAFF_TRACK_REPORT") {

              
                   popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft='+ frmMain.create_form_type.value + '&sid=' + frmMain.form_staff.value + '&search=' + chosenSearchValue.replace(/\s/g, "^") + '&od=' + chosenOpenDateValue + '&cd=' +  chosenClosedDateValue + '&status=' + chosenStatusValue,800,1000,frmMain.create_form_type.value + '' + frmMain.form_staff.value);
               

               } else
                {
                   popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft='+ frmMain.create_form_type.value + '&sid=' + frmMain.form_staff.value,800,1000,frmMain.create_form_type.value + '' + frmMain.form_staff.value);
                }                
            }            
        }
        else
        {
            if (frmMain.create_form_type.value == "CALL_LOG" || frmMain.create_form_type.value == "GIAP" || frmMain.create_form_type.value == "KHTP" || frmMain.create_form_type.value == "LRA" || frmMain.create_form_type.value == "DMTC" || frmMain.create_form_type.value == "AF")
            {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft='+ frmMain.create_form_type.value + '&sid=' + frmMain.form_staff.value + '&lfid=' + chosenParentValue,800,1400,frmMain.create_form_type.value + '' + frmMain.form_staff.value);
            }
            else
            {
                popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft='+ frmMain.create_form_type.value + '&sid=' + frmMain.form_staff.value + '&lfid=' + chosenParentValue,800,1000,frmMain.create_form_type.value + '' + frmMain.form_staff.value);                  
            }
            
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
	    window.location.href= "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&lCode=" + document.clientForm.location.value + "&fb=" + filterName + "&sb=" + '<%=Request.QueryString("sb")%>';
	}
	
	function undoFilter(undo_form_id)
	{
	    window.location.href= "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&foid=" + undo_form_id + '&tf=<%=Request.QueryString("tf")%>';
	}	
	
	function filterNotes(filterName)
	{
	    window.location.href= "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&lCode=" + document.clientForm.location.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + filterName;
	}
	
	function filterTimeFrame(filterName) 
	{
	    window.location.href = "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&tf=" + filterName;
	}
	
	function filterEpisodes(filterName)
	{
	    window.location.href= "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&lCode=" + document.clientForm.location.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + '<%=Request.QueryString("fn")%>' + "&episodes=" + filterName;
	}
	
	function filterFormGroup(filterName)
	{
	    window.location.href= "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + '<%=Request.QueryString("sb")%>' + "&fn=" + filterName + "&episodes=" + '<%=Request.QueryString("episodes")%>';
	}
	
	function filterFormProgram(filterName)
	{

      if($("#assigned_to_me").prop('checked')) {
           assigned_to_me=1
         } else {
          assigned_to_me=0
        }


	    window.location.href= "index_staff.asp?pCode=" + filterName + "&sid=" + '<%=Request.QueryString("sid")%>' + "&sd=" + '<%=Request.QueryString("sd")%>' + "&csid=" + '<%=cur_staff_id%>' + "&status=" + $("#form_status").val() + "&atm=" + assigned_to_me;
	}

    function filterFormStatus(filterName)
	{

       
        if($("#assigned_to_me").prop('checked')) {
           assigned_to_me=1
         } else {
          assigned_to_me=0
        }

	    window.location.href= "index_staff.asp?status=" + filterName + "&sid=" + '<%=Request.QueryString("sid")%>' + "&sd=" + '<%=Request.QueryString("sd")%>' + "&csid=" + '<%=cur_staff_id%>' + "&pCode=" + $("#notes_filter").val() + "&atm=" + assigned_to_me;
	}
	
	function sortList(sortName)
	{
	    window.location.href= "index_staff.asp?cid=" + '<%=Request.QueryString("cid")%>' + "&pCode=" + document.clientForm.program.value + "&lCode=" + document.clientForm.location.value + "&fb=" + '<%=Request.QueryString("fb")%>' + "&sb=" + sortName;
	}
		
	function decideShowDiv(formID, sigType)
	{
	    if(sigType == 'Provider' || sigType == 'MD' || sigType == 'Supervisor' || sigType == 'Additional' || sigType == 'Family Partner' || sigType == 'Family Partner Supervisor' || sigType == 'MDT 1' || sigType == 'MDT 2' || sigType == 'MDT 3' || sigType == 'MDT 4')
	    {
	        showdiv3(formID,sigType);
	    }
	    if(sigType == 'Parent Guardian')
	    {
	        showdiv5(formID,sigType);
	    }
	    if(sigType == 'Person Served')
	    {
	        showdiv4(formID,sigType);
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

function showToolTip(e,text){
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
                popupwindow('edit_client.asp?from_home=1&cid=' + document.clientForm.client_id.value,600,1100,'editClientWindow');
            }
	    }
	    function assignClient()
	    {
	        if("" == document.clientForm.client_id.value)
	        {
	            alert('Please select a client to assign');
	        }
	        else
	        {
	            popupwindow('client_program_assign.asp?from_home=1&cid=' + document.clientForm.client_id.value,600,1100,'assignClientWindow');
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

/*--IE 6 PNG Fix--
img{ behavior: url(iepngfix.htc) }
*/
</style>

</head>

<body>
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
    <div id="dialogDiv" style="display:none;">
      
      <div><iframe src="http://emrhelpdesk.massemr.org/sylli.asp?sid=<%=cur_staff_id %>&nocache=<%=Replace(Now(),":","") %>" frameborder="0" width="996" height="800" scrolling="auto"></iframe></div>
    </div>
<script>


   

function updatePhrase() {
    if (signFormClient.client_sig_type.value == 'Pass Phrase') {
        signFormClient.client_pass_phrase.disabled = false;
        if (document.getElementById("phrase_hint")) {
            signFormClient.phrase_hint.disabled = false;
        }
    } else {
        signFormClient.client_pass_phrase.disabled = true;
        if (document.getElementById("phrase_hint")) {
            signFormClient.phrase_hint.disabled = true;
        }
    }
}

function checkClientSign() {
    if (signFormClient.client_sig_type.value == 'Pass Phrase' && (signFormClient.client_pass_phrase.value == '' || signFormClient.client_pass_phrase.value == ' ')) {
        alert("Please enter a pass phrase if choosing 'Pass Phrase' as the signature type.");
        return false;
    }
}

$(document).ready(function() {


  


     $("#assigned_to_me").change(function () {
   

       var include_closed=0;

         if ($("#include_closed").prop('checked')) {

             include_closed=1;
         }


    
      if ($(this).prop('checked')) {
       window.location.href= "index_staff.asp?status=" + $('#form_status').val() + "&sid=" + '<%=Request.QueryString("sid")%>' + "&sd=" + '<%=Request.QueryString("sd")%>' + "&csid=" + '<%=cur_staff_id%>' + "&pCode=" + $("#notes_filter").val() + "&atm=1" + "&ic=" + include_closed;
        } else {
        window.location.href= "index_staff.asp?status=" + $('#form_status').val() + "&sid=" + '<%=Request.QueryString("sid")%>' + "&sd=" + '<%=Request.QueryString("sd")%>' + "&csid=" + '<%=cur_staff_id%>' + "&pCode=" + $("#notes_filter").val() + "&atm=0" + "&ic=" + include_closed;
        }
    
    });



      $("#include_closed").change(function () {
   
    
         var assigned_to_me=0;

         if ($("#assigned_to_me").prop('checked')) {

             assigned_to_me=1;
         }


    

      if ($(this).prop('checked')) {
       window.location.href= "index_staff.asp?status=" + $('#form_status').val() + "&sid=" + '<%=Request.QueryString("sid")%>' + "&sd=" + '<%=Request.QueryString("sd")%>' + "&csid=" + '<%=cur_staff_id%>' + "&pCode=" + $("#notes_filter").val() + "&atm=" + assigned_to_me + "&ic=1";
        } else {
        window.location.href= "index_staff.asp?status=" + $('#form_status').val() + "&sid=" + '<%=Request.QueryString("sid")%>' + "&sd=" + '<%=Request.QueryString("sd")%>' + "&csid=" + '<%=cur_staff_id%>' + "&pCode=" + $("#notes_filter").val() + "&atm=" + assigned_to_me + "&ic=0";
        }
    
    });



     $("#autocomplete_placeholder").html("");

        var autoCompleteInput = "<input id='tags' style='width:285px'>";

                var $autoCompleteInput = $(autoCompleteInput);
                //    $autoCompleteInput.click(function() { alert('a'); });



                  <%
             Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	         Set rs3 = Server.CreateObject ("ADODB.Recordset")
                        
  	         SQLStmt3.CommandText = "exec get_helpdesk_search "
  	         SQLStmt3.CommandType = 1
  	         Set SQLStmt3.ActiveConnection = conn
  	         'response.write "SQL = " & SQLStmt3.CommandText
  	         rs3.Open SQLStmt3
                	
          
            search_string = ""
            search_cnt = 0
            search = ""
            cnt = ""
             search_string =  "  var availableTags = ['"
             search_submissions_disable ="disabled='disabled'"  
	         Do Until rs3.EOF

                 search_cnt =  search_cnt + 1

               search =  Replace(rs3("search"),"'","\'")
                cnt = rs3("cnt")           	  
      
	          if cnt <> search_cnt then
                   search_string = search_string & search & "','"
             else
                search_string = search_string & search & "'"
             end if

               search_submissions_disable =""  
                
	            rs3.MoveNext
                	        
	         Loop 
    
              search_string = search_string & "];"

           if search_cnt = 0 then
             search_string = " var availableTags = ''"
           end if
    
           response.write search_string
    %>    

               
              




                $autoCompleteInput.autocomplete({
                        source: availableTags

                });




                $("#autocomplete_placeholder").append($autoCompleteInput);















    var is_emr = '<%=Request.QueryString("sid") %>';

    if (is_emr == "1344") {
        //  $("#create_form_label").hide();
        //  $("#create_form_div").hide();
        $("#help_desk_form").show();

  
         $("#search_submissions").show();
     $("#emr_label").show();   
     $("#emr_divider").show();
 
      $("#create_form_label").hide();
    $("#create_form_div").hide();
    
    
    } else {   

        // $("#create_form_label").show();
        //  $("#create_form_div").show();
        $("#help_desk_form").hide();
      $("#search_submissions").hide();
    $("#emr_label").hide();
     $("#emr_divider").hide();
     $("#create_form_label").show();
     $("#create_form_div").show();
   
    }



    var url = "scripts/FilterStaffList.asp?csid=" + '<%=user_staff_id %>' + "&sd=" + '<%=Request.QueryString("sd") %>' + "&sid=" + '<%=Request.QueryString("sid") %>'

    $.ajax({
        url: url,
        cache: false,
        success: function(result) {

            $("#form_staff_placeholder").html(result);
            //  alert($("#form_staff").find(":selected").text());
            //  alert($("#form_staff").find(":selected").text());
        }
    });



    $("#help_desk_form").click(function() {
        popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft=STAFF_EMR_HELPDESK&sid=' + frmMain.form_staff.value + '&csid=' + '<%=cur_staff_id %>', 800, 1000, 'STAFF_EMR_HELPDESK');
        return false;

    });

    


    $("#sylli_link").click(function() {
    
   $("#dialogDiv").dialog({
          
            width: 1050,
            height: 900,
            title: "Since Your Last Log In Dashboard",
            autoOpen: false,
            modal : true,
           open: function( event, ui ) {
    
         

            }});

        

        


     $("#dialogDiv").dialog("open");



    return false;
    });



    $("#search_submissions").click(function() {

   

       // alert($("#search_submissions").attr("disabled"));


    var attr = $(this).attr("disabled");

      if (typeof attr !== typeof undefined && attr !== false) {
        return false;
    }



    var chosenSearchValue = $("#tags").val();


 
   
     if($("#search_assigned_to_me").prop('checked')) {
      search_assigned_to_me=1;

    } else {
       search_assigned_to_me=0;
    }



      if($("#search_include_closed").prop('checked')) {

   
      search_include_closed=1;

    } else {
       search_include_closed=0;
    }


       popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&ft=STAFF_TRACK_REPORT&sid=<%=cur_staff_id %>&search=' +   chosenSearchValue.replace(/\s/g, "^") + '&od=' +  $("#open_date").val() + '&cd=' +  $("#closed_date").val() + '&status=' + $("#status").val() + '&submission_type=' + encodeURIComponent($("#submission_type").val()) + '&agency=' + $("#agency").val() + '&notify=' + $("#notify").val() + '&aging=' + $("#aging").val() + "&atm=" + search_assigned_to_me + "&ic=" + search_include_closed,800,1000,'STAFF_TRACK_REPORT');
               


      return false;

    });




});
      



</script>
    
  
   


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
<form name="frmMain" action="index_staff.asp" method="post">

   <div id="container">
	<table cellpadding="0" cellspacing="0" border="0" width="100%">
	  <tr>
	    <td align="left" valign="bottom"><font style="font-weight:bold; font-size:22px; color:#dc7d18;">&nbsp;STAFF</font>
	    </td>
	    <td align="right">  
	        <!--#include file="nav_menu.asp" -->
            <script language="javascript"> 
            <!--
            buildsubmenus_horizontal();
            //-->
            </script>
        </td>
  </tr>  
</table>
   <div id="spacer5px"></div>
   	
	<div class="pod_add_new_form" align="left">
    
    <span style="float:left;"><b>STAFF NAME:</b>&nbsp;&nbsp;</span>
        <div style="float:left;" id="form_staff_placeholder"></div>&nbsp;
	       <div style="float:left;padding-left:10px;"><b>|</b>&nbsp;</div>
              <div style="float:left;padding-left:5px;" >
                <% if Session("active_staff_manage_staff_auth") = "1" THEN %><img src="/images/icon_add_person.png" width="7" height="7" border="0" alt="Add Staff" title="Add Staff" style="padding-right: 4px;" /><a href="#" onclick="popupwindow('manage_staff.asp?',600,1300,'manageStaffWindow');" class="addPerson">ADD NEW STAFF</a>
                <%end if %>
         </div>
         <div style="clear:left;" ></div>
         <div style="float:left;padding-left:86px;padding-top:7px;padding-bottom:7px;"><input type="checkbox" class="radio" name="show_discharged" onclick="filterStaffFormSD()" <%if Request.QueryString("sd") = "true" THEN %>checked<%end if %> />  Show Inactive </div>
	  <table cellpadding="3" cellspacing="3" border="0" width="100%" class="pod_selected_client">
	    <tr>
	        
	        <td valign="middle" align="left" style="padding-left: 2px;" bgcolor="#E4EDF6" width="124">
		      <table cellpadding="0" cellspacing="0" width="100" height="115" class="pod">
		        <tr>
			      <td align="center">
			      <% if staff_has_picture = "Yes" THEN %>
			        <a href="#" onclick="popupwindow('display_orig_image.asp?sid=<%=Request.QueryString("sid")%>',800,1000,'manageStaffWindow');"><img src="display_image.asp?sid=<%=Request.QueryString("sid")%>" width="95" height="120" border="0" /></a>
			      <% else %>  
			      Person Photo
			      <% end if %>
			      </td>
			    </tr>
		      </table>
		    </td>
		    
	        <td align="left" colspan="2" bgcolor="#E4EDF6" class="blackFontSmall" style="font-weight:normal;" valign="top">
                  <%if Request.QueryString("sid") = "1344" then %>
                   <div style="float:left;padding-top:5px;"><span id="emr_label"><b>EMR HELPDESK:&nbsp;</b></span></div>
                   <div style="float:left;padding-top:5px;"><a href="" id="sylli_link"><b>[SYLLI Dashboard]</b></a><br /></div><br />

<table cellpadding="3" cellspacing="1" border="0" width="100%" class="pod_selected_client">
	    <tr  bgcolor="#d0d8e5">
              <td ><b>Note Status</td>
              <td><b>0-7 Days</td>
              <td><b>8-15 Days</td>
              <td><b>16-30 Days</td>
              <td><b>30 + Days</td>
              <td><b>TOTAL</td>
           </tr>

    <%
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "exec get_note_status_aging '" & agency_program_id & "'"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	      '  response.write SQLStmtV.CommandText
  	            Do Until rsV.EOF
  	                open1to7 = rsV("open1to7")
                    open8to15 = rsV("open8to15")
                    open15to30 = rsV("open15to30")
                    open30plus = rsV("open30plus")
                    open_total = rsV("open_total")

                      assigned1to7 = rsV("assigned1to7")
                    assigned8to15 = rsV("assigned8to15")
                    assigned15to30 = rsV("assigned15to30")
                    assigned30plus = rsV("assigned30plus")
                    assigned_total = rsV("assigned_total")

                     hold1to7 = rsV("hold1to7")
                    hold8to15 = rsV("hold8to15")
                    hold15to30 = rsV("hold15to30")
                    hold30plus = rsV("hold30plus")
                    hold_total = rsV("hold_total")

                      future1to7 = rsV("future1to7")
                    future8to15 = rsV("future8to15")
                    future15to30 = rsV("future15to30")
                    future30plus = rsV("future30plus")
                    future_total = rsV("future_total")

                     next1to7 = rsV("next1to7")
                    next8to15 = rsV("next8to15")
                    next15to30 = rsV("next15to30")
                    next30plus = rsV("next30plus")
                    next_total = rsV("next_total")

                     total1to7 = rsV("total1to7")
                    total8to15 = rsV("total8to15")
                    total15to30 = rsV("total15to30")
                    total30plus = rsV("total30plus")
                    total_total = rsV("total_total")
  	               
                    total_closed = rsV("total_closed")
  	            
  	            
  	            rsV.MoveNext
  	            Loop
  	            
        
        
         %>




	    <tr>
              <td>OPEN</td>
              <td align="center" <%if open1to7 > 0 then response.write "style='font-weight: bold;background-color:#E3E763'" end if  %>><%=open1to7 %></td>
              <td align="center" <%if open8to15 > 0 then response.write "style='font-weight: bold;background-color:#F5ACAE'" end if  %>><%=open8to15 %></td>
              <td align="center" <%if open15to30 > 0 then response.write "style='font-weight: bold;background-color:#F5ACAE'" end if  %> ><%=open15to30 %></td>
              <td align="center" <%if open30plus > 0 then response.write "style='font-weight: bold;background-color:#F5ACAE'" end if  %>><%=open30plus %></td>
              <td align="center"><%=open_total %></td>
           </tr>

    


	    <tr>
              <td>ASSIGNED</td>
              <td align="center"><%=assigned1to7 %></td>
              <td align="center" <%if assigned8to15 > 0 then response.write "style='font-weight: bold;background-color:#E3E763'" end if  %>><%=assigned8to15 %></td>
              <td align="center" <%if assigned15to30 > 0 then response.write "style='font-weight: bold;background-color:#F5ACAE'" end if  %>><%=assigned15to30 %></td>
              <td align="center" <%if assigned30plus > 0 then response.write "style='font-weight: bold;background-color:#F5ACAE'" end if  %>><%=assigned30plus %></td>
              <td align="center"><%=assigned_total %></td>
           </tr>
	    <tr>
              <td>NEXT</td>
              <td align="center"><%=next1to7 %></td>
              <td align="center" <%if next8to15 > 0 then response.write "style='font-weight: bold'" end if  %>><%=next8to15 %></td>
              <td align="center" <%if next15to30 > 0 then response.write "style='font-weight: bold;background-color:#E3E763'" end if  %>><%=next15to30 %></td>
              <td align="center" <%if next30plus > 0 then response.write "style='font-weight: bold;background-color:#E3E763'" end if  %>><%=next30plus %></td>
              <td align="center"><%=next_total %></td>
           </tr>
	    <tr>
              <td>FUTURE</td>
              <td align="center"><%=future1to7 %></td>
              <td align="center" <%if future8to15 > 0 then response.write "style='font-weight: bold'" end if  %>><%=future8to15 %></td>
              <td align="center" <%if future15to30 > 0 then response.write "style='font-weight: bold'" end if  %>><%=future15to30 %></td>
              <td align="center" <%if future30plus > 0 then response.write "style='font-weight: bold;background-color:#E3E763'" end if  %>><%=future30plus %></td>
              <td align="center"><%=future_total %></td>
           </tr>

	    <tr>
              <td>HOLD</td>
              <td align="center"><%=hold1to7 %></td>
              <td align="center" <%if hold8to15 > 0 then response.write "style='font-weight: bold'" end if  %>><%=hold8to15 %></td>
              <td align="center" <%if hold15to30 > 0 then response.write "style='font-weight: bold'" end if  %>><%=hold15to30 %></td>
              <td align="center" <%if hold30plus > 0 then response.write "style='font-weight: bold;background-color:#E3E763'" end if  %>><%=hold30plus %></td>
              <td align="center"><%=hold_total %></td>
           </tr>
	    <tr>
              <td align="right"><b>TOTAL:</td>
               <td align="center"><%=total1to7 %></td>
              <td align="center"><%=total8to15 %></td>
              <td align="center"><%=total15to30 %></td>
              <td align="center"><%=total30plus %></td>
              <td align="center"><b><%=total_total %></b></td>
           </tr>


</table>

                <div style="float:left;"><b>CLOSED COUNT:<%=total_closed %></b></div>

                <div style="float:right;padding-left:10px;"><a id="help_desk_form" href="" target="_blank"><div style="float:left;"><img src="images/help_desk.jpg" border="0" alt="New iCentrix Change Tracking" title="New iCentrix Change Tracking"></div><span style="float:left;padding-top:7px;padding-left:5px;"><b>[NEW ENTRY]</b></span></a></div>                  
<% else %>  
                    <b><u>Staff Details</u></b>
	            Credentials: <b><%=choosen_staff_creds %> </b><br />
	            Role: <b><%=choosen_staff_role_name %> </b><br />
	            User Name: <b><%=choosen_staff_username %> </b><br />
	            External Staff ID: <b><%=choosen_staff_external %> </b><br />
	            MAP Certified: <b><%=choosen_map_cert %></b><br /><br />
<%end if %>

                    
                
                   
                  
	        </td>	        
	        
	      
                 <td width="420px;" colspan="2" bgcolor="#E4EDF6" class="blackFontSmall">
                
              <%if Request.QueryString("sid") = "1344" then %>

                
             <div style="float:left"><div style="padding-top:2px;"><b>Search:</b></div></div>
                          <div style="float:left;padding-left:10px;padding-bottom:5px;"><div id="autocomplete_placeholder"></div></div>

                    
              
              

                <div style="clear:left"></div>
                    
                  
      	        
                  <div style="float:left;"><div style="padding-top:2px;"><b>Submission Type:</b></div></div>
      	           <div style="float:left;padding-left:7px;padding-bottom:5px;"> <select id="submission_type" name="submission_type">
      	            <option value="ALL">ALL</option>


                   <%       
                   
                    Set SQLStmt4 = Server.CreateObject("ADODB.Command")
                    Set rs4 = Server.CreateObject ("ADODB.Recordset")
        
                    SQLStmt4.CommandText = "select code_name,long_desc from code_map where code_type=(select top 1 code_type from Code_Def where form_sid='Submission_Type')" 
                    SQLStmt4.CommandType = 1
                    Set SQLStmt4.ActiveConnection = conn
                  ' response.write "SQL = " & SQLStmt3.CommandText
                    rs4.Open SQLStmt4

                        Do Until rs4.EOF 
                          
                               %>
                        <option value="<%=rs4("code_name") %>"><%=rs4("long_desc")%></option>
                <%
                            
                    rs4.MoveNext
                    Loop
                  
                  
                  
                   %>

                    
      	         </select></div>


                         
      	           <div style="float:right;padding-right:5px;padding-bottom:5px;"> <select id="status" name="status" style="">
      	            <option value="ALL">ALL</option>
                     <option value="ASSIGNED" selected>ASSIGNED</option>
                         <option value="CLOSED">CLOSED</option>
                          <option value="FUTURE">FUTURE</option>
                          <option value="HOLD">HOLD</option>
                          <option value="NEXT">NEXT</option>
                      <option value="OPEN">OPEN</option>
                      
      	         </select></div>
                  <div style="float:right"><div style="padding-top:2px;padding-right:5px;"><b>Note Status:</b></div></div>


               <div style="clear:left"></div>
                 <div style="float:left"><div style="padding-top:2px;"><b>Open Date:</b></div></div>
      	            <div style="float:left;padding-left:5px;padding-bottom:5px;"><select id="open_date" name="open_date" style="width:100px">
      	            <option value="ALL">ALL</option>
                
                   <%
                  
                   
                    Set SQLStmt4 = Server.CreateObject("ADODB.Command")
                    Set rs4 = Server.CreateObject ("ADODB.Recordset")
        
                    SQLStmt4.CommandText = "exec get_helpdesk_open_dates " 
                    SQLStmt4.CommandType = 1
                    Set SQLStmt4.ActiveConnection = conn
                  ' response.write "SQL = " & SQLStmt3.CommandText
                    rs4.Open SQLStmt4

                        Do Until rs4.EOF 
                          
                               %>
                        <option value="<%=rs4("open_date") %>"><%=rs4("open_date")%></option>
                <%
                            
                    rs4.MoveNext
                    Loop
                  
                  
                  
                   %>
                
      	         </select>
      	        </div>


                     
               
      	            <div style="float:right;padding-right:5px;padding-bottom:5px;"><select id="closed_date" name="closed_date" style="width:100px">
      	            <option value="ALL">ALL</option>
                
                   <%
                  
                    Set SQLStmt5 = Server.CreateObject("ADODB.Command")
                    Set rs5 = Server.CreateObject ("ADODB.Recordset")
        
                    SQLStmt5.CommandText = "exec get_helpdesk_closed_dates " 
                    SQLStmt5.CommandType = 1
                    Set SQLStmt5.ActiveConnection = conn
                  ' response.write "SQL = " & SQLStmt3.CommandText
                    rs5.Open SQLStmt5

                        Do Until rs5.EOF 
                            
                               %>
                        <option value="<%=rs5("closed_date") %>"><%=rs5("closed_date")%></option>
                <%
                         rs5.MoveNext
                    Loop
                  
                  
                  
                   %>
                
      	         </select>
      	         </div>
                      <div style="float:right;padding-right:5px;"> <div style="padding-top:2px;"><b>Closed Date:</b></div></div>

                      <div style="clear:left"></div>

              
                 
                  
                      <div style="float:left"> <div style="padding-top:2px;"><b>Notify:</b></div></div>
      	            <div style="float:left;padding-left:32px;padding-bottom:5px;">
                  <select id="notify" name="notify">
                       <option value="All" >All</option>
                     <option value="-4" >Agency Contact(s)</option>
                     <option value="-2" >Global Broadcast</option>
                      <option value="-3" >C4C</option>
                   
      	          
      	          
                   <%
                  
                    Set SQLStmt5 = Server.CreateObject("ADODB.Command")
                    Set rs5 = Server.CreateObject ("ADODB.Recordset")
        
                    SQLStmt5.CommandText = "select Staff_ID,Last_Name,First_Name from staff_master where emr_staff is not null order by Last_Name" 
                    SQLStmt5.CommandType = 1
                    Set SQLStmt5.ActiveConnection = conn
                   response.write "SQL = " & SQLStmt5.CommandText
                    rs5.Open SQLStmt5

                        Do Until rs5.EOF 

                       
                        Staff_ID=rs5("Staff_ID")
                       Last_Name=rs5("Last_Name")
                       First_Name=rs5("First_Name")
                            
                            %>
                        <option value="<%=Staff_ID %>" ><%=Last_Name %>, <%=First_Name %></option>
                       
                          <%  
                       
                         rs5.MoveNext
                    Loop
                  
                  
                  
                   %>
                
      	         </select>
      	         </div>


               

                     
      	            <div style="float:right;padding-right:5px;padding-bottom:5px;"><select id="aging" name="aging" style="max-width:90px;">
      	            <option value="ALL">ALL</option>
                   
                           <option value="1">1-7 Days</option>
                           <option value="2">8-15 Days</option>
                           <option value="3">16-30 Days</option>
                          <option value="4">30+ Days</option>
                 
                
      	         </select>
      	         </div>
                      <div style="float:right;padding-right:5px;"> <div style="padding-top:2px;"><b>Aging:</b></div></div>



                       <div style="clear:left"></div>

                          <div style="float:left"> <div style="padding-top:2px;"><b>Agency:</b></div></div>
      	            <div style="float:left;padding-left:25px;padding-bottom:5px;"><select id="agency" name="agency" style="width:165px;">
      	         <option value="-1">ALL</option>
                
                   <%
                  
                    Set SQLStmt5 = Server.CreateObject("ADODB.Command")
                    Set rs5 = Server.CreateObject ("ADODB.Recordset")
        

                  if agency_program_id = -1 then
                    SQLStmt5.CommandText = "select Program_ID,Program_Name from Program_Master where CHARINDEX('-',Program_Name)=0 and PROGRAM_NAME<>'INTAKE' and program_id in(select main_program from staff_forms_master where form_type='STAFF_EMR_HELPDESK') order by Program_Name asc" 
                   else
                       SQLStmt5.CommandText = "select Program_ID,Program_Name from Program_Master where CHARINDEX('-',Program_Name)=0 and PROGRAM_NAME<>'INTAKE' and program_id in(select main_program from staff_forms_master where form_type='STAFF_EMR_HELPDESK') and program_id in(select Program_ID from Staff_Program_Assign where staff_id=" & cur_staff_id & ") order by Program_Name asc" 
                    end if
                    SQLStmt5.CommandType = 1
                    Set SQLStmt5.ActiveConnection = conn
                  ' response.write "SQL = " & SQLStmt3.CommandText
                    rs5.Open SQLStmt5

                        Do Until rs5.EOF 

                        Program_ID=rs5("Program_ID")
                        Program_Name=rs5("Program_Name")
                            
                         %>
                        <option value="<%=Program_ID %>" <%if agency_program_id <> -1 and agency_program_id=Program_ID then response.write "selected" end if %>><%=Program_Name %></option>
                       
                          <%  
                       
                         rs5.MoveNext
                    Loop
                  
                  
                  
                   %>
                
      	         </select>
      	         </div>





           <div style="float:right">
                      <%if agency_contact<>1 then %>
	           <div style="float:left;padding-left:10px;"><input type="checkbox" name="search_assigned_to_me" value="1" id="search_assigned_to_me" checked><label for="search_assigned_to_me">&nbsp;&nbsp;Assigned to Me</label></div>
               <div style="clear:both;"></div> 
               <div style="float:left;padding-left:10px;"><input type="checkbox" name="search_include_closed" value="1" id="search_include_closed" ><label for="search_include_closed">&nbsp;&nbsp;Include Closed</label></div>
            <%end if %>
            </div>


                       <div style="clear:left"></div>


                  



                   <div style="float:right;padding-right:10px;"><a id="search_submissions" href="" target="_blank" <%response.write search_submissions_disable %>><img src="images/search-icon48.png" height="34" border="0" alt="Search iCentrix Change Tracking" title="Search iCentrix Change Tracking" ><span style="float:right;padding-top:7px;padding-left:5px;"><b>[SEARCH]</b></span></a>

                   </div>


              

                     
                
               
              

            <%else %>
                <td class="blueFontSmall" width="45%" colspan="2" bgcolor="#E4EDF6" align="center">
                   <span id="create_form_label"><b>NEW STAFF FORM TO ADD:</b></span><br />
                
        
                
                
                 <div id="possible_parent_forms"> </div>  
	            <div id="create_form_div">
		            <% 
		            form_avail_count = -1
		            
		            if Request.QueryString("sid") <> "" THEN 
		               
                            form_family_level = "0"
                    		
                    		form_avail_count = 0
                    			                        			    
                            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
                            Set rs2 = Server.CreateObject ("ADODB.Recordset")
                                
  	                        SQLStmt2.CommandText = "exec get_creatable_forms_for_staff_and_user " & Request.QueryString("sid") & ",'" & Session("user_name") & "'"
  	                        SQLStmt2.CommandType = 1
  	                        Set SQLStmt2.ActiveConnection = conn
  	                        SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                     '  response.write "SQL = " & SQLStmt2.CommandText
  	                        rs2.Open SQLStmt2
  	                        Do Until rs2.EOF
							
							cur_required_forms = rs2("Required_Forms")
  	                        cur_form_family = rs2("Form_Family")
  	                
  	                            if form_avail_count = 0 THEN
  	                            %>
  	                            <select name="create_form_type" id="create_form_type">
		                            <option value="">Select Form Type to Create</option>
  	                            <%
  	                            end if
  	                            
  	                                form_avail_count = form_avail_count + 1
                                	                    
	                                if (cur_form_family = "1" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "1"
	                                    %>
	                                    <optgroup label="Employment" id="EMPOYMENT">
	                                    <%  
	                                elseif (cur_form_family = "2" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "2"
	                                    %>
	                                    <optgroup label="Legal Forms" id="LF">
	                                    <%  
	                                elseif (cur_form_family = "3" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "3"
	                                    %>
	                                    <optgroup label="Comprehensive Assessment Forms" id="CAF">
	                                    <%  
	                                elseif (cur_form_family = "4" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "4"
	                                    %>
	                                    <optgroup label="Addendum/Assessment Forms" id="AAF">
	                                    <%  
	                                elseif (cur_form_family = "5" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "5"
	                                    %>
	                                    <optgroup label="Individualized Action Plan Forms" id="IAP">
	                                    <%  
	                                elseif (cur_form_family = "6" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "6"
	                                    %>
	                                    <optgroup label="Progress and Service Notes" id="PSN">
	                                    <%  
	                                elseif (cur_form_family = "7" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "7"
	                                    %>
	                                    <optgroup label="Transition and Discharge" id="TD">    
	                                    <%  
	                                elseif (cur_form_family = "8" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "8"
	                                    %>
	                                    <optgroup label="CBHI" id="CBHI">    
	                                    <%  
	                                elseif (cur_form_family = "9" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "9"
	                                    %>
	                                    <optgroup label="Brien - Acute Care/ESP/MCI/Respite/Stab Forms" id="BAC">    
	                                    <%  
	                                elseif (cur_form_family = "x10" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "x10"
	                                    %>
	                                    <optgroup label="Brien - Adult Outpatient/MH/SU Forms" id="BAO">
	                                    <%
	                                elseif (cur_form_family = "x11" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "x11"
	                                    %>
	                                    <optgroup label="Brien - Child &amp; Adolescent/CSA/CSU/CBHI Forms" id="BCA">
	                                    <%
	                                elseif (cur_form_family = "x12" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "x12"
	                                    %>
	                                    <optgroup label="Brien - Community Services/ADH/CBFS Forms" id="BCS">
	                                    <%
	                                elseif (cur_form_family = "x13" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "x13"
	                                    %>
	                                    <optgroup label="Brien - Psychiatric Services/MD/CNS Forms" id="BPS">    
	                                    <%  
	                                elseif (cur_form_family = "x14" and cur_form_family <> form_family_level) THEN
	                                    form_family_level = "x14"
	                                    %>
	                                    <optgroup label="Other Forms" id="Other">
	                                    <%  
	                                end if	                             
	                                    %>
	                                    <option value="<%=rs2("Form_Type")%>"><%Response.write rs2("Form_Description")%></option>
	                                    <%
                    	            
	                                disabled_flag = ""
                    	                                                	                   
                               rs2.MoveNext
                           Loop
                        else
                        %>
                        Choose a Staff 
		                <% end if  %>                    	
                    <% if form_avail_count = 0 THEN %>
                    No forms available for this staff
                    <% elseif form_avail_count <> -1 THEN %>
                    </select>
                    &nbsp;&nbsp;<span id="add_new_form_button_span"><a href="javascript: void(0);" onclick="checkForm();" id="add_new_form_button"><img src="images/button_add_new.jpg" width="28" height="16" border="0" class="textmiddle" alt="Create New Form" title="Create New Form" /></a></span>
                    <% else %>
                    &nbsp;
                    <% end if  %>		    
	                
                    
	            </div>	            
	            <br />
	        </td>	  


             </td>
          <%end if%>
                
                
                
                
                
                    
	    </tr>
	    <% if Request.QueryString("sid") <> "" and Session("active_staff_manage_staff_auth") = "1" THEN %>
	    <tr>
	      <td align="left" colspan="4"><a href="edit_staff.asp?sid=<%=choosen_staff_id %>&fromModule=1" target="_blank">[<b>Edit Staff</b>]</a></td>         
	    </tr>
	    <% end if %>
        </table>
       
	 
	</div>
	<%
	qs_pid = Request.QueryString("fn")
	
	if qs_pid = "" THEN
	    qs_pid = "1000"
	end if   	
	%>
    <div id="spacer6"></div>
	
   	  
    <table id="forms">

	 <!--PUT ASP FOR PARENT FORM LOOP HERE-->
	  <%
	  if choosen_staff_id <> "&nbsp;" and choosen_staff_id <> "" THEN
            
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



          cur_program_id=-1

          if Request.QueryString("pCode")<>"" then
           cur_program_id=Request.QueryString("pCode")
          else
           cur_program_id=agency_program_id
          end if


           if Request.QueryString("status")<>"" then
           cur_note_status=Request.QueryString("status")
          else
           cur_note_status="ASSIGNED"
          end if

           if Request.QueryString("atm")="0" or agency_contact=1 then
           assigned_to_me=0
          else
           assigned_to_me=1
          end if


           if Request.QueryString("ic")="1" or agency_contact=1 then
           include_closed=1
          else
           include_closed=0
          end if

          


        
        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	    SQLStmt2.CommandText = "exec get_form_list_for_staff " & choosen_staff_id & ",'" & Session("user_name") & "','" & cur_filt_by & "'," & form_group_filter & ",-1" & ",'" & Request.QueryString("sb") & "'," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & ",'" & Request.QueryString("tf") & "','" & Request.QueryString("ufid") & "'," & cur_program_id & ",'" & cur_note_status & "'," & assigned_to_me & "," & include_closed
  	    Set SQLStmt2.ActiveConnection = conn
  	    SQLStmt2.CommandTimeout = 45 'Timeout per Command
   'response.write "SQL 1 = " & SQLStmt2.CommandText
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
    
  	        SQLStmt3.CommandText = "exec get_staff_form_info_without_content " & rs2("Unique_Form_ID")
  	        SQLStmt3.CommandType = 1
  	        Set SQLStmt3.ActiveConnection = conn
  	        SQLStmt3.CommandTimeout = 45 'Timeout per Command
  	       ' response.write "SQL = " & SQLStmt3.CommandText
  	        rs3.Open SQLStmt3

            cur_copied_form_id = rs3("Parent_Form_ID")
            blob_val = rs3("File_Content")
            cur_create_date = rs3("Create_Date")
            has_children = rs3("has_children")            
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
	        
	        'cur_hide_show_count = rs2("hide_show_count")
	      
         
            
	        cur_req_met = 0
	        cur_req_msg = ""
    	    
    	    cur_form_hover = "This form has " & cur_num_pages & " page(s)" ' and was created on " & cur_create_date & "."
    	    
	        if cur_req_msg <> "" THEN
	            cur_req_msg = "This form requires the completion of " & cur_req_msg & " before it can be started"
	        end if
    	    
	        if use_header = 1 and use_secondary_header <> 1 THEN
	      %>
    	  
	      <tr>
	        <td class="formRow" colspan="7" valign="bottom" <% if cur_type = "PAPER" THEN %>style="background-color:#FFFBD2;"<% end if %>>
	        
	        <%if cur_type <> "PAPER" and cur_hide_show_count > 1 and (INT(Request.QueryString("ulfid")) <> INT(cur_linked_form_id) or Request.QueryString("uft") <> cur_type) THEN %>
	        <a href="javascript:void();" onclick="undoFilter('<%=cur_form_id%>','<%=cur_linked_form_id %>','<%=cur_type%>');" class="filterOff" style="text-decoration:none;" alt="Showing Filtered Forms"><font style="font-size:14px;" >+</font></a>
	        <% end if  %>
	        
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
	              <%if cur_type = "GCFTR" THEN
                 
                   Set SQLStmtComp = Server.CreateObject("ADODB.Command")
  	            Set rsComp = Server.CreateObject ("ADODB.Recordset")

  	            SQLStmtComp.CommandText = "select Staff_ID from Staff_Master where user_name = '" & cur_update_user & "'"
  	            
  	            SQLStmtComp.CommandType = 1
  	            Set SQLStmtComp.ActiveConnection = conn
  	            SQLStmtComp.CommandTimeout = 45 'Timeout per Command
  	            rsComp.Open SQLStmtComp

                if NOT rsComp.EOF THEN
  	                Staff_ID = rsComp("Staff_ID")
  	             
  	            end if
                 
                  %>
                 <img src="images/form_iconSmall.gif" border="0"> 
	            <a onclick="popupwindow('get_form_copy.asp?uid=<%=cur_form_id%>&ft=<%=cur_type%>&cid=<%=client_id%>&pid=<%=cur_main_prog_id2 %>&sid=<%=Staff_ID %>&dp=&if=Fix',800,1400,'<%=cur_form_id%>');" href="#"><img title="Update Form" alt="" src="images/plus_sign.gif" border="0"></a>
                 <%else %> 
                 <img src="images/form_iconSmall.gif" border="0">
                 <%end if %>
	        <%end if %>
	        
	        <%if cur_access = "V" or cur_status = "Group Lock" THEN%>
	            <font color="#777777"><%=cur_form_name%></font>
	        <%else%>
	           <%if cur_type = "PAPER" THEN%>
	             <font color="#333300"><%=cur_form_name%></font>	 
	             <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&fromStaff=1',800,1000,'getAttachmentsWindow');">&nbsp;&nbsp;<%if cur_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg" border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg" border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>            
	           <%else
                   
                   notified_users = ""

                   if cur_type = "STAFF_EMR_HELPDESK" THEN
                         Set SQLStmtComp = Server.CreateObject("ADODB.Command")
  	                    Set rsComp = Server.CreateObject ("ADODB.Recordset")

  	                    SQLStmtComp.CommandText = "exec get_emrhelpdesk_users_notified " & cur_form_id 
  	            
  	                    SQLStmtComp.CommandType = 1
  	                    Set SQLStmtComp.ActiveConnection = conn
  	                    SQLStmtComp.CommandTimeout = 45 'Timeout per Command
  	                    rsComp.Open SQLStmtComp


                        open_closed_icon = "OPEN"

                       nu_count=0
                       Do Until rsComp.EOF 
                   
                      if nu_count=0 then
                       notified_users =  " - " & rsComp("user_names")
                      else
                       notified_users = notified_users & ", " & rsComp("user_names")
                      end if 	
  	                     
  	                   
                        nu_count=nu_count+1
  	                    rsComp.MoveNext
                        Loop
                  end if
                   
                   %>
	             <%response.write cur_form_name & notified_users %>

                <%if open_closed_icon="CLOSED2" then  %>

                  &nbsp; <img title='CLOSED' alt="" src='images/closed.jpg' border='0' />
               <%elseif open_closed_icon="OPEN2" then  %>
                &nbsp; <img title='OPEN' alt="" src='images/open.jpg' border='0'>
              
	           <%
                  end if
                   end if%>    
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
	            <font color="#777777"><%=cur_form_name%></font>
	        <%else%>
	           <%if cur_type = "PAPER" THEN%>
	             <font color="#333300"><%=cur_form_name%></font>
	             <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&fromStaff=1',800,1000,'getAttachmentsWindow');">&nbsp;&nbsp;<%if cur_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg" border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg" border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>
	           <%else%>
	             <%=cur_form_name%>
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
            <%if (cur_staff_role_desc = "Administrator") or (cur_status = "In-Process" and (cur_access = "E" or cur_access = "L"))  THEN 
            %>
                <% if (cur_create_user = Session("user_name") or cur_staff_role_desc = "Administrator") and cur_type <> "PAPER" THEN %>
                <a href="#" onclick="confirmFileDelete('<%=Replace(cur_form_name,"'","")%>',<%=cur_form_id%>);"><img title="Delete in-process form" src="images/delete_file.gif" border="0"/></a>
                <%end if %>
            <%end if %>
    
            <!-- NEW FORM COPY PROCESS FOR IAPs-->
            <%
            if cur_type <> "PAPER" and cur_type <> "INTAKE" and cur_type <> "STAFF_EMR_HELPDESK" THEN %>
                <a href="choose_parent_for_copy.asp?width=700&height=300&uid=<%=cur_form_id%>&cid=<%=client_id%>&usid=<%=user_staff_id %>&ft=<%=cur_type %>" title="FAQs" class="thickbox"><img src="images/MSDPCopyIcon.jpg" title="Make a copy of this form" border="0"></a>
            <%end if %> 

            <%if cur_type <> "PAPER" THEN %>
                <a href="#" onclick="popupwindow('view_history.asp?uid=<%=cur_form_id%>',800,1000,'getHistoryWindow');"><img src="images/button_history.jpg"  border="0" alt="History" title="History" /></a>

                 <%if cur_type <> "STAFF_EMR_HELPDESK" THEN %>
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
                 <%end if %>
                <%if ((cur_status = "Finalized" or cur_status = "CLOSED") or cur_access = "V" or cur_status = "Group Lock") THEN%>
                    <% if cur_type = "BTMS" or cur_type = "CONTRACTS" or cur_type = "DMTC" or cur_type = "DHSPS" or cur_type = "FDIP" or cur_type = "FFSE" or cur_type = "GCFTR" or cur_type = "FMTP" or cur_type = "FTR" or cur_type = "HIPPA" or cur_type = "IMMUNE" or cur_type ="MEDLIST" or cur_type ="MRC" or cur_type ="RTF" or cur_type ="SPMU" or cur_type ="MS" or cur_type ="PAS" or cur_type ="MS_V2" or cur_type ="MS_PROC" or cur_type ="MMR" or cur_type ="DDS" THEN%>
                        <div align="right" style="float:right;"><a href="javascript:void()" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=cur_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                    <% else %>
                        <div align="right" style="float:right;"><a href="javascript:void()" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=cur_form_id%>',800,1000,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                    <% end if %>
                <%elseif (cur_status = "In-Process" or cur_status = "OPEN" or cur_status = "ASSIGNED" or cur_status = "HOLD" or cur_status = "FUTURE" or cur_status = "NEXT") and (cur_access = "E" or cur_access = "L") THEN%>
                    <% if cur_type = "BTMS" or cur_type = "CONTRACTS" or cur_type = "DMTC" or cur_type = "DHSPS" or cur_type = "FDIP" or cur_type = "FFSE" or cur_type = "FMTP" or cur_type = "GCFTR" or cur_type = "FTR" or cur_type = "HIPPA" or cur_type = "IMMUNE" or cur_type ="MEDLIST" or cur_type ="MRC" or cur_type ="RTF" or cur_type ="SPMU" or cur_type ="MS" or cur_type ="PAS" or cur_type ="MS_V2" or cur_type ="MS_PROC" or cur_type ="MMR" or cur_type ="DDS" THEN%>
                        <div align="right" style="float:right;"><a href="javascript:void()" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=cur_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                    <% else %>
                        <div align="right" style="float:right;"><a href="javascript:void()" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=cur_form_id%>',800,1000,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                    <% end if %>
                <%end if %>
            <% end if %>
 
            
            <% if cur_type <> "PAPER" THEN %>
            <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&fromStaff=1',800,1000,'getAttachmentsWindow');">
           <%if cur_has_attachments = "Yes" THEN %>
                <img src="images/MSDPattachmentYES.jpg"  border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" />
              <%else %>
                <img src="images/MSDPattachment.jpg"  border="0" title="No Attachments, Click to Add" alt="Attachments" />
             <%end if %>

            </a>&nbsp; 
        
              <%if cur_status = "CLOSED" THEN %>
             <a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=copy_staff&uid=<%=cur_form_id%>&sid=<%=Request.QueryString("sid") %>&if=Fix',800,1000,'<%=cur_form_id%>');""> <img src="images/plus_sign.gif"  border="0" title="Click to Copy" alt="Attachments" /></a>&nbsp; 
                  
            
            <% end if
                end if  %>    
        </td>
    <% end if  %>
    
    <%if cur_type = "PAPER" THEN %>
            
              <td colspan="7" class="blackFontSmall" align="center" style="background-color:#FFFBD2;"> 
              <%
             Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	         Set rs3 = Server.CreateObject ("ADODB.Recordset")
                        
  	         SQLStmt3.CommandText = "exec get_doc_type_counts_for_staff " & Request.QueryString("sid")
  	         SQLStmt3.CommandType = 1
  	         Set SQLStmt3.ActiveConnection = conn
  	       '  response.write "SQL = " & SQLStmt3.CommandText
  	         rs3.Open SQLStmt3
                	            
	         Do Until rs3.EOF
               	    
	            cur_doc_category = rs3("category")
	            cur_doc_desc = rs3("desc")
	            cur_doc_count = rs3("doc_count")

                  if (cur_doc_category<>"Excel_Reports" and cur_doc_category<>"Misc" and cur_doc_category<>"To_Be_Discussed" and cur_doc_category<>"User_Documentation" and cur_doc_category<>"Wish_List") or Request.QueryString("sid") = "1344" then
	            %>
	                <span style="white-space: nowrap;"><% if cur_doc_count <> 0 THEN %><a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=cur_form_id%>&doc_type=<%=cur_doc_category %>&doc_subtype=All&fromStaff=1',800,1000,'getAttachmentsWindow');"><b><%=cur_doc_desc %></b></a><% else %><%=cur_doc_desc %> <% end if  %> (<b><%=cur_doc_count%></b>)</span>&nbsp;
	            <%
                  end if

	            rs3.MoveNext
                	        
	         Loop %>              
</td>
            
    <% else %>
		    <td class="status">
		    <%if has_the_needs = "1" THEN%><font color="#993300"><%=cur_status%></font>
		    <%elseif cur_status = "Finalized" or cur_status = "CLOSED" THEN%><span class="finalized"><%=cur_status%></span>
		    <%else%><%=cur_status%>
		    <%end if%>
		    <br />
		    <%if cur_copied_form_id <> "" and cur_copied_form_id <> "0" THEN %>
		    <div onmouseover="showToolTip(event,'Copied From <%=cur_copied_form_type%> with create date of <br / > <%=cur_copied_form_create_date%> <br/><br /> Created on <%=cur_create_date%> <br /> By <%=REPLACE(cur_create_user_name, "'", "\")%> <br /> Program: <%=cur_main_prog %> <br /> Staff: <%=REPLACE(cur_main_staff,"'", "\'") %>')" onmouseout="hideToolTip()">
		    <%else%>
		    <div onmouseover="showToolTip(event,'Created on <%=cur_create_date%> <br /> By <%=REPLACE(cur_create_user_name, "'", "\'")%> <br /> Program: <%=cur_main_prog %> <br /> Staff: <%=REPLACE(cur_main_staff, "'", "\'") %>')" onmouseout="hideToolTip()">
		    <%end if %>
		    <%if has_the_needs = "1" THEN%><font color="#993300"><%=cur_last_update%></font>
		    <%elseif cur_status = "Finalized" or cur_status = "CLOSED" THEN%><span class="finalized"><%=cur_last_update%></span>
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
  	        'response.write "SQL = " & SQLStmtSI.CommandText
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
	      <%end if %>
	      
	 <%if cur_type = "PAPER" THEN%>

       	 <%if Request.QueryString("sid")="1344" THEN%>
	      <tr>
         <% else %>
             <tr style="display:none;">
         <% end if
             
            ' response.write cur_status
             
              %>
	           <td border="0" width="100%" colspan="7" style="font-size: 7pt; font-family: verdana; font-weight: bold;" valign="bottom">	
                   
                   
                    <div style="float:left">
                   
                    &nbsp;Note Status:&nbsp;&nbsp;  
                    
                  <select id="form_status" name="form_status" onchange="filterFormStatus(this.value);">
      	            <option value="ALL" <% if cur_note_status="ALL" then response.write "selected" end if %>>ALL</option>
                     <option value="ASSIGNED" <% if cur_note_status="ASSIGNED" then response.write "selected" end if %>>ASSIGNED</option>
                         <option value="CLOSED" <% if cur_note_status="CLOSED" then response.write "selected" end if %>>CLOSED</option>
                       <option value="CURRENT" <% if cur_note_status="CURRENT" then response.write "selected" end if %>>CURRENT</option>
                          <option value="FUTURE" <% if cur_note_status="FUTURE" then response.write "selected" end if %>>FUTURE</option>
                          <option value="HOLD" <% if cur_note_status="HOLD" then response.write "selected" end if %>>HOLD</option>
                          <option value="NEXT" <% if cur_note_status="NEXT" then response.write "selected" end if %>>NEXT</option>
                      <option value="OPEN" <% if cur_note_status="OPEN" then response.write "selected" end if %>>OPEN</option>
                      
      	         </select>
	          &nbsp;&nbsp;

	          Agency:&nbsp;&nbsp;<select name="notes_filter" id="notes_filter" onchange="filterFormProgram(this.value);">

               
	             <option value="-1">Show All Agencies</option>
               
	                <%
	                if Request.QueryString("sid") <> "" THEN
	                    Set SQLStmt2a = Server.CreateObject("ADODB.Command")
  	                    Set rs2a = Server.CreateObject ("ADODB.Recordset")
                        if agency_program_id = -1 then
                          SQLStmt2a.CommandText = "select Program_ID,Program_Name from Program_Master where CHARINDEX('-',Program_Name)=0 and PROGRAM_NAME<>'INTAKE' and program_id in(select main_program from staff_forms_master where form_type='STAFF_EMR_HELPDESK') order by Program_Name asc"  
                        else
                          SQLStmt2a.CommandText = "select Program_ID,Program_Name from Program_Master where CHARINDEX('-',Program_Name)=0 and PROGRAM_NAME<>'INTAKE' and program_id in(select main_program from staff_forms_master where form_type='STAFF_EMR_HELPDESK') and program_id in(select Program_ID from Staff_Program_Assign where staff_id=" & cur_staff_id & ") order by Program_Name asc" 
                        end if
  	                    SQLStmt2a.CommandType = 1
  	                    Set SQLStmt2a.ActiveConnection = conn
  	                    SQLStmt2a.CommandTimeout = 45 'Timeout per Command
  	                    'response.write "SQL = " & SQLStmt2.CommandText
  	                    rs2a.Open SQLStmt2a
  	                    Do Until rs2a.EOF
      	                
  	                    cur_pid = rs2a("Program_ID")
	                    %>
	                    <option value="<%=cur_pid%>" <% if Int(cur_pid) = Int(cur_program_id) THEN %>selected<% end if %> ><%Response.write rs2a("Program_Name")%></option>
	                    <%
	                    rs2a.MoveNext
                        Loop
                    end if
	                %>
	            </select> 
	          </div>

            <%if agency_contact<>1 then %>
	           <div style="float:left;padding-left:10px;"><input type="checkbox" name="assigned_to_me" value="1" id="assigned_to_me" <%if assigned_to_me=1 then response.write "checked" end if %>><label for="assigned_to_me">&nbsp;&nbsp;Assigned to Me</label></div>
                   <div style="float:left;padding-left:10px;"><input type="checkbox" name="include_closed" value="1" id="include_closed" <%if include_closed=1 or Request.QueryString("status")="CLOSED" then response.write "checked" end if %>><label for="include_closed">&nbsp;&nbsp;Include Closed</label></div>
            <%end if %>
            
	        </td>
	   </tr>
	        
	  <tr>
	    <td class="tabBarLeft"><img src="images/FormsList5.gif" width="349" height="18" border="0" alt="Form Name" title="Form Name" /></td>
		<td class="tabBarCenter"><img src="images/Status5.gif" width="72" height="18" border="0" alt="Status" title="Status" /></td>
		<td class="tabBarCenter" colspan="5"><img src="images/SignaturePanel5a.gif" width="543" height="18" border="0" alt="Signatures" title="Signatures" /></td>
	  </tr>  
   <%end if
	      
    	  	
	  	    if cur_form_id <> "" and has_children = "Yes" THEN 
	        '---------------START INNER FORMS
	        inner_form_name = ""
	        cur_type = ""
            
            Set SQLStmt30 = Server.CreateObject("ADODB.Command")
  	        Set rs30 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt30.CommandText = "exec get_form_list_for_staff " & choosen_staff_id & ",'" & Session("user_name") & "','" & cur_filt_by & "'," & form_group_filter & "," & cur_form_id & ",'" & Request.QueryString("sb") & "'," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & ",'" & Request.QueryString("tf") & "','" & Request.QueryString("ufid") & "'"
  	        SQLStmt30.CommandType = 1
  	        Set SQLStmt30.ActiveConnection = conn
  	        SQLStmt30.CommandTimeout = 45 'Timeout per Command
  	        ' if cur_form_id = "13087" THEN
  	         '   response.write "SQL 1 = " & SQLStmt30.CommandText
  	        ' end if
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
        
  	            SQLStmt3.CommandText = "exec get_staff_form_info_without_content " & rs30("Unique_Form_ID")
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
	        	        
	            inner_display_status = ""
	            inner_display_date = ""
	        
	            'DETERMINE DISPLAY STATUS/DATE HERE
                if inner_status = "In-Process" or inner_status = "OPEN" or inner_status = "CLOSED"   THEN
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
                
                if inner_display_status = "" THEN
                    inner_display_status = "Finalized"
                    inner_display_date = inner_last_update
                end if	            
	                   	    
	            inner_form_hover = "This form has " & inner_num_pages & " page(s)"
	                    	    
	            if inner_new_form_type = 1 and inner_type <> "STAFF_EMR_HELPDESK" THEN
	            %>
	          <tr><td colspan="7" class="childDivider"></td></tr>
	          <tr>
	            <td class="formRowChild" valign="bottom">
	            <%if cur_type <> "PAPER" and inner_hide_show_count > 1 and (INT(Request.QueryString("ulfid")) <> INT(inner_linked_form_id) or Request.QueryString("uft") <> inner_type) THEN %>
	        <a href="javascript:void();" onclick="undoFilter('<%=inner_form_id%>','<%=inner_linked_form_id %>','<%=inner_type%>');" class="filterOff" style="text-decoration:none;" alt="Showing Filtered Forms"><font style="font-size:14px;" >+</font></a>
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
	            <%else
                   
                     %>
	                <img src="images/form_iconSmall.gif" border="0">
	            <%end if 
                %>
	            
	            <%
                   
                    if inner_access = "V" THEN%>
	                <font color="#777777"><%=inner_form_name%></font>
	            <%else%>
	                <%=inner_form_name%>
	            <%end if
                  f%>
                </td>
		        <td  class="NoneStatus"><%if inner_blob_val = "" or inner_linked_form_id <> cur_form_id THEN %><b>NONE</b><%else %>&nbsp;<%end if %></td>

		        <td class="client">&nbsp;</td>
		        <td class="provider">&nbsp;</td>
		        <td class="guardian">&nbsp;</td>
		        <td class="md">&nbsp;</td>
		        <td class="supervisor">&nbsp;</td>

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
            <a href="#" onclick="confirmFileDelete('<%=Replace(cur_form_name,"'","")%>',<%=inner_form_id%>);"><img alt="Delete in-process form" src="images/delete_file.gif" border="0"/></a>
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
        <%if (inner_status = "Finalized" or inner_status = "OPEN" or inner_status = "CLOSED") or inner_access = "V" THEN%>
            <% if inner_type ="DDS" THEN%>
                        <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=inner_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                <% else %>
            <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=inner_form_id%>',800,1000,'<%=inner_form_id%>');"><img src="images/button_view.jpg"  border="0" alt="View" title="View" /></a></div>
                <% end if %>
        <%elseif inner_status = "In-Process" and (inner_access = "E" or inner_access = "L") THEN%>
            <% if inner_type ="DDS" THEN%>
                <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=inner_form_id%>',800,1400,'<%=cur_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
                <% else %>
            <div align="right" style="float:right;"><a href="#" onclick="popupwindow('http://<%=url_org_name%>:9080/samples/WebformProxy?pathtype=staff&uid=<%=inner_form_id%>',800,1000,'<%=inner_form_id%>');"><img src="images/button_edit.jpg"  border="0" alt="Edit" title="Edit" /></a></div>
            <% end if %>
        <%end if %>
        <a href="#" onclick="popupwindow('view_attachments.asp?uid=<%=inner_form_id%>&fromStaff=1',800,1000,'getAttachmentsWindow');"><%if inner_has_attachments = "Yes" THEN %><img src="images/MSDPattachmentYES.jpg"  border="0" title="Attachments are present, Click to View/Edit" alt="Attachments" /><%else %><img src="images/MSDPattachment.jpg"  border="0" title="No Attachments, Click to Add" alt="Attachments" /><%end if %></a>&nbsp; 

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
  	            'response.write "SQL = " & SQLStmtSI.CommandText
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
	            
	        rs30.MoveNext
            Loop            
            
	        '---------------END INNER FORMS
	        end if
	        last_cur_type = cur_type
	    end if
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
