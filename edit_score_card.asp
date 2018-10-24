<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>iCentrix Corp. Electronic Medical Records</title>
    
    <%      if Request.QueryString("action") = "delete_category" then
           
            Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC.CommandText = "DELETE FROM Score_Card_Form_Content WHERE form_type_id='" & Request.QueryString("form_type_id") & "' AND Item_Cat='"  & Request.QueryString("category") & "'"
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	        'response.write(Request.QueryString("form_type_id"))
  	        rsSC.Open SQLStmtSC
             Response.Redirect "edit_score_card.asp?form_type_id=" & Request("form_type_id")
           end if
           
            if Request.QueryString("action") = "delete_item" then
           
            Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC.CommandText = "DELETE FROM Score_Card_Form_Content WHERE ID='"  & Request.QueryString("item_id") & "'" 
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rsSC.Open SQLStmtSC
             Response.Redirect "edit_score_card.asp?form_type_id=" & Request.QueryString("form_type_id")
           end if
            
           if Request.QueryString("action") = "edit_cat_seq" then
           
            Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC.CommandText = "UPDATE Score_Card_Form_Content SET cat_seq='"  & Request.QueryString("cat_seq") & "' WHERE Item_Cat='" & Request.QueryString("cat_id") & "' AND Item=''"
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rsSC.Open SQLStmtSC
            'Response.Write("test")
            Response.Redirect "edit_score_card.asp?form_type_id=" & Request.QueryString("form_type_id")
           end if
           
           if Request.QueryString("action") = "edit_item_seq" then
           
            Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC.CommandText = "UPDATE Score_Card_Form_Content SET item_seq='"  & Request.QueryString("item_seq") & "' WHERE ID=" & Request.QueryString("item_id") 
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	        'response.write "SQL = " & SQLStmt2.CommandText
  	        rsSC.Open SQLStmtSC
            'Response.Write("test")
            
            Response.Redirect "edit_score_card.asp?form_type_id=" & Request.QueryString("form_type_id")
           end if
               
           
            Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC.CommandText = "exec get_score_card_grade_range '"  & Request.QueryString("form_type_id") & "'" 
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	       ' response.write "SQL = " & SQLStmtSC.CommandText
  	        rsSC.Open SQLStmtSC
  	        
  	       
  	        If Not rsSC.EOF Then
  	        
  	         
  	        Form_Type_ID = rsSC("Form_Type_ID")
  	        Form_Name = rsSC("Form_Name")
  	        Is_Active = rsSC("Is_Active")
  	        Creator = rsSC("Creator")
  	        Date_of_Creation = rsSC("Date_of_Creation")
  	        VG_Low = rsSC("VG_Low")
  	        VG_High = rsSC("VG_High")
  	        Adequate_Low = rsSC("Adequate_Low")
  	        Adequate_High = rsSC("Adequate_High")
  	        NIA_Low = rsSC("NIA_Low")
  	        NIA_High = rsSC("NIA_High")
  	        Item_Type = rsSC("Item_Type")
  	        Max_Score = rsSC("Max_Count")
  	       
  	        
  	        End If 
  	        
  	         
  	        forms_check = ""
  	        attachments_checked = ""
  	        manual_checked = ""
  	        
  	        if Item_Type = "forms" or Item_Type = "" or IsNull(rsSC("Item_Type")) then
  	        forms_checked = "checked"
  	        end if
  	        
  	        if Item_Type = "attachments" then
  	        attachments_checked = "checked"
  	        end if
  	        
  	        if Item_Type = "manual" then
  	        manual_checked  = "checked"
  	        end if
  	        
  	        on_checked = ""
  	        off_checked = ""
  	        
  	         if Is_Active = "True" then
  	         on_checked = "checked"
  	         end if
  	         
  	         if Is_Active = "False" then
  	         off_checked = "checked"
  	         end if
 	        
  	        if Request("page") = "add_category" then
  	        
  	       Response.Write(Request("category"))
  	       Response.Write(Request("form_type_id"))
  	        
  	        Set SQLStmtC = Server.CreateObject("ADODB.Command")
  	       Set rsC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtC.CommandText = "exec update_score_card_cat '"  & Request("form_type_id") & "','" & Request("category") & "'"
  	        SQLStmtC.CommandType = 1
  	        Set SQLStmtC.ActiveConnection = conn
  	        response.write "SQL = " & SQLStmt2.CommandText
  	        rsC.Open SQLStmtC
  	        
  	        Response.Redirect "edit_score_card.asp?form_type_id=" & Request("form_type_id")
  	        
  	        end if    
  	        
  	        
  	      if Request("page") = "add_item" then
  	      
  	    '  Response.Write(Request("form_name"))
  	     '  Response.Write(Request("doc_type"))
  	       
  	        Set SQLStmtC = Server.CreateObject("ADODB.Command")
  	       Set rsC = Server.CreateObject ("ADODB.Recordset")
           If Request("form_name") <> "" and Request("doc_type") = ""  Then 
  	        SQLStmtC.CommandText = "exec update_score_card_item '"  & Request("form_type_id") & "','" & Request("item_cat") & "','" & Replace(Replace(Request("form_name"),"'","''"),"*","")  & "','','','" & Request("create_form_type") & "','" & Request("doc_type") & "','" & Request("doc_subtype") & "'"
  	        Elseif Request("doc_type") <> "" Then
  	         SQLStmtC.CommandText = "exec update_score_card_item '"  & Request("form_type_id") & "','" & Request("item_cat") & "','" & Replace(Replace(Request("form_name"),"'","''"),"*","")  & "','','','PAPER','" & Request("doc_type") & "','" & Request("doc_subtype") & "'"
  	        Else
  	        SQLStmtC.CommandText = "exec update_score_card_item '"  & Request("form_type_id") & "','" & Request("item_cat") & "','" & Replace(Request("item"),"'","''") & "','','','" & Request("create_form_type") & "'"
  	        End If
  	        SQLStmtC.CommandType = 1
  	        Set SQLStmtC.ActiveConnection = conn
  	        response.write "SQL = " & SQLStmtC.CommandText
  	        rsC.Open SQLStmtC
  	        
  	        Response.Redirect "edit_score_card.asp?form_type_id=" & Request("form_type_id")
  	        
  	       ' Response.Write(Request("item"))
  	      '  Response.Write(Request("item_cat"))
  	      '  Response.Write(Request("form_type_id"))
  	        
  	      
  	     
  	        
  	        end if   
  	        
  	        if Request("page") = "edit_cat_name" then
  	       ' Response.Write(Request("form_type_id"))
  	         ' Response.Write(Request("edit_cat_name"))
  	          ' Response.Write(Request("old_cat"))
  	        Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC.CommandText = "UPDATE Score_Card_Form_Content SET Item_Cat='"  & Request("edit_cat_name") & "' WHERE Item_Cat='" & Request("old_cat") & "' AND Form_Type_ID='" & Request("form_type_id") & "'"
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	        response.write "SQL = " & SQLStmt2.CommandText
  	        rsSC.Open SQLStmtSC
            Response.Write("test")
            Response.Redirect "edit_score_card.asp?form_type_id=" & Request("form_type_id")
  	        
  	        
  	        end if
  	        
  	        
  	     if Request("page") = "save_changes" then
  	      
  	      hidden_item_id = 0
  	      item_dor = ""
  	      item_cr = ""
  	       For i = 0 To Request("total_items") - 1
  	       item_dor = "item_dorr" & i
         '  response.write(Request.Form(item_dor))
           item_cr = "item_cr" & i
          ' response.write(Request.Form(item_cr))
            hidden_item_id = "hidden_item_id" & i
         '  response.write(Request.Form(hidden_item_id))
           
           Set SQLStmtC = Server.CreateObject("ADODB.Command")
  	        Set rsC = Server.CreateObject ("ADODB.Recordset")
            SQLStmtC.CommandText = "exec update_score_card_grid "  & Request.Form(hidden_item_id) & ",'" & Request.Form(item_cr) & "','" & Request.Form(item_dor) & "'"
  	        SQLStmtC.CommandType = 1
  	        Set SQLStmtC.ActiveConnection = conn
  	      '  response.write "SQL = " & SQLStmtC.CommandText
  	        rsC.Open SQLStmtC
           Next
  	     
  	        
  	        
  	        Set SQLStmtSC = Server.CreateObject("ADODB.Command")
  	        Set rsSC = Server.CreateObject ("ADODB.Recordset")
  	        
  	        is_active = 0
  	        If Request.Form("is_active") = "on" Then 
  	        is_active = 1
  	        End If
  	        

  	        SQLStmtSC.CommandText = "UPDATE Score_Card_Master SET Form_Name='" & Request("form_name_title") & "',Is_Active=" & is_active & ",VG_Low='" & Request("VG_Low") & "',VG_High='" & Request("VG_High") & "',Adequate_Low='" & Request("Adequate_Low") & "',Adequate_High='" & Request("Adequate_High") & "',NIA_Low='" & Request("NIA_Low") & "',NIA_High='" & Request("NIA_High") & "',Max_Score='" & Request("Max_Score") & "' WHERE Form_Type_ID='"  & Request("form_type_id") & "'" 
  	        SQLStmtSC.CommandType = 1
  	        Set SQLStmtSC.ActiveConnection = conn
  	       ' response.write "SQL = " & SQLStmtSC.CommandText
  	        rsSC.Open SQLStmtSC
  	        
  	        'response.write(Request.Form("is_active"))
  	        
  	        
  	        
  	       Response.Redirect "edit_score_card.asp?form_type_id=" & Request("form_type_id")
  	        
  	      
  	        
  	      
  	     
  	        
  	        end if  
 	            
  	          
            
        
    
    
    
    
    
    %>
    
 <script type="text/javascript">

     function submitCategory() {
        
            document.scorecardForm.page.value = "add_category";
            document.scorecardForm.submit();
        }

        function saveCatName() {


           // alert(document.scorecardForm.old_cat.value);
            document.scorecardForm.page.value = "edit_cat_name";
            document.scorecardForm.submit();
          // alert("test");
       }
       
	    
    </script>  
<link href="includes/styles.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="includes/jquery-latest.pack.js"></script>
<script type="text/javascript" src="includes/thickbox.js"></script>
<link rel="stylesheet" href="includes/jquery-ui-1.12.1/jquery-ui.min.css" type="text/css" />
<script src="includes/jquery-ui-1.12.1/external/jquery/jquery.js" type="text/javascript"></script>
<script src="includes/jquery-ui-1.12.1/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" src="includes/nav.js"></script>
<script type="text/javascript" src="js/tw-sack.js"></script>
</head>
<body>

<form action="edit_score_card.asp?Form_Type_ID=<%=Request.QueryString("Form_Type_ID")%>" name="scorecardForm" method="POST">
<input type="hidden" name="page" value="add_category" />
<input type="hidden" name="form_name" value="" />
<input type="hidden" name="old_cat" value="" />
<!--#include file="includes/header_client.asp" -->
 <div id="dialog" class="dialog"  title=""></div>
 <script type="text/javascript">



     $(document).ready(function() {

         $("#save_changes").click(function() {

             var VG_High = parseInt($('input[name="VG_High"]').val())
             var VG_Low = parseInt($('input[name="VG_Low"]').val())
             var Adequate_High = parseInt($('input[name="Adequate_High"]').val())
             var Adequate_Low = parseInt($('input[name="Adequate_Low"]').val())
             var NIA_High = parseInt($('input[name="NIA_High"]').val())
             var NIA_Low = parseInt($('input[name="NIA_Low"]').val())

             if (VG_Low > VG_High) {
                 $("#dialog").html('<div>Please correct Final Score Card Grade Range type Very Good. The from value must to less then the to value</div>')
                 $("#dialog").dialog({
                     modal: true,
                     height: 200,
                     width: 300,
                     close: function(event, ui) { },
                     buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            // window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                 });


                 return false
             } else if (Adequate_Low > Adequate_High) {
                 $("#dialog").html('<div>Please correct Final Score Card Grade Range type Adequate. The from value must to less then the to value</div>')
                 $("#dialog").dialog({
                     modal: true,
                     height: 200,
                     width: 300,
                     close: function(event, ui) { },
                     buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            // window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                 });
                 return false
             } else if (NIA_Low > NIA_High) {
                 $("#dialog").html('<div>Please correct Final Score Card Grade Range type Needs Immediate Attention. The from value must to less then the to value</div>')
                 $("#dialog").dialog({
                     modal: true,
                     height: 200,
                     width: 300,
                     close: function(event, ui) { },
                     buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            // window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                 });
                 return false
             } else {

                 document.scorecardForm.page.value = "save_changes";
                 document.scorecardForm.submit();
             }

         });

         $("#add_item").click(function() {
             var item_cat = $('select[name="item_cat"]').val();
             var form_dropdown = $(':radio[name="form_dropdown"]:checked').val();
             var item = $('input[name="item"]').val();
             var form_type_id = $('input[name="form_type_id"]').val();
             var create_form_type = $('#create_form_type').val();
             var create_form_type_text = $('#create_form_type option:selected').text();
             var doc_type = $('select[name="doc_type"]').val();
             var doc_subtype = $('select[name="doc_subtype"]').val();
             //alert($("#item_cat").val());


             //  return false

             if (form_dropdown == 'forms') {

                 if (item_cat == '' && create_form_type == '') {

                     $("#dialog").html('<div>Must select a category and form.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false



                 } else if (item_cat == '' && create_form_type != '') {

                     $("#dialog").html('<div>Must select a category.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false


                 } else if (item_cat != '' && create_form_type == '') {

                     $("#dialog").html('<div>Must select a form.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false

                 } else {
                     document.scorecardForm.page.value = "add_item";
                     document.scorecardForm.form_name.value = create_form_type_text;
                     document.scorecardForm.submit();

                     //alert(document.scorecardForm.form_name.value);

                 }


             } else if (form_dropdown == 'attachments') {
                 //alert(doc_type);
                 //alert(doc_subtype);
                 if (doc_type == '') {

                     $("#dialog").html('<div>Must select an attchament.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false

                 } else if (item_cat == '' && doc_type != '') {

                     $("#dialog").html('<div>Must select an category.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false


                 } else {

                     var doc_subtype = $('select[name="doc_subtype"] option:selected').text()

                     //  alert(doc_subtype);
                     if (doc_subtype == 'All') {
                         create_form_type_text = $('select[name="doc_type"] option:selected').text();

                     } else {
                         create_form_type_text = $('select[name="doc_type"] option:selected').text() + " - " + $('select[name="doc_subtype"] option:selected').text();

                     }


                     document.scorecardForm.page.value = "add_item";
                     document.scorecardForm.form_name.value = create_form_type_text;
                     document.scorecardForm.submit();



                 }
             } else {
                 //   alert(item);
                 //   return false
                 if (item_cat == '' && item == '') {

                     $("#dialog").html('<div>Must select a category and enter an item.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false



                 } else if (item_cat == '' && item != '') {

                     $("#dialog").html('<div>Must select a category.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false


                 } else if (item_cat != '' && item == '') {

                     $("#dialog").html('<div>Must enter and item.')
                     $("#dialog").dialog({
                         modal: true,
                         height: 200,
                         width: 300,
                         close: function(event, ui) { },
                         buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                     });


                     return false




                 } else {
                     document.scorecardForm.page.value = "add_item";
                     document.scorecardForm.form_name.value = item;
                     document.scorecardForm.submit();

                     // alert(document.scorecardForm.form_name.value);

                 };


             };




         });

         $(".submit.cat_delete").click(function() {

             //  document.scorecardForm.page.value = "delete_category";
             // document.scorecardForm.category.value = (this).attr("category");
             // document.scorecardForm.submit();
             // alert($(this).attr("category"));
             window.location.href = "edit_score_card.asp?action=delete_category&category=" + $(this).attr("category") + "&form_type_id=" + $("input[name=form_type_id]").val()


         });





         $(".submit.item_delete").click(function() {


             //alert($("input[name=form_type_id]").val());
             window.location.href = "edit_score_card.asp?action=delete_item&item_id=" + $(this).attr("item_id") + "&Form_Type_ID=" + $("input[name=form_type_id]").val()


         });

         $(".submit.deletecat").click(function() {

             //     alert($(this).attr("category"))
             //  alert($(this).attr("item"))
             window.location.href = "edit_score_card.asp?ID=" + scorecardForm.ID.value + "&Item_Cat=" + $(this).attr("category") + "&Form_Type_ID=" + scorecardForm.Form_Type_ID.value + "&page=delete_score_card_cat";


         });

       
        var item_type = '<%=Item_Type%>'
          
         //  alert(item_type);
         if (item_type == 'attachments') {
                 // alert($('#create_form_type').html());
                 // $('#doc_type').show()
                 $('select[name="doc_type"]').show()
                 $('#doc_subtype_div').show()
                 $('#create_form_type').hide()
                 $('#item').hide()



             } else if (item_type == 'manual') {
                 //  $('#doc_type').hide()
                 $('select[name="doc_type"]').hide()
                 $('#doc_subtype_div').hide()
                 $('#create_form_type').hide()
                 $('#item').show()


             } else {

                 //  $('#doc_type').hide()
                 $('#doc_subtype_div').hide()
                 $('select[name="doc_type"]').hide()
                 $('#create_form_type').show()
                 $('#item').hide()
               }


         $(':radio[name="form_dropdown"]').change(function() {
             var category = $(this).filter(':checked').val();
             if (category == 'attachments') {
                 // alert($('#create_form_type').html());
                 // $('#doc_type').show()
                 $('select[name="doc_type"]').show()
                 $('#doc_subtype_div').show()
                 $('#create_form_type').hide()
                 $('#item').hide()



             } else if (category == 'manual') {
                 //  $('#doc_type').hide()
                 $('select[name="doc_type"]').hide()
                 $('#doc_subtype_div').hide()
                 $('#create_form_type').hide()
                 $('#item').show()


             } else {

                 //  $('#doc_type').hide()
                 $('#doc_subtype_div').hide()
                 $('select[name="doc_type"]').hide()
                 $('#create_form_type').show()
                 $('#item').hide()

             };

             $.ajax({ url: "manage_score_cards.asp?page=update_item_type&item_type=" + category + "&form_type_id=" + $("input[name=form_type_id]").val(), success: function(result) { } });
         });

         $('.item_seq').change(function() {
             var item_seq = $(this).val()
             var item_id = $(this).attr("item_id")

             window.location.href = "edit_score_card.asp?action=edit_item_seq&item_id=" + item_id + "&item_seq=" + item_seq + "&form_type_id=" + $("input[name=form_type_id]").val()
         });

         $('.cat_seq').change(function() {
             var cat_seq = $(this).val()
             var cat_id = $(this).attr("category")

             window.location.href = "edit_score_card.asp?action=edit_cat_seq&cat_id=" + cat_id + "&cat_seq=" + cat_seq + "&form_type_id=" + $("input[name=form_type_id]").val()
         });

         $('input[name="VG_Low"]').change(function() {
             var input = parseInt($(this).val()) + 1

             $(this).val(input)
         });
         $('input[name="VG_High"]').change(function() {
             var VG_High = parseInt($(this).val())
             var VG_Low = parseInt($('input[name="VG_Low"]').val())

             if (VG_Low > VG_High) {
                 $("#dialog").html('<div>The from value must to less the to value</div>')
                 $("#dialog").dialog({
                     modal: true,
                     height: 200,
                     width: 300,
                     close: function(event, ui) { },
                     buttons: [

                    {
                        text: "Close",
                        click: function() {
                            $(this).dialog('close');
                            // window.location.href = "edit_score_card.asp?form_type_id=" + form_type_id
                        }
}]
                 });


                 return false
             }

         });
         //  $('input[name="Adequate_Low"]').change(function() {
         //    var input = parseInt($(this).val()) + 1

         //  $(this).val(input)
         // });
         $('input[name="Max_Score"]').change(function() {
           
             $('input[name="VG_High"]').val($('input[name="Max_Score"]').val());

         });
         $('input[name="Adequate_High"]').change(function() {
             var Adequate_High = parseInt($(this).val())

            // if (Adequate_High != 0) {
                 $('input[name="VG_Low"]').val(Adequate_High + 1);
            // };

         });


         //$('input[name="NIA_Low"]').change(function() {
         // var input = parseInt($(this).val()) + 1

         // $(this).val(input)
         // });

         $('input[name="NIA_High"]').change(function() {
             var NIA_High = parseInt($(this).val())
           //  if (NIA_High != 0) {
                 $('input[name="Adequate_Low"]').val(NIA_High + 1);
            // };


         });

         $('select[name="doc_type"]').change(function() {

             //   alert($(this).val());

             $.ajax({ url: "scripts/ShowDocumentSubTypes.asp?dt=" + $(this).val(), success: function(result) {

                 $("#doc_subtype_div").html(result);

             }
             });


         });


         $('input[name="cat_submit"]').click(function() {

             document.scorecardForm.page.value = "add_category";
             document.scorecardForm.submit();

         });

         $('.submit.cat_edit').click(function() {
             var container = $(this).parent()
             var cat_name = container.find('div').text()
             var old_cat = $(this).attr("category")
             $("input[name=old_cat]").val(old_cat);
             // alert(container.find('div').text());
             container.html("<input type='text' name='edit_cat_name' value='" + cat_name + "' /><input type='button' id='edit_cat_name' onclick='saveCatName();' value='Save' class='submit' /> ")

         });




     });
</script>
 
    <table cellpadding="4" cellspacing="0" border="0" width="969" align="center" id="box">
	    <tr>
	        <td colspan="6"><h1>Edit Score Card:<input style="font-size:14pt;" type="text" name="form_name_title" size="50" value="<%=Form_Name%>" /></h1> </td>
	    </tr>
	    <tr>
	        <td colspan="6" valign="middle"><img src="images/back_arrow.gif" width="17" height="16" border="0" alt="Back" title="Back" /> <a href="manage_score_cards.asp?sf=<%=Form_Name%>">Back to Score Card Search</a></td>
	    </tr>
	    <tr>
	    <td><b>Form Type ID:</b><input type="hidden" name="form_type_id" value="<%=Form_Type_ID%>" />&nbsp;&nbsp;<%=Form_Type_ID%></td>
	    <td><b>Creator:</b>&nbsp;&nbsp;<%=Creator%></td>
	     <td><b>Date of Creation:</b>&nbsp;&nbsp;<%=Date_of_Creation%></td>
	     <td><b>Active:</b>&nbsp;&nbsp;<input type="radio" name="is_active" value="on" <%=on_checked%>>On<input type="radio" name="is_active" value="off" <%=off_checked%>>Off</td>
	    </tr>
	      </table>  
	  <table cellpadding="4" cellspacing="0" border="1" width="969" align="center" id="Table1">
	    <tr>
	        <td colspan="6"><div style="float:left">Final Score Card Grade Range:</div><div style="float:right;margin-right:580px;">Max Score:<input type="text" style="background-color : #d1d1d1;"  onkeypress='return event.charCode >= 48 && event.charCode <= 57' name="Max_Score" size="5" value="<%=Max_Score %>" readonly></div></td>
	    </tr>
	    <tr>
	        <td style="width:200px;" colspan="2" align="center" ><B>VERY GOOD</B></td>
	         <td style="width:200px;" colspan="2" align="center" ><B>ADEQUATE</B></td>
	          <td colspan="2" align="center" ><B>NEEDS IMMEDIATE ATTENTION</B></td>
	    </tr>  
	    <tr>
	        <td style="width:200px;" bgcolor="#008000" colspan="2" ><font color="#FFFFFF">from:</font><input type="text" style="background-color : #d1d1d1;" onkeypress='return event.charCode >= 48 && event.charCode <= 57' name="VG_Low" size="5" value="<%=VG_Low%>" readonly><font color="#FFFFFF">to:</font><input type="text" style="background-color : #d1d1d1;"  name="VG_High" size="5" value="<%=Max_Score%>" readonly></td>
	         <td style="width:200px;" bgcolor="#FFFF00" colspan="2" >from:<input style="background-color : #d1d1d1;"  type="text" onkeypress='return event.charCode >= 48 && event.charCode <= 57' name="Adequate_Low" size="5" value="<%=Adequate_Low%>" readonly>to:<input type="text" name="Adequate_High" size="5" value="<%=Adequate_High%>"></td>
	          <td colspan="2" bgcolor="#FF0000" ><font color="#FFFFFF">from:</font><input type="text" style="background-color : #d1d1d1;" onkeypress='return event.charCode >= 48 && event.charCode <= 57' name="NIA_Low" size="5" value="0" readonly><font color="#FFFFFF">to:</font><input type="text" name="NIA_High" size="5" value="<%=NIA_High%>"></td>
	    </tr>                     
	 </table> 
	 
	    <table cellpadding="4" cellspacing="0" border="1" width="969" align="center" id="Table2">
	    <tr>
	        <td style="width:300px;" nowrap><B>Section:</B><input type="text" name="category" size="50" /><input type="button" name="cat_submit" value="Add" class="submit" /></td>
	        <td style="width:300px;">
	        <select name="item_cat">
               <option value="">Select Section</option>
               
               <%
                Set SQLStmtSC_Ca = Server.CreateObject("ADODB.Command")
  	        Set rsSC_Ca = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmtSC_Ca.CommandText = "SELECT Item_Cat,cat_seq FROM Score_Card_Form_Content where Form_Type_ID='" & Form_Type_ID & "' AND Item='' ORDER BY cat_seq" 
  	        SQLStmtSC_Ca.CommandType = 1
  	        Set SQLStmtSC_Ca.ActiveConnection = conn
  	       ' response.write "SQL = " & SQLStmtSC_Ca.CommandText
  	        rsSC_Ca.Open SQLStmtSC_Ca
  	        
  	        category_counter = 0
  	        
                  Do Until rsSC_Ca.EOF
               
                %>
                <option value="<%=rsSC_Ca("Item_Cat")%>"><%=rsSC_Ca("Item_Cat")%></option>
                
                <%
                 category_counter = category_counter + 1
                 rsSC_Ca.MoveNext
                Loop
                
                 %>
               
               
              
            </select> 
	       <b><div style="padding-top:5px;"><input type="radio" name="form_dropdown" value="forms" <%=forms_checked %>>Forms<input type="radio" name="form_dropdown" value="attachments" <%=attachments_checked %> >Attachments
	       <input type="radio" name="form_dropdown" value="manual" <%=manual_checked %>>Manual</div></b>
	       <div style="padding-top:5px;">
	      <select name="create_form_type" id="create_form_type">
		                            <option value="">Select Form Type to Create</option>
		            <% 
		         form_family_level = "0"
		           Set SQLStmt2 = Server.CreateObject("ADODB.Command")
    	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
        
  	                    SQLStmt2.CommandText = "exec get_creatable_forms_for_score_card '" & Session("user_name") & "','" & Form_Type_ID & "'"
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                    response.write "SQL = " & SQLStmt2.CommandText
  	                    rs2.Open SQLStmt2
  	                    Do Until rs2.EOF
      	                
  	                        cur_required_forms = rs2("Required_Forms")
  	                        cur_form_family = rs2("Form_Family")
      	                    
  	                       
  	                
  	                         %>
  	                           
  	                           
  	                            
  	                            
                                	                    
	                               <%if (cur_form_family = "1" and cur_form_family <> form_family_level) THEN
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
	                            end if      '  form_type = ""
	                                   ' if rs2("Other_Form_Type_ID") <> "" Then 
	                                  '  form_type = "*" & rs2("Form_Description") 
	                                   ' else
	                                   ' form_type = rs2("Form_Description") 
	                                   ' END IF  
	                                   if rs2("Other_Form_Type_ID") <> "" Then        
	                                    %>
	                                   
	                                    <option value="<%=rs2("Form_Type")%>"><% Response.write "*" & rs2("Form_Description") %></option>
	                                    <%
                    	                 else            
	                                    %>
	                                    <option value="<%=rs2("Form_Type")%>"><%Response.write  rs2("Form_Description")%></option>
	                                    <%
	                                    end if
                    	            
	                                                                 	                   
                               rs2.MoveNext
                           Loop%>
                    
                    </select>
            
                 <select name="doc_type" style="display:none;">
              <option value="">--Select Document Type--</option>
	        <option value="FormRelated">Form Related</option>
	        <%
	                Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                Set rs3 = Server.CreateObject ("ADODB.Recordset")
                    
  	                SQLStmt3.CommandText = "exec get_doc_types"
  	                SQLStmt3.CommandType = 1
  	                Set SQLStmt3.ActiveConnection = conn
  	                'response.write "SQL = " & SQLStmt3.CommandText
  	                rs3.Open SQLStmt3
            	            
	                Do Until rs3.EOF
            	    
	                    cur_doc_category = rs3("category")
	                    cur_doc_desc = rs3("desc")
	                %>
	                <option value="<%=cur_doc_category%>" ><%=cur_doc_desc %></option>
	                <%
	                    rs3.MoveNext
            	        
	                Loop
	    %>
	    </select>
	      	    <div name="doc_subtype_div" id="doc_subtype_div" style="display:none;">
    	        <select name="doc_subtype">
    	            <option value="All">All</option>
	            </select>
	          </div>
                  
	       <input type="text"  id="item" name="item" size="50" value="" style="display:none;"/>
	       
	     </div>  <div style="float:right;"><input type="button" id="add_item" value="Add" class="submit" /></div>
	       
	       
	       
	        </td>
	    </tr>
	                     
	 </table> 
	 
	   <table cellpadding="4" cellspacing="0" border="1" width="969" align="center" id="Table3">
	    
               
               <%
               
               cat_counter = 0
               item_counter = 0
              
              if category_counter <> 0 then
                rsSC_Ca.MoveFirst
                  Do Until rsSC_Ca.EOF
               
                %>
               <tr bgcolor="#808080">
               <td colspan="4"><div style="float:left;"><div style="float:left;"><font color="#FFFFFF"><%=rsSC_Ca("Item_Cat")%></font></div>&nbsp;&nbsp;<input type="button" name="cat_edit<%=cat_counter %>" category="<%=rsSC_Ca("Item_Cat")%>" value="Edit" class="submit cat_edit" /> </div>
                     <select style="padding-bottom:2px;" class="cat_seq" category="<%=rsSC_Ca("Item_Cat")%>">
                    <%
                    
                     Set SQLStmtSeq = Server.CreateObject("ADODB.Command")
  	                    Set rsSeq = Server.CreateObject ("ADODB.Recordset")

  	                    SQLStmtSeq.CommandText = "SELECT COUNT(cat_seq) as cat_count FROM Score_Card_Form_Content WHERE Form_Type_ID='" & Form_Type_ID & "' AND Item = ''"
  	                    SQLStmtSeq.CommandType = 1
  	                    Set SQLStmtSeq.ActiveConnection = conn
  	                   ' response.write "SQL = " & SQLStmtSeq.CommandText
  	                    rsSeq.Open SQLStmtSeq
                    
                        For i = 1 To rsSeq("cat_count") 
                     %>
                    <option value="<%=i %>" <%If rsSC_Ca("cat_seq")=i Then Response.Write("selected") End If %>><%=i %></option>
                    
                    <%
                       
                     Next
                    
                    
                     %>
                    
                   </select>
              
               
               
               <div style="float:right;padding-right:22px;"><input type="button" name="cat_delete<%=cat_counter %>" category="<%=rsSC_Ca("Item_Cat")%>" value="Delete" class="submit cat_delete" /></div>
               </td>
                 </tr>
             
                <%
                
                
             Set SQLStmtSC_Item = Server.CreateObject("ADODB.Command")
  	        Set rsSC_Item = Server.CreateObject ("ADODB.Recordset")

  	       ' SQLStmtSC_Item.CommandText = "SELECT DISTINCT Item,Current_Rule,Date_of_Origin_Rule,Other_Form_Type_ID FROM Score_Card_Form_Content WHERE Item <> ''" 
  	       SQLStmtSC_Item.CommandText = "SELECT * FROM Score_Card_Form_Content WHERE Form_Type_ID='" & Form_Type_ID & "' AND Item_Cat='"  & rsSC_Ca("Item_Cat")& "' AND Item <> '' ORDER BY item_seq" 
  	        SQLStmtSC_Item.CommandType = 1
  	        Set SQLStmtSC_Item.ActiveConnection = conn
  	      '  response.write "SQL = " & SQLStmtSC_Item.CommandText
  	        rsSC_Item.Open SQLStmtSC_Item
  	        
  	        
  	        cat_item_count = 0
  	        manual_item = 0
  	        
                  Do Until rsSC_Item.EOF

                    if rsSC_Item("Other_Form_Type_ID") = "" and rsSC_Item("doc_type") = ""  then
                      manual_item = 1
                     else
                     manual_item = 0
                    end if
                  
                  
                    if rsSC_Item("Other_Form_Type_ID") <> "PAPER" then
                      
                         Set SQLStmtDORR = Server.CreateObject("ADODB.Command")
  	                    Set rsDORR = Server.CreateObject ("ADODB.Recordset")

  	                    SQLStmtDORR.CommandText = "exec get_data_of_origin_rule '" & rsSC_Item("Other_Form_Type_ID") & "'"
  	                    SQLStmtDORR.CommandType = 1
  	                    Set SQLStmtDORR.ActiveConnection = conn
  	                   ' response.write "SQL = " & SQLStmtDORR.CommandText
  	                    rsDORR.Open SQLStmtDORR
  	                    
  	                    
  	                    
  	                    dos = rsDORR("dos")
  	                    doi = rsDORR("doi")
  	                    
  	                 end if   
                   
                   if cat_item_count = 0 then
                  
                   %>
                      <tr bgcolor="#F5F5FF"><td></td><td align="center">Date of Origin Rule</td><td  align="center">Current Rule</td><td></td></tr>
                    
                    <%end if %>
                 
                    <tr><td><%=rsSC_Item("Item")%>
                    
                    <input type="hidden" name="hidden_item_id<%=item_counter %>" value="<%=rsSC_Item("ID")%>" />
                    <select class="item_seq" item_id="<%=rsSC_Item("ID")%>">
                    <%
                    
                     Set SQLStmtSeq = Server.CreateObject("ADODB.Command")
  	                    Set rsSeq = Server.CreateObject ("ADODB.Recordset")

  	                    SQLStmtSeq.CommandText = "SELECT COUNT(item_seq) as item_count FROM Score_Card_Form_Content WHERE Form_Type_ID='" & Form_Type_ID & "' AND Item_Cat= '" & rsSC_Ca("Item_Cat")& "' AND Item <> ''"
  	                    SQLStmtSeq.CommandType = 1
  	                    Set SQLStmtSeq.ActiveConnection = conn
  	                  '  response.write "SQL = " & SQLStmtSC_Item.CommandText
  	                    rsSeq.Open SQLStmtSeq
                    
                        For i = 1 To rsSeq("item_count") 
                     %>
                    <option value="<%=i %>" <%If rsSC_Item("item_seq")=i Then Response.Write("selected") End If %>><%=i %></option>
                    
                    <%
                       
                     Next
                    
                    
                     %>
                    
                   </select></td><td align="center">
                 
                    
                    
                   <select name="item_dorr<%=item_counter %>">
                   <% if rsSC_Item("Other_Form_Type_ID") <> "PAPER" and manual_item <> 1 then %>
                        <option value="">--SELECT--</option>
                           <%If doi="1" then %>
                           <option value="DOI" <%If rsSC_Item("Date_Of_Origin_Rule")="DOI" Then Response.Write("selected") End If %>>DOI</option>
                         <%End If %>
                         <%If dos="1" then %>
                         <option value="DOS"  <%If rsSC_Item("Date_Of_Origin_Rule")="DOS" Then Response.Write("selected") End If %>>DOS</option>
                         
                         <%End If %>
                       
                          <option value="Create_Date" <%If rsSC_Item("Date_Of_Origin_Rule")="Create_Date" Then Response.Write("selected") End If %>>Create Date</option>
                    <%elseif manual_item <> 1 then  %>
                         <option value="Action_Date" selected>Action Date</option>
                    <%else %>
                        <option value="">--SELECT--</option>
                    <%end if %>
                    </select> 
                    </td>
                    <td align="center">              
                     <select name="item_cr<%=item_counter %>">
                      <% if manual_item = 1 then %>
                        <option value="">--SELECT--</option>
                      <%else %>
                          <option value="">--SELECT--</option>
                         <option value="Semi_Annual" <%If rsSC_Item("Current_Rule")="Semi_Annual" Then Response.Write("selected") End If %>>Semi Annual</option>
                         <option value="Annual" <%If rsSC_Item("Current_Rule")="Annual" Then Response.Write("selected") End If %>>Annual</option>
                         <option value="BI_Annual" <%If rsSC_Item("Current_Rule")="BI_Annual" Then Response.Write("selected") End If %>>Bi-Annual</option>
                         <option value="Monthly"  <%If rsSC_Item("Current_Rule")="Monthly" Then Response.Write("selected") End If %>>Monthly</option>
                         <option value="One_Time" <%If rsSC_Item("Current_Rule")="One_Time" Then Response.Write("selected") End If %>>One Time</option>
                          <option value="Quarterly" <%If rsSC_Item("Current_Rule")="Quarterly" Then Response.Write("selected") End If %>>Quarterly</option>
                        <%end if %> 
                    </select>  
                    </td>   
                    <td align="center">
                    <input type="button" name="item_delete<%=item_counter %>" value="Delete" class="submit item_delete" item_id="<%=rsSC_Item("ID")%>" /> 
                  </td> 
                </tr> 
                <%
                
                cat_item_count = cat_item_count + 1
                item_counter = item_counter + 1
                
                
                    rsSC_Item.MoveNext
                Loop
                
                cat_counter = cat_counter + 1
                
                 rsSC_Ca.MoveNext
                Loop
                
                end if
                
                 %>
               
               
              
       
	             <input type="hidden" name="total_items" value="<%=item_counter%>" /> 
	                   
	 </table> 
	 
	  <table cellpadding="4" cellspacing="0" border="0" width="969" align="center" id="Table4">
	              <tr><td align="right">
	              <input type="button" name="save_changes" id="save_changes" value="Save Changes" class="submit" />  
	              </td></tr> 
	   </table>               
	
	        
  
   
	           
    
</form>
</body>
</html>
