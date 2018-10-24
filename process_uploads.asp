<%@ Language=VBScript %>
<!--#include file="security_check.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>iCentrix Corp. Electronic Medical Records</title>
    <!--#include file="upload.asp" -->
    <% 
         if Session("user_name") = "pwcard" THEN
           Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
           Set rsDebug = Server.CreateObject ("ADODB.Recordset")
           SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'top of process uploads at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
           SQLStmtDebug.CommandType = 1
           'response.Write "sql = " & SQLStmtNewMd.CommandText
           Set SQLStmtDebug.ActiveConnection = conn
           SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
           rsDebug.Open SQLStmtDebug  
        end if 

        'UPLOAD TO TEMP DIR IF NO TEMP FILE(TF) ALREADY EXISTS
    IF Session("tf") = "" and Request.QueryString("tf") = "" THEN
    
        dim MyUploader
	    Set MyUploader = New FileUploader 
	    MyUploader.Upload()

        thecount= MyUploader.Count() -1

        ' do file first
	    Dim File
	    For each File in MyUploader.Files.Items
            if (File.FileName <> "") Then
                'SAVE THE FILE TO THE TEMP DIR USING THE USER NAME, AND TIMESTAMP
                Orig_name = File.FileName
                'response.write "old name = " & Orig_name
                New_name = Session("user_name") & "_" & Year(Date()) & Month(Date()) & Day(Date()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now())
                'response.write "new name = " & New_name
                File.FileName = New_name & ".xfdl"
                File.SaveToDisk form_root_path & "web_root\temp_forms"
  	        end if
  	    Next
  	    
    ELSE
    
        if Session("tf") <> "" THEN
            New_name = Session("tf")
        else
            New_name = Request.QueryString("tf")
        end if
    END IF
    
    form_FormID = ""
    foundFormID = 0
    form_FormStatus = ""
    foundFormStatus = 0
    foundCopyForm = 0
    foundPreloadEntry = 0
       
    'INSERT INTO XML FOR UNIQUE FORM ID LOGIC
    Set fs = CreateObject("Scripting.FileSystemObject") 
    fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
    Set wfile = fs.OpenTextFile(fileToOpen) 
     	        
    xml_val = wfile.ReadAll
    new_xml_Contents = Replace(xml_val,"'","''")
        
    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	Set rsI = Server.CreateObject ("ADODB.Recordset")
  	SQLStmtI.CommandText = "insert into xml_repository(temp_file_name, xml_string, varchar_string) select '" & New_name & "','" & new_xml_Contents & "','" & new_xml_Contents & "'"
  	SQLStmtI.CommandType = 1
  	Set SQLStmtI.ActiveConnection = conn
  	SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	'if Session("user_name") = "pwcard" THEN
    '  	Response.Write "sql =  " & SQLStmtI.CommandText
  	'end if
  	rsI.Open SQLStmtI
  	
  	Set SQLStmtMod = Server.CreateObject("ADODB.Command")
  	Set rsMod = Server.CreateObject ("ADODB.Recordset")
  	SQLStmtMod.CommandText = "check_modify_flag '" & New_name & "'"
  	SQLStmtMod.CommandType = 1
  	Set SQLStmtMod.ActiveConnection = conn
  	SQLStmtMod.CommandTimeout = 45 'Timeout per Command
  	'if Session("user_name") = "pwcard" THEN
  	'    Response.Write "sql =  " & SQLStmtMod.CommandText
  	'end if
  	rsMod.Open SQLStmtMod
  	
  	modify_flag = rsMod("modify_flag")
    form_exists=0
  	
  	if modify_flag = "on" THEN
  	
  	    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	    SQLStmtI.CommandText = "exec get_hash_and_form_id_for_file '" & New_name & "'"
  	    SQLStmtI.CommandType = 1
  	    Set SQLStmtI.ActiveConnection = conn
  	    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	    'Response.Write "sql =  " & SQLStmtI.CommandText
  	    rsI.Open SQLStmtI

        precheck_form_value = rsI("form_value")
        precheck_mime_value = rsI("mime_value")
        precheck_form_type = rsI("form_type")


        if InStr(precheck_form_type,"STAFF_") <> 0 or InStr(precheck_form_type,"_STAFF") <> 0 THEN
    %>
        <!--#include file="process_uploads_staff.asp" -->
    <%
        elseif InStr(precheck_form_type,"GROUP_") <> 0 or InStr(precheck_form_type,"_GROUP") <> 0 THEN %>

       <!--#include file="process_uploads_group.asp" -->
    <%
        elseif InStr(precheck_form_type,"LOCATION_") <> 0 or InStr(precheck_form_type,"_LOCATION") <> 0 THEN %>

       <!--#include file="process_uploads_location.asp" -->
    <%
        elseif InStr(precheck_form_type,"PROVIDER_") <> 0 or InStr(precheck_form_type,"_PROVIDER") <> 0 THEN 
   '  response.write "test2"
      	'  response.end%>
       <!--#include file="process_uploads_provider.asp" -->

       <% else

         '  response.write precheck_form_type
      	'  response.end

  	        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	        Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	        SQLStmt2.CommandText = "exec does_form_exist " & precheck_form_value
  	        SQLStmt2.CommandType = 1
  	        Set SQLStmt2.ActiveConnection = conn
  	        SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	        'response.Write "sql = " & SQLStmt2.CommandText
  	        rs2.Open SQLStmt2
          	
            'IF FORM ID IS FOUND IN SYSTEM        
  	        IF rs2("form_id_exists") = "Yes" THEN
      	        form_exists=1
  	            'response.Write "yes exists found"
      	           
                'CHECK STATUS OF EXISTING FORM
                Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	            Set rs3 = Server.CreateObject ("ADODB.Recordset")
	            SQLStmt3.CommandText = "exec get_form_info " & precheck_form_value
                SQLStmt3.CommandType = 1
                Set SQLStmt3.ActiveConnection = conn
                SQLStmt3.CommandTimeout = 45 'Timeout per Command
                'response.Write "sql = " & SQLStmt3.CommandText
                rs3.Open SQLStmt3
        	
    	        checked_form_id = rs3("unique_form_id")
    	        checked_form_client = rs3("Client_Id")
    	        checked_form_type = rs3("Form_Type")
        	    
                'IF EXISTING FORM IS FINALIZED       
  	            IF rs3("Status") = "Finalized" THEN
      	        
  	                if precheck_mime_value <> "" THEN
  	                    'CHECK DATA LOCK HASH QUERY GOES HERE
                        Set SQLStmtD = Server.CreateObject("ADODB.Command")
  	                    Set rsD = Server.CreateObject ("ADODB.Recordset")
	                    SQLStmtD.CommandText = "exec check_datalock_hash " & precheck_form_value & ",'" &  precheck_mime_value & "'"
                        SQLStmtD.CommandType = 1
                        Set SQLStmtD.ActiveConnection = conn
                        SQLStmtD.CommandTimeout = 45 'Timeout per Command
                        'if Session("user_name") = "pwcard" THEN
                        '    response.Write "sql = " & SQLStmtD.CommandText
                        'end if 
                        rsD.Open SQLStmtD
                        
                        'IF EXISTING NON EMPTY HASH STILL MATCHES FORM IS SAME SO JUST UPDATE TRANS HISTORY
                        IF rsD("hash_match") <> "Yes" THEN
          	            
                            'FIND NEXT UID
                            Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                        Set rs3 = Server.CreateObject ("ADODB.Recordset")
                            SQLStmt3.CommandText = "exec get_next_uid"
                            SQLStmt3.CommandType = 1
                            Set SQLStmt3.ActiveConnection = conn
                            SQLStmt3.CommandTimeout = 45 'Timeout per Command
                            rs3.Open SQLStmt3
                                          	     
                            form_FormID = rs3("next_uid")
                            old_form_FormID = precheck_form_value
                            'response.Write "MIME NEW ID IS " & form_FormID
                            foundFormID = 1
                            foundCopyForm = 1
                        else
                            foundSameFormHashMatch = 1
                        End if
  	                ELSE
  	                    'OLD FORM WAS FINALIZED THIS ONE IS NOT NEW ID NEEDED
  	                    'FIND NEXT UID
                        Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                    Set rs3 = Server.CreateObject ("ADODB.Recordset")
                        SQLStmt3.CommandText = "exec get_next_uid"
                        SQLStmt3.CommandType = 1
                        Set SQLStmt3.ActiveConnection = conn
                        SQLStmt3.CommandTimeout = 45 'Timeout per Command
                        rs3.Open SQLStmt3
                                          	     
                        form_FormID = rs3("next_uid")
                        old_form_FormID = precheck_form_value
                        'response.Write "NO MIME NEW ID IS " & form_FormID
                        foundFormID = 1
                        foundCopyForm = 1
  	                end if           
               end if
           end if
              
           if foundSameFormHashMatch = 1 THEN 'THIS MEANS THE SAME FORM WAS SUBMITTED WITH NO SIGNED CHANGES ON IT, PIF AND ESP COMPS ALLOW UNSIGNED CHANGES SO SAVE FORM CONTENT ONLY
       
           'response.Write "SAME HASH"
       
                'INSERT INTO TRANS HISTORY
                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "exec insert_transaction_history " & checked_form_id & ",'" & Session("user_name") & "','Edit','" & new_xml_Contents & "'"
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            rsI.Open SQLStmtI
                     
                'UPDATE RECORD IN FORMS MASTER
                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "exec update_existing_form_content " & checked_form_id & ",'" & new_xml_Contents & "','" & Session("user_name") & "'"
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            'response.Write "sql = " & SQLStmtI.CommandText
  	            rsI.Open SQLStmtI
  	        
  	            '***CHECK FORM FOR ALERT RULES
                Set SQLStmtAlertCheck = Server.CreateObject("ADODB.Command")
  	            Set rsAlertCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtAlertCheck.CommandText = "exec alert_rules_check " & checked_form_id
  	            SQLStmtAlertCheck.CommandType = 1
  	            Set SQLStmtAlertCheck.ActiveConnection = conn
  	            'response.Write "sql = " & SQLStmtAlertCheck.CommandText
  	            SQLStmtAlertCheck.CommandTimeout = 45 'Timeout per Command
  	            rsAlertCheck.Open SQLStmtAlertCheck
       
           else
                '****************************************************************************
                    
                'PARSE FORM FOR FORM_TYPE, FORM_ID, CLIENT_ID, SIGNATURES
                Set fs = CreateObject("Scripting.FileSystemObject") 
  	            fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
  	            Set wfile = fs.OpenTextFile(fileToOpen) 
                '**********USE TO LOOP THROUGH FILE TO FIND THINGS (MAY BE DESTRUCTIVE TO SIGNATURES)
                 foundClientName = 0
                foundClientID = 0    
                foundLinkedFormID = 0 
                foundParentFormID = 0 
                foundFormType = 0
                foundDataLock = 0
                foundDHSPStaffID = 0
                foundProgramID = 0
                foundSubmitForm = 0
                foundBackID = 0
                foundFowardID = 0
                foundDeleteForm = 0
                foundFixFlag = 0
                foundFixedID = 0
                foundPersonFirst = 0 
                foundPersonLast = 0
                                        
                form_ClientName = ""
                form_ClientID = ""
                form_LinkedFormID = ""
                form_ParentFormID = ""
                form_FormType = ""
                form_DataLock = ""
                form_ProgramID=""
                form_DHSPStaffID=""
                form_QuestionID=""
                form_GoalID=""
                form_IsSubmit=""
                form_BackID=""
                form_FowardID=""
                fix_FormID=""
                person_firstFormID =""
                person_lastFormID =""
            
                foundAssessedNeed1 = 0
                foundAssessedNeed2 = 0
                foundAssessedNeed3 = 0
                foundAssessedNeed4 = 0
                foundAssessedNeed5 = 0
                foundAssessedNeed6 = 0
                    
                need_cleanup_done = 0
                
                isDataLockDone = 0
            
                update_prescribers = 0 
                            
                temp_DataLock = ""
                            
                do while not wfile.AtEndOfStream 
                        
                    singleline=wfile.readline 
                    'response.Write "line = " & singleline & "<br>"
                                  
                     if InStr(singleline,"<client_name>") and foundClientName = 0 THEN
                
                        foundClientName = 1
                        client_name_tag_start = InStr(singleline,"<client_name>")
                        client_name_tag_end = InStr(singleline,"</client_name>")
                        total_tag_length = client_name_tag_end - client_name_tag_start
                        form_ClientName = Mid(singleline,(client_name_tag_start+13),(total_tag_length-13))
                    
                    end if

                    if InStr(singleline,"<client_name>") and foundClientName = 0 THEN
                
                        foundClientName = 1
                        client_name_tag_start = InStr(singleline,"<client_name>")
                        client_name_tag_end = InStr(singleline,"</client_name>")
                        total_tag_length = client_name_tag_end - client_name_tag_start
                        form_ClientName = Mid(singleline,(client_name_tag_start+13),(total_tag_length-13))

                    
                    end if


                   
                    if InStr(singleline,"<delete_form>1</delete_form>") THEN
                       foundDeleteForm = 1
                    end if

                   if InStr(singleline,"<submit_form>1</submit_form>") THEN
                       foundSubmitForm = 1
                    end if


                     if InStr(singleline,"<goal_num>") THEN
                
                       ' foundSubmitForm = 0
                        goal_id_tag_start = InStr(singleline,"<goal_num>")
                        goal_id_tag_end = InStr(singleline,"</goal_num>")
                        goal_total_tag_length = goal_id_tag_end - goal_id_tag_start
                        form_GoalID = Mid(singleline,(goal_id_tag_start+10),(goal_total_tag_length-10))
         
                    end if
               
                         
         


                    if InStr(singleline,"<client_id>") and foundClientID = 0 THEN
                
                        foundClientID = 1
                        client_id_tag_start = InStr(singleline,"<client_id>")
                        client_id_tag_end = InStr(singleline,"</client_id>")
                        total_tag_length = client_id_tag_end - client_id_tag_start
                        form_ClientID = Mid(singleline,(client_id_tag_start+11),(total_tag_length-11))
                    
                    end if
                    if InStr(singleline,"<form_id>") and foundFormID = 0 THEN
                
                        'response.Write "found form id"
                    
                        foundFormID = 1
                        form_id_tag_start = InStr(singleline,"<form_id>")
                        form_id_tag_end = InStr(singleline,"</form_id>")
                        total_tag_length = form_id_tag_end - form_id_tag_start
                        form_FormID = Mid(singleline,(form_id_tag_start+9),(total_tag_length-9))
                     
                    end if

                


                      if InStr(singleline,"<fixed_id>") and foundFixedID = 0 THEN
               
                        foundFixedID = 1
                        fix_id_tag_start = InStr(singleline,"<fixed_id>")
                        fix_id_tag_end = InStr(singleline,"</fixed_id>")
                        fix_total_tag_length = fix_id_tag_end - fix_id_tag_start
                        fix_FormID = Mid(singleline,(fix_id_tag_start+10),(fix_total_tag_length-10))
                    
                    end if

                      if InStr(singleline,"<fix>1</fix>") THEN
         
                       foundFixFlag=1
                      
                      end if
             


                     if InStr(singleline,"<form_status>") and foundFormStatus = 0 THEN
                
                        'response.Write "found form id"
                    
                        foundFormID = 1
                        form_id_tag_start = InStr(singleline,"<form_status>")
                        form_id_tag_end = InStr(singleline,"</form_status>")
                        total_tag_length = form_id_tag_end - form_id_tag_start
                        form_FormStatus = Mid(singleline,(form_id_tag_start+13),(total_tag_length-13))
                    
                        'response.Write "form id = " & form_FormID
                    
                    end if

                   
                    if InStr(singleline,"<program_id>") and foundProgramID = 0 THEN
                
                        foundProgramID = 1
                        program_id_tag_start = InStr(singleline,"<program_id>")
                        program_id_tag_end = InStr(singleline,"</program_id>")
    
                        total_tag_length = program_id_tag_end - program_id_tag_start
                        form_ProgramID = Mid(singleline,(program_id_tag_start+12),(total_tag_length-12))

                    end if

                   if InStr(singleline,"<form_title>") and foundFormType = 0 THEN
                
                        foundFormType = 1
                        form_name_tag_start = InStr(singleline,"<form_title>")
                        form_name_tag_end = InStr(singleline,"</form_title>")
                        total_tag_length = form_name_tag_end - form_name_tag_start
                        form_FormType = Mid(singleline,(form_name_tag_start+12),(total_tag_length-12))
                    
                    end if  
             

                    if InStr(singleline,"<staff_id>") and foundDHSPStaffID = 0  THEN
                
                        foundDHSPStaffID = 1
                        staff_id_tag_start = InStr(singleline,"<staff_id>")
                        staff_id_tag_end = InStr(singleline,"</staff_id>")
    
                        staff_total_tag_length = staff_id_tag_end - staff_id_tag_start
                        form_DHSPStaffID = Mid(singleline,(staff_id_tag_start+10),(staff_total_tag_length-10))

                    end if


                    if InStr(singleline,"<linked_form_id>") and foundLinkedFormID = 0 THEN
                
                        foundLinkedFormID = 1
                        linked_form_id_tag_start = InStr(singleline,"<linked_form_id>")
                        linked_form_id_tag_end = InStr(singleline,"</linked_form_id>")
                        total_tag_length = linked_form_id_tag_end - linked_form_id_tag_start
                        form_LinkedFormID = Mid(singleline,(linked_form_id_tag_start+16),(total_tag_length-16))
                    
                    end if



              if InStr(singleline,"<next_page>") and foundNextPageVal = 0 THEN
                

                    
                    
                    next_page_tag_start = InStr(singleline,"<next_page>")
                    next_page_tag_end = InStr(singleline,"</next_page>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+11),(total_tag_length-11))


                   if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if
                    
                end if

                if InStr(singleline,"<next_page2>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page2>")
                    next_page_tag_end = InStr(singleline,"</next_page2>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if

             if InStr(singleline,"<next_page3>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page3>")
                    next_page_tag_end = InStr(singleline,"</next_page3>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if




                if InStr(singleline,"<next_page4>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page4>")
                    next_page_tag_end = InStr(singleline,"</next_page4>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if



              if InStr(singleline,"<next_page5>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page5>")
                    next_page_tag_end = InStr(singleline,"</next_page5>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if



               if InStr(singleline,"<next_page6>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page6>")
                    next_page_tag_end = InStr(singleline,"</next_page6>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if



               if InStr(singleline,"<next_page7>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page7>")
                    next_page_tag_end = InStr(singleline,"</next_page7>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if



              if InStr(singleline,"<next_page8>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page8>")
                    next_page_tag_end = InStr(singleline,"</next_page8>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if



             if InStr(singleline,"<next_page9>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page9>")
                    next_page_tag_end = InStr(singleline,"</next_page9>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+12),(total_tag_length-12))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if



          if InStr(singleline,"<next_page10>") and foundNextPageVal = 0 THEN
                
                    foundNextPageVal = 1
                    next_page_tag_start = InStr(singleline,"<next_page10>")
                    next_page_tag_end = InStr(singleline,"</next_page10>")
                    total_tag_length = next_page_tag_end - next_page_tag_start
                    next_pageFormID = Mid(singleline,(next_page_tag_start+13),(total_tag_length-13))

                  if next_pageFormID <>"" then
                      foundNextPageVal = 1
                   else
                     foundNextPageVal = 0
                   end if

                    
                end if


                       if InStr(singleline,"<person_first>") and foundPersonFirst = 0 THEN
                
                        foundPersonFirst = 1
                        person_first_id_tag_start = InStr(singleline,"<person_first>")
                        person_first_id_tag_end = InStr(singleline,"</person_first>")
                        total_tag_length =  person_first_id_tag_end - person_first_id_tag_start
                        person_firstFormID = Mid(singleline,(person_first_id_tag_start+14),(total_tag_length-14))
                    
                    end if

                    if InStr(singleline,"<person_last>") and foundPersonLast = 0 THEN
                
                        foundPersonLast = 1
                        person_last_id_tag_start = InStr(singleline,"<person_last>")
                        person_last_id_tag_end = InStr(singleline,"</person_last>")
                        total_tag_length =  person_last_id_tag_end - person_last_id_tag_start
                        person_lastFormID = Mid(singleline,(person_last_id_tag_start+13),(total_tag_length-13))
                    
                    end if


                    if InStr(singleline,"<parent_form_id>") and foundParentFormID = 0 THEN
                
                        foundParentFormID = 1
                        parent_form_id_tag_start = InStr(singleline,"<parent_form_id>")
                        parent_form_id_tag_end = InStr(singleline,"</parent_form_id>")
                        total_tag_length = parent_form_id_tag_end - parent_form_id_tag_start
                        form_ParentFormID = Mid(singleline,(parent_form_id_tag_start+16),(total_tag_length-16))
                    
                    end if

                 

                    if InStr(singleline,"<back_id>") and foundBackID = 0 THEN
                
                      
                            foundBackID = 1
                            back_id_tag_start = InStr(singleline,"<back_id>")
                            back_id_tag_end = InStr(singleline,"</back_id>")
                            back_id_total_tag_length = back_id_tag_end - back_id_tag_start
                            form_BackID = Mid(singleline,(back_id_tag_start+9),(back_id_total_tag_length-9))

                    
                    end if

                 
                   if InStr(singleline,"<forward_id>") and foundFowardID = 0 THEN
                
                            foundFowardID = 1
                            foward_id_tag_start = InStr(singleline,"<forward_id>")
                            foward_id_tag_end = InStr(singleline,"</forward_id>")
                            foward_id_total_tag_length = foward_id_tag_end - foward_id_tag_start
                            form_FowardID = Mid(singleline,(foward_id_tag_start+12),(foward_id_total_tag_length-12))

                    
                    end if



                      if foundDeleteForm = 1 and form_BackID<>"" and form_FormID<>"" and form_ParentFormID<>"" and fix_FormID="" then

                            Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                        Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmt2.CommandText = "exec delete_dhsp_goal " & form_FormID & "," & form_ParentFormID
  	                        SQLStmt2.CommandType = 1
  	                        Set SQLStmt2.ActiveConnection = conn
  	                        SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                        'response.Write "sql = " & SQLStmt2.CommandText
  	                        rs2.Open SQLStmt2

                           Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_BackID
                    end if



                    if InStr(singleline,"<page_readonly>1</page_readonly>") and foundFixFlag=0 and fix_FormID ="" THEN
                
                       Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_BackID
                       
                    end if 


                    if InStr(singleline,"<next_page>1</next_page>") and fix_FormID ="" THEN
                 

                     if foundFixFlag=0 then
                       Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FowardID
                      else


                         if form_FowardID<>"" then
                              Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                            Set rs3 = Server.CreateObject ("ADODB.Recordset")
	                            SQLStmt3.CommandText = "exec get_form_info " & form_FowardID
                                SQLStmt3.CommandType = 1
                                Set SQLStmt3.ActiveConnection = conn
                                SQLStmt3.CommandTimeout = 45 'Timeout per Command
          
                                rs3.Open SQLStmt3
        	
                                Form_Type = rs3("Form_Type")
                                Linked_Form_ID = rs3("Linked_Form_ID")
                                status = rs3("Status")
    	                        Main_Program = rs3("Main_Program")
    	                        Main_Staff = rs3("Main_Staff")

                               if status="Finalized" then
                                 Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?pathtype=copy&uid=" & form_FowardID & "&if=Fix&pid=" & Main_Program & "&sid=" & Main_Staff & "&ft=" & Form_Type & "&lid=" & Linked_Form_ID & "&cid=" & form_ClientID
                               else
                                    Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FowardID & "&if=Fix"
                               end if
    	       
                            else
                              Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FowardID & "&if=Fix"
                           end if


                       'Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FowardID & "&if=fix"

                      end if
                       
                    end if 


                    if InStr(singleline,"<prev_page>1</prev_page>") THEN
                       
                      
                      if foundFixFlag=0 then
                       Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_BackID
                      else
                       Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_BackID & "&if=Fix"

                      end if

                       
                    end if 


                    if InStr(singleline,"<hcef>1</hcef>") THEN
                
                   
                        form_HCEF = "1"
                    
                    end if  
                    if InStr(singleline,"<form_title>") and foundFormType = 0 THEN
                
                        foundFormType = 1
                        form_name_tag_start = InStr(singleline,"<form_title>")
                        form_name_tag_end = InStr(singleline,"</form_title>")
                        total_tag_length = form_name_tag_end - form_name_tag_start
                        form_FormType = Mid(singleline,(form_name_tag_start+12),(total_tag_length-12))
                    
                    end if                 
                
                    if InStr(singleline,"<prescribing_md>New") <> 0 or InStr(singleline,"<cme_prescribing_md>New") <> 0  THEN
                        update_prescribers = 1   
                    end if 
                
                    if InStr(singleline,"<button sid=""DataLock"">") and foundDataLock = 0 THEN
                
                        foundDataLock = 1                    
                                   
                    end if
                
                    if (InStr(singleline,"<mimedata encoding=""base64-gzip"">") and foundDataLock = 1) or foundDataLock = 2 THEN
                
                        'foundProviderSig = 1 means found the start of Signature tag, 2= found start of mimetag, 3= done with mime tag 
                        if InStr(singleline,"</mimedata>") <> 0 THEN
                    
                            temp_DataLock = temp_DataLock & singleline
                            form_DataLock_tag_start = InStr(temp_DataLock,"<mimedata encoding=""base64-gzip"">")
                            form_DataLock_tag_end = InStr(temp_DataLock,"</mimedata>")
                            total_tag_length = form_DataLock_tag_end - form_DataLock_tag_start
                            form_DataLock = Mid(temp_DataLock,(form_DataLock_tag_start+33),(total_tag_length-11))
                            foundDataLock = 3
                        
                            if form_DataLock <> "" THEN
                        
                                isDataLockDone = 1
                            
                            end if
                        
                        else
                    
                            foundDataLock = 2
                            temp_DataLock = temp_DataLock & singleline
                        
                        end if
                    
                    
                    
                    end if
                
                loop 
            
                wfile.close 
                Set wfile=nothing 
                Set fs=nothing
            
                Set fs = CreateObject("Scripting.FileSystemObject") 
                fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
                Set wfile = fs.OpenTextFile(fileToOpen) 
             	        
                blob_val = wfile.ReadAll
              	        
                new_file_Contents = Replace(blob_val,"'","''")
           
                wfile.close 
                Set wfile=nothing 
                Set fs=nothing
           
                'IF FORM TYPE IS NOT KNOWN OR EMPTY REJECT UPLOAD
                IF form_FormType = "" THEN
            
                    'response.Write "UNKNOWN OR EMPTY FORM TYPE"
                
                END IF
            
                if form_LinkedFormID = "" THEN
                    form_LinkedFormID = -1
                end if
            
                'IF CLIENT ID IS NOT KNOWN OR EMPTY FORCE TO SELECT
                IF form_ClientID = "" AND Session("client_id") = "" and form_FormType <> "CALL_LOG" and form_FormType <> "CALL_LOG_SUMMARY" THEN
            
                    'response.redirect "fix_client_submit.asp?tf=" & Server.URLEncode(New_name)
                
                END IF
            
                IF form_ClientID = "" AND Session("client_id") <> "" THEN
            
                    form_ClientID = Session("client_id")
                
                END IF


              if form_ClientID <> "" and form_ClientID <> "-1" and form_FormID <> "" and form_FormID <> "-1"  THEN
            
                   Set SQLStmtAgencyCheck = Server.CreateObject("ADODB.Command")
  	               Set rsAgencyCheck = Server.CreateObject ("ADODB.Recordset")
  	               SQLStmtAgencyCheck.CommandText = "select case when exists(select 1 from Forms_Master_Preload where Unique_Form_ID=" & form_FormID & " and client_id=" & form_ClientID & ") then 'YES' else 'NO' end as Correct_Agency" 
  	               SQLStmtAgencyCheck.CommandType = 1
  	               Set SQLStmtAgencyCheck.ActiveConnection = conn
  	               SQLStmtAgencyCheck.CommandTimeout = 45 'Timeout per Command
          
           
  	               rsAgencyCheck.Open SQLStmtAgencyCheck
               
                   if rsAgencyCheck("Correct_Agency") = "NO" then



                       if person_firstFormID<>"" and person_lastFormID<>"" then
            
                          Set SQLStmtAgencyCheck2 = Server.CreateObject("ADODB.Command")
  	                       Set rsAgencyCheck2 = Server.CreateObject ("ADODB.Recordset")
  	                       SQLStmtAgencyCheck2.CommandText = "select case when exists(select 1 from Client_Master where Client_ID=" & form_ClientID & " and First_Name='" & person_firstFormID & "' and Last_Name='" & person_lastFormID & "') then 'YES' else 'NO' end as Correct_Agency" 
  	                       SQLStmtAgencyCheck2.CommandType = 1
  	                       Set SQLStmtAgencyCheck2.ActiveConnection = conn
  	                       SQLStmtAgencyCheck2.CommandTimeout = 45 'Timeout per Command
          
           
  	                       rsAgencyCheck2.Open SQLStmtAgencyCheck2
               
                           if rsAgencyCheck2("Correct_Agency") = "NO" then
                                response.write "Refresh dashboard of the Agency that this form came from before submitting."
                                response.end
                           end if
                       else

                               response.write "Refresh dashboard of the Agency that this form came from before submitting."
                                response.end

                       end if

                      
                   end if

              end if      




            
                IF form_FormType = "CALL_LOG_SUMMARY" THEN
                    'INSERT THIS CALL LOG BY UID
                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "insert into forms_master(unique_form_id, form_type, client_id, file_content, create_date, create_user, update_date, update_user, status, main_program, main_staff) select " & form_FormID & ", 'CALL_LOG_SUMMARY', -1, '" & new_file_Contents & "', getDate(), '" & Session("user_name") & "', getDate(), '" & Session("user_name") & "', 'In-Process', -1, -1"
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
  	                
  	                Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "insert into Form_XML_Data(unique_form_id, xml_data_section) select " & form_FormID & ", SUBSTRING(File_Content, (CHARINDEX('<data>',File_Content)+6) , ( CHARINDEX('</data>',File_Content) - CHARINDEX('<data>',File_Content) - 6 ))from forms_master where unique_form_id = " & form_FormID 
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2  
  	            
  	                Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "delete from Forms_Master_Preload where unique_form_id = " & form_FormID 
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
                
                    'update call log table info
                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "update_calls_from_summary " & form_FormID & ",'" & Session("user_name") & "'"
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2        
            
                elseif form_FormType = "CALL_LOG" THEN
                'SPECIAL CASE FOR FORM WITHOUT CLIENT/PROGRAM/STAFF ETC
                
                    'LOOK FOR FORM ID IN SYSTEM
                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "exec does_form_exist " & form_FormID 
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
              	
      	            'IF FORM ID IS FOUND IN SYSTEM        
  	                IF rs2("form_id_exists") = "Yes" THEN  	            
                        'UPDATE THIS CALL LOG BY UID
                        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmt2.CommandText = "update forms_content_master set file_content = '" & new_file_Contents & "', update_date = getDate(), update_user = '" & Session("user_name") & "' where unique_form_id = " & form_FormID 
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                    'response.Write "sql = " & SQLStmt2.CommandText
  	                    rs2.Open SQLStmt2
  	                
  	                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmt2.CommandText = "update Form_XML_Data set xml_data_section = SUBSTRING(File_Content, (CHARINDEX('<data>',File_Content)+6) , ( CHARINDEX('</data>',File_Content) - CHARINDEX('<data>',File_Content) - 6 )) from forms_master f where f.unique_form_id = " & form_FormID & " and form_xml_data.unique_form_id = " & form_FormID
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                    'response.Write "sql = " & SQLStmt2.CommandText
  	                    rs2.Open SQLStmt2
  	                                
                    else
                        'INSERT THIS CALL LOG BY UID
                        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmt2.CommandText = "insert into forms_master(unique_form_id, form_type, client_id, file_content, create_date, create_user, update_date, update_user, status, main_program, main_staff) select " & form_FormID & ", 'CALL_LOG', -1, '" & new_file_Contents & "', getDate(), '" & Session("user_name") & "', getDate(), '" & Session("user_name") & "', 'In-Process', -1, -1"
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                    'response.Write "sql = " & SQLStmt2.CommandText
  	                    rs2.Open SQLStmt2
  	                
  	                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmt2.CommandText = "insert into Form_XML_Data(unique_form_id, xml_data_section) select " & form_FormID & ", SUBSTRING(File_Content, (CHARINDEX('<data>',File_Content)+6) , ( CHARINDEX('</data>',File_Content) - CHARINDEX('<data>',File_Content) - 6 ))from forms_master where unique_form_id = " & form_FormID 
  	                    SQLStmt2.CommandType = 1
  	                    Set SQLStmt2.ActiveConnection = conn
  	                    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                    'response.Write "sql = " & SQLStmt2.CommandText
  	                    rs2.Open SQLStmt2
  	                
                    end if
                
                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "delete from Forms_Master_Preload where unique_form_id = " & form_FormID 
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
                
                    'update call log table info
                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "update_call_log " & form_FormID & ",'" & Session("user_name") & "'"
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
  	                
                elseif form_FormID = "" THEN
                'IF FORM ID IS EMPTY 
            
                    '***NEW FORM FROM LOCAL BLANK FILE***
                     form_Action = "Create"
                 
                    'FIND NEXT UID
                    Set SQLStmtU = Server.CreateObject("ADODB.Command")
  	                Set rsU = Server.CreateObject ("ADODB.Recordset")
                    SQLStmtU.CommandText = "exec get_next_uid"
                    SQLStmtU.CommandType = 1
                    Set SQLStmtU.ActiveConnection = conn
                    SQLStmtU.CommandTimeout = 45 'Timeout per Command
                    rsU.Open SQLStmtU
              	            
                    form_FormID = rsU("next_uid")
                
                    'INSERT INTO TRANS HISTORY
                    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtI.CommandText = "exec insert_transaction_history " & form_FormID & ",'" & Session("user_name") & "','" & form_Action & "',''"
  	                SQLStmtI.CommandType = 1
  	                Set SQLStmtI.ActiveConnection = conn
  	                SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                rsI.Open SQLStmtI
                    
                    'INSERT INTO FORMS MASTER
                    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtI.CommandText = "exec insert_update_forms_master '" & form_FormType & "'," & form_ClientID & "," & form_FormID & ",0,'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Add'," & form_LinkedFormID
  	                SQLStmtI.CommandType = 1
  	                Set SQLStmtI.ActiveConnection = conn
  	                SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                rsI.Open SQLStmtI
                
                    'INSERT INTO FORMS SIG COMPLETED
                    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtI.CommandText = "exec insert_update_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Add'"
  	                SQLStmtI.CommandType = 1
  	                Set SQLStmtI.ActiveConnection = conn
  	                SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                rsI.Open SQLStmtI
          	    
                ELSE
                    'response.Write "form id not empty"
                
                    if foundCopyForm = 1 THEN
                        idToFind = old_form_FormID
                    else
                        idToFind = form_FormID
                    end if
                
                    'LOOK FOR FORM ID IN SYSTEM
                    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	                Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmt2.CommandText = "exec does_form_exist " & idToFind 
  	                SQLStmt2.CommandType = 1
  	                Set SQLStmt2.ActiveConnection = conn
  	                SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmt2.CommandText
  	                rs2.Open SQLStmt2
              	
      	            'IF FORM ID IS FOUND IN SYSTEM        
  	                IF rs2("form_id_exists") = "Yes" THEN
          	    
  	                    'response.Write "exists"
          	    
                        'UPDATE/CHANGE TO FORM
                        form_Action = "Edit"
                    
                        'CHECK STATUS OF EXISTING FORM
                        Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                    Set rs3 = Server.CreateObject ("ADODB.Recordset")
	                    SQLStmt3.CommandText = "exec get_form_info " & idToFind
                        SQLStmt3.CommandType = 1
                        Set SQLStmt3.ActiveConnection = conn
                        SQLStmt3.CommandTimeout = 45 'Timeout per Command
                        'response.Write "sql = " & SQLStmt3.CommandText
                        rs3.Open SQLStmt3
              	    
      	                'IF EXISTING FORM IS FINALIZED       
  	                    IF rs3("Status") = "Finalized" THEN
          	        
                            'CHECK DATA LOCK HASH QUERY GOES HERE
                            Set SQLStmtD = Server.CreateObject("ADODB.Command")
  	                        Set rsD = Server.CreateObject ("ADODB.Recordset")
	                        SQLStmtD.CommandText = "exec check_datalock_hash " & idToFind & ",'" & form_DataLock & "'"
                            SQLStmtD.CommandType = 1
                            Set SQLStmtD.ActiveConnection = conn
                            SQLStmtD.CommandTimeout = 45 'Timeout per Command
                            'response.Write "sql = " & SQLStmtD.CommandText
                            rsD.Open SQLStmtD
                        
                            'IF EXISTING NON EMPTY HASH STILL MATCHES FORM IS SAME SO JUST UPDATE TRANS HISTORY
                            IF rsD("hash_match") = "Yes" THEN
                        
                                'response.Write "FORM IS IDENTICAL SO DO NOTHING"          
                        
                            ELSE
          	            
                                'FIND NEXT UID
                                'Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                            'Set rs3 = Server.CreateObject ("ADODB.Recordset")
                                'SQLStmt3.CommandText = "exec get_next_uid"
                                'SQLStmt3.CommandType = 1
                                'Set SQLStmt3.ActiveConnection = conn
                                'rs3.Open SQLStmt3
                          	
                  	            'old_form_FormID = form_FormID            
                                'form_FormID = rs3("next_uid")
                            
                                new_form_Action = "Create"
                            
                                if form_DataLock <> "" THEN
                                    new_form_child_status = "Finalized"
                                else
                                    new_form_child_status = "In-Process"
                                end if
                            
                                'INSERT INTO TRANS HISTORY FOR NEW UID
                                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtI.CommandText = "exec insert_transaction_history " & form_FormID & ",'" & Session("user_name") & "','" & new_form_Action & "',''"
  	                            SQLStmtI.CommandType = 1
  	                            Set SQLStmtI.ActiveConnection = conn
  	                            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                            rsI.Open SQLStmtI
                            
                                'INSERT INTO TRANS HISTORY FOR OLD UID
                                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtI.CommandText = "exec insert_transaction_history " & old_form_FormID & ",'" & Session("user_name") & "','" & form_Action & "','NA'"
  	                            SQLStmtI.CommandType = 1
  	                            Set SQLStmtI.ActiveConnection = conn
  	                            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                            rsI.Open SQLStmtI
                            
                                'INSERT INTO FORMS MASTER WITH NEW UID AND PARENT ID OF EXISTING FORM
                                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtI.CommandText = "exec insert_update_forms_master '" & form_FormType & "'," & form_ClientID & "," & form_FormID & "," & old_form_FormID & ",'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','" & new_form_child_status & "','Add'," & form_LinkedFormID
  	                            SQLStmtI.CommandType = 1
  	                            'response.Write "sql = " & SQLStmtI.CommandText
  	                            Set SQLStmtI.ActiveConnection = conn
  	                            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                            rsI.Open SQLStmtI
          	                
  	                            'FIX FILE ID IN FILE
                                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtI.CommandText = "update forms_content_master set File_Content = REPLACE(File_Content,'<form_id>" & old_form_FormID & "</form_id>','<form_id>" & form_FormID & "</form_id>') where unique_form_id = " & form_FormID
  	                            SQLStmtI.CommandType = 1
  	                            Set SQLStmtI.ActiveConnection = conn
  	                            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                            rsI.Open SQLStmtI  	                
                            
                                'INSERT INTO FORMS SIG COMPLETE FOR NEW UID
                                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtI.CommandText = "exec insert_update_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Add'"
  	                            SQLStmtI.CommandType = 1
  	                            Set SQLStmtI.ActiveConnection = conn
  	                            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                            rsI.Open SQLStmtI
          	                
                            END IF
                               
                        ELSEIF rs3("Status") = "In-Process" THEN
                     
                         'response.Write "in process old form"
                     
                            'INSERT INTO TRANS HISTORY
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "exec insert_transaction_history " & form_FormID & ",'" & Session("user_name") & "','" & form_Action & "',''"
  	                        SQLStmtI.CommandType = 1
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        rsI.Open SQLStmtI
                     
                            'UPDATE RECORD IN FORMS MASTER
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "exec insert_update_forms_master '" & form_FormType & "'," & form_ClientID & "," & form_FormID & ",0,'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Update'," & form_LinkedFormID
  	                        SQLStmtI.CommandType = 1
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        if Session("user_name") = "pwcard" THEN 
  	                            response.Write "sql = " & SQLStmtI.CommandText
  	                        end if 
  	                        rsI.Open SQLStmtI
                        
                            'IF NEW FORM HAS DATA LOCK HASH UPDATE STATUS TO FINALIZED
                            IF form_DataLock <> "" THEN
                        
                                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtI.CommandText = "update forms_master set status = 'Finalized' where unique_form_id = " & form_FormID
  	                            SQLStmtI.CommandType = 1
  	                            Set SQLStmtI.ActiveConnection = conn
  	                            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                            rsI.Open SQLStmtI


                               '   Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                           ' Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                           ' SQLStmtI.CommandText = "update forms_content_master set File_Content = REPLACE(File_Content,'<is_finalized></is_finalized>','<is_finalized>1</is_finalized>') where unique_form_id = " & form_FormID
  	                          '  SQLStmtI.CommandType = 1
  	                          '  Set SQLStmtI.ActiveConnection = conn
  	                          '  SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                          '  rsI.Open SQLStmtI  
          	                
                            END IF
                                            
                            'UPDATE FORMS SIG COMPLETED
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "exec insert_update_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Update'"
  	                        SQLStmtI.CommandType = 1
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        'response.Write "updating datalock with sql = " & SQLStmtI.CommandText
  	                        rsI.Open SQLStmtI
          	                
                        END IF
                
                    ELSE
                
                        'response.Write "new form with id not in system"
                
                        'NEW FORM
                        form_Action = "Create"
                    
                        'INSERT INTO TRANS HISTORY
                        Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtI.CommandText = "exec insert_transaction_history " & form_FormID & ",'" & Session("user_name") & "','" & form_Action & "',''"
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    rsI.Open SQLStmtI
                    
                        'CHECK FOR PRELOAD EXISTENCE
                        Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtI.CommandText = "select case when EXISTS(select * from forms_master_preload where unique_form_id =" & form_FormID & ") THEN 'Yes' ELSE 'No' End as preload_found" 
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    rsI.Open SQLStmtI
          	        
                        if rsI("preload_found") = "Yes" THEN
                            foundPreloadEntry = 1
                        end if
                    
                        'INSERT INTO FORMS MASTER
                        Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtI.CommandText = "exec insert_update_forms_master '" & form_FormType & "'," & form_ClientID & "," & form_FormID & ",0,'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Add'," & form_LinkedFormID
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    rsI.Open SQLStmtI
                    
                        'INSERT INTO FORMS SIG COMPLETED
                        Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtI.CommandText = "exec insert_update_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Add'"
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    rsI.Open SQLStmtI
          	        
  	                    'IF NEW FORM HAS DATA LOCK CHANGE STATUS TO FINALIZED
  	                     IF form_DataLock <> "" THEN
                        
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "update forms_master set status = 'Finalized' where unique_form_id = " & form_FormID
  	                        SQLStmtI.CommandType = 1
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        rsI.Open SQLStmtI

                         END IF
          	                
                    END IF
                
                END IF
            
           if Session("user_name") = "pwcard" THEN

               Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
               Set rsDebug = Server.CreateObject ("ADODB.Recordset")
               SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'start bottom at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
               SQLStmtDebug.CommandType = 1
               'response.Write "sql = " & SQLStmtNewMd.CommandText
               Set SQLStmtDebug.ActiveConnection = conn
               SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
               rsDebug.Open SQLStmtDebug  
           end if

                if form_FormType = "INTAKE" then    'update_prescribers = 1 or
                   'Check for new prescribing MDs and add to DB
                    Set SQLStmtNewMd = Server.CreateObject("ADODB.Command")
                    Set rsNewMd = Server.CreateObject ("ADODB.Recordset")
                    SQLStmtNewMd.CommandText = "exec insert_prescriber " & form_ClientID & "," & form_FormID
                    SQLStmtNewMd.CommandType = 1
                    'response.Write "sql = " & SQLStmtNewMd.CommandText
                    Set SQLStmtNewMd.ActiveConnection = conn
                    SQLStmtNewMd.CommandTimeout = 45 'Timeout per Command
                    rsNewMd.Open SQLStmtNewMd   
               end if     
              	        
  	            '***CHECK FORM FOR MEDICATIONS UPDATES
                Set SQLStmtMedicationsCheck = Server.CreateObject("ADODB.Command")
  	            Set rsMedicationsCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtMedicationsCheck.CommandText = "exec update_medications_check " & form_FormID
  	            SQLStmtMedicationsCheck.CommandType = 1
  	            Set SQLStmtMedicationsCheck.ActiveConnection = conn
  	            SQLStmtMedicationsCheck.CommandTimeout = 45 'Timeout per Command
  	            rsMedicationsCheck.Open SQLStmtMedicationsCheck
                        
                '***CHECK FORM FOR ALLERGIES UPDATES
                Set SQLStmtAllergiesCheck = Server.CreateObject("ADODB.Command")
  	            Set rsAllergiesCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtAllergiesCheck.CommandText = "exec update_allergies_check " & form_FormID
  	            SQLStmtAllergiesCheck.CommandType = 1
  	            Set SQLStmtAllergiesCheck.ActiveConnection = conn
  	            SQLStmtAllergiesCheck.CommandTimeout = 45 'Timeout per Command
  	           ' Response.Write "sql = " & SQLStmtAllergiesCheck.CommandText
         '  response.end
  	            rsAllergiesCheck.Open SQLStmtAllergiesCheck
  	        
  	            '***CHECK FORM FOR ALLERGIES UPDATES
           if form_FormType = "CONTRACTS" THEN 
                Set SQLStmtContractsCheck = Server.CreateObject("ADODB.Command")
  	            Set rsContractsCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtContractsCheck.CommandText = "exec update_contracts_check " & form_FormID
  	            SQLStmtContractsCheck.CommandType = 1
  	            Set SQLStmtContractsCheck.ActiveConnection = conn
  	            SQLStmtContractsCheck.CommandTimeout = 45 'Timeout per Command
  	            'Response.Write "sql = " & SQLStmtAllergiesCheck.CommandText
  	            rsContractsCheck.Open SQLStmtContractsCheck
  	      end if 
          
           if Session("user_name") = "pwcard" THEN

           Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
           Set rsDebug = Server.CreateObject ("ADODB.Recordset")
           SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'before alerts check at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
           SQLStmtDebug.CommandType = 1
           'response.Write "sql = " & SQLStmtNewMd.CommandText
           Set SQLStmtDebug.ActiveConnection = conn
           SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
           rsDebug.Open SQLStmtDebug  
           end if 
                         
                '***CHECK FORM FOR ALERT RULES
                Set SQLStmtAlertCheck = Server.CreateObject("ADODB.Command")
  	            Set rsAlertCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtAlertCheck.CommandText = "exec alert_rules_check " & form_FormID
  	            SQLStmtAlertCheck.CommandType = 1
  	            Set SQLStmtAlertCheck.ActiveConnection = conn
  	            'response.Write "sql = " & SQLStmtAlertCheck.CommandText
  	            SQLStmtAlertCheck.CommandTimeout = 45 'Timeout per Command
  	            rsAlertCheck.Open SQLStmtAlertCheck
        
           if Session("user_name") = "pwcard" THEN
          Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
           Set rsDebug = Server.CreateObject ("ADODB.Recordset")
           SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'before episodes check at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
           SQLStmtDebug.CommandType = 1
           'response.Write "sql = " & SQLStmtNewMd.CommandText
           Set SQLStmtDebug.ActiveConnection = conn
           SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
           rsDebug.Open SQLStmtDebug  
            end if 
             	
          	    Set SQLStmtEpisodesCheck = Server.CreateObject("ADODB.Command")
  	            Set rsEpisodesCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtEpisodesCheck.CommandText = "exec episode_rules_check " & form_FormID
  	            SQLStmtEpisodesCheck.CommandType = 1
  	            Set SQLStmtEpisodesCheck.ActiveConnection = conn
  	            SQLStmtEpisodesCheck.CommandTimeout = 45 'Timeout per Command
  	            rsEpisodesCheck.Open SQLStmtEpisodesCheck
          	
  	            'Populate meds table from INTAKE and MEDREV forms
                if form_FormType = "DHSAT" THEN
                    Set SQLStmtGoalsObjs = Server.CreateObject("ADODB.Command")
  	                Set rsGoalsObjs = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtGoalsObjs.CommandText = "exec generate_goals_objectives_check " & form_FormID
  	                SQLStmtGoalsObjs.CommandType = 1
  	                Set SQLStmtGoalsObjs.ActiveConnection = conn
  	                SQLStmtGoalsObjs.CommandTimeout = 45 'Timeout per Command
  	                rsGoalsObjs.Open SQLStmtGoalsObjs
  	            end if
           
                  if form_FormType = "HCEF" and form_DataLock <> "" THEN
                    Set SQLStmtHCEF = Server.CreateObject("ADODB.Command")
  	                Set rsHCEF = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtHCEF.CommandText = "exec generate_hcef_visit_dates " & form_FormID
  	                SQLStmtHCEF.CommandType = 1
  	                Set SQLStmtHCEF.ActiveConnection = conn
  	                SQLStmtHCEF.CommandTimeout = 45 'Timeout per Command
  	                rsHCEF.Open SQLStmtHCEF
  	            end if

              if form_FormType = "DHSP_DT" and (form_exists="" or form_exists=0) and fix_FormID ="" THEN
                    Set SQLStmtGoalsObjs = Server.CreateObject("ADODB.Command")
  	                Set rsGoalsObjs = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtGoalsObjs.CommandText = "insert_goal_page_mapping " & form_FormID & "," & form_ParentFormID & "," & form_GoalID
  	                SQLStmtGoalsObjs.CommandType = 1
  	                Set SQLStmtGoalsObjs.ActiveConnection = conn
  	               SQLStmtGoalsObjs.CommandTimeout = 45 'Timeout per Command
  	                rsGoalsObjs.Open SQLStmtGoalsObjs

                  response.write SQLStmtGoalsObjs.CommandText 
  	            end if


  	                if (form_FormType = "DHSP" or form_FormType = "DHSP_DT") and fix_FormID ="" THEN


                       if foundSubmitForm=1 then
                    
                              Set SQLStmtDHSAT = Server.CreateObject("ADODB.Command")
                                Set rsDHSAT = Server.CreateObject("ADODB.Recordset")                
                                SQLStmtDHSAT.CommandText = "exec get_dhsat_specific_goals_objectives  " & form_ClientID
                                SQLStmtDHSAT.CommandType = 1 
                                Set SQLStmtDHSAT.ActiveConnection = conn
                               ' response.Write "sql 1= " & SQLStmtDHSAT.CommandText
                                SQLStmtDHSAT.CommandTimeout = 45 'Timeout per Command
                                rsDHSAT.Open SQLStmtDHSAT
                
                   
                                dhsat_unique_form_id = rsDHSAT("unique_form_id")
                 


                                 if dhsat_unique_form_id<>"" then
                                      Set SQLStmtDHSAT2 = Server.CreateObject("ADODB.Command")
                                        Set rsDHSAT2 = Server.CreateObject("ADODB.Recordset")        
                                           SQLStmtDHSAT2.CommandText = "insert into DHSAT_Goal_Map (unique_form_id,goal_id,dhsat_id) select top 1 " & form_FormID & ",g.Goal_ID," & dhsat_unique_form_id & " from Goals_Master g where  g.Goal_id not in(select Goal_id from DHSAT_Goal_Map where dhsat_id=" & dhsat_unique_form_id & ") and g.dhsp <> 'true' and g.unique_form_id=" & dhsat_unique_form_id & " order by g.Goal_ID asc"
                                        SQLStmtDHSAT2.CommandType = 1 
                                          Set SQLStmtDHSAT2.ActiveConnection = conn
                                        SQLStmtDHSAT2.CommandTimeout = 45 'Timeout per Command
                                        rsDHSAT2.Open SQLStmtDHSAT2
                                 end if

                       end if









           if form_DataLock <> "" and  (form_FormType = "DHSP" or form_FormType = "DHSP_DT") and foundSubmitForm=0 and fix_FormID ="" then
  	             Set SQLStmtDHSP = Server.CreateObject("ADODB.Command")
  	                    Set rsDHSP = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtDHSP.CommandText = "exec get_objective_for_provider_support " & form_ClientID
  	                    SQLStmtDHSP.CommandType = 1
  	                    Set SQLStmtDHSP.ActiveConnection = conn
  	                    SQLStmtDHSP.CommandTimeout = 45 'Timeout per Command
  	                
  	            
  	                    rsDHSP.Open SQLStmtDHSP
      	            
  	                
                       count_forms=0

  	                    Do Until rsDHSP.EOF

                          goal=rsDHSP("goal")
                           service_need=rsDHSP("service_need")
                           objective_name=rsDHSP("objective_name")
                            parent_id=rsDHSP("parent_id")

                             if parent_id <> "" and count_forms=0 and form_FormType = "DHSP_DT" then

                                Set SQLStmtDHSP2 = Server.CreateObject("ADODB.Command")
  	                            Set rsDHSP2 = Server.CreateObject ("ADODB.Recordset")
  	                            SQLStmtDHSP2.CommandText = "exec update_dhsp_form_status_for_fix " & parent_id
  	                            SQLStmtDHSP2.CommandType = 1
  	                            Set SQLStmtDHSP2.ActiveConnection = conn
  	                            SQLStmtDHSP2.CommandTimeout = 45 'Timeout per Command

                                 rsDHSP2.Open SQLStmtDHSP2


                             end if
  	                
  	                
  	                
  	                      Set SQLStmtGoalsObjs = Server.CreateObject("ADODB.Command")
  	                        Set rsGoalsObjs = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtGoalsObjs.CommandText = "exec generate_goals_objectives_check_dhsp " & parent_id & ",'" & Replace(service_need,"'","''") & "','"  &  Replace(goal,"'","''") & "','" &  Replace(objective_name,"'","''") & "'"
  	                        SQLStmtGoalsObjs.CommandType = 1
  	                        Set SQLStmtGoalsObjs.ActiveConnection = conn
  	                        SQLStmtGoalsObjs.CommandTimeout = 45 'Timeout per Command
  	                    
  	                   
  	                        rsGoalsObjs.Open SQLStmtGoalsObjs
  	                  
      	             
                            count_forms=count_forms+1

  	                        rsDHSP.MoveNext  	            
      	            
  	                    Loop
      	      
  	            end if  
           end if
           
           
           
            if form_FormType = "QNR" THEN
  	             Set SQLStmtQNRGoals = Server.CreateObject("ADODB.Command")
  	                    Set rsQNRGoals = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtQNRGoals.CommandText = "exec get_discontinued_goals_qnr " & form_FormID
  	                    SQLStmtQNRGoals.CommandType = 1
  	                    Set SQLStmtQNRGoals.ActiveConnection = conn
  	                    SQLStmtQNRGoals.CommandTimeout = 45 'Timeout per Command
  
  	            
  	                    rsQNRGoals.Open SQLStmtQNRGoals
      	            
  	                
  	                    Do Until rsQNRGoals.EOF
  	                
  	                
  	                
  	                      Set SQLStmtGoalsObjs = Server.CreateObject("ADODB.Command")
  	                        Set rsGoalsObjs = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtGoalsObjs.CommandText = "update Goals_Master set goal_status='Discontinued' where goal_id=" & rsQNRGoals("goal_id")
  	                        SQLStmtGoalsObjs.CommandType = 1
  	                        Set SQLStmtGoalsObjs.ActiveConnection = conn
  	                        SQLStmtGoalsObjs.CommandTimeout = 45 'Timeout per Command
  	                    
  	                   
  	                        rsGoalsObjs.Open SQLStmtGoalsObjs
  
      	             
  	                        rsQNRGoals.MoveNext  	            
      	            
  	                    Loop
      	      
  	            end if 
           
           
           
           
           
            if form_FormType = "DDS_ISP" THEN
  	                    Set SQLStmtDHSP = Server.CreateObject("ADODB.Command")
  	                    Set rsDHSP = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtDHSP.CommandText = "exec get_goals_objectives_from_dds_isp_nonisp " & form_ClientID
  	                    SQLStmtDHSP.CommandType = 1
  	                    Set SQLStmtDHSP.ActiveConnection = conn
  	                    SQLStmtDHSP.CommandTimeout = 45 'Timeout per Command
  	                
  	            
  	                    rsDHSP.Open SQLStmtDHSP
      	            
  	                
  	                    Do Until rsDHSP.EOF
  	                
  	                
  	                
  	                      Set SQLStmtGoalsObjs = Server.CreateObject("ADODB.Command")
  	                        Set rsGoalsObjs = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtGoalsObjs.CommandText = "exec generate_goals_objectives_check_dds_isp " & form_FormID & ",'" & Replace(rsDHSP("service_need"),"'","''") & "','"  & Replace(rsDHSP("goal"),"'","''") & "','" & Replace(rsDHSP("objective"),"'","''") & "'," & rsDHSP("obj_num")
  	                        SQLStmtGoalsObjs.CommandType = 1
  	                        Set SQLStmtGoalsObjs.ActiveConnection = conn
  	                        SQLStmtGoalsObjs.CommandTimeout = 45 'Timeout per Command
  	                    
  	                   
  	                        rsGoalsObjs.Open SQLStmtGoalsObjs
  	                  
      	             
  	                        rsDHSP.MoveNext  	            
      	            
  	                    Loop
      	      
  	            end if  
           
           
            
           
            
    	if Session("user_name") = "pwcard" THEN
           Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
           Set rsDebug = Server.CreateObject ("ADODB.Recordset")
           SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'before attendance check at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
           SQLStmtDebug.CommandType = 1
           'response.Write "sql = " & SQLStmtNewMd.CommandText
           Set SQLStmtDebug.ActiveConnection = conn
           SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
           rsDebug.Open SQLStmtDebug  
        end if 
                      
  	            if form_ClientID <> "-1" THEN
  	                Set SQLStmtAttendanceCheck = Server.CreateObject("ADODB.Command")
  	                Set rsAttendanceCheck = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtAttendanceCheck.CommandText = "exec on_demand_attendance_update " & form_ClientID & "," & form_FormID
  	                SQLStmtAttendanceCheck.CommandType = 1
  	                Set SQLStmtAttendanceCheck.ActiveConnection = conn
  	                SQLStmtAttendanceCheck.CommandTimeout = 150 'Timeout per Command
  	                'Response.Write "sql = " & SQLStmtAttendanceCheck.CommandText
  	                rsAttendanceCheck.Open SQLStmtAttendanceCheck
                end if

           if Session("user_name") = "pwcard" THEN
           Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
           Set rsDebug = Server.CreateObject ("ADODB.Recordset")
           SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'before form counts at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
           SQLStmtDebug.CommandType = 1
           'response.Write "sql = " & SQLStmtNewMd.CommandText
           Set SQLStmtDebug.ActiveConnection = conn
           SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
           rsDebug.Open SQLStmtDebug  
        end if 

                Set SQLStmtEpisodesCheck = Server.CreateObject("ADODB.Command")
  	            Set rsEpisodesCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtEpisodesCheck.CommandText = "exec update_client_form_counts " & form_ClientID
  	            SQLStmtEpisodesCheck.CommandType = 1
  	            Set SQLStmtEpisodesCheck.ActiveConnection = conn
  	            SQLStmtEpisodesCheck.CommandTimeout = 45 'Timeout per Command
  	            rsEpisodesCheck.Open SQLStmtEpisodesCheck

            if Session("user_name") = "pwcard" THEN
           Set SQLStmtDebug = Server.CreateObject("ADODB.Command")
           Set rsDebug = Server.CreateObject ("ADODB.Recordset")
           SQLStmtDebug.CommandText = "insert into debugging_log(debug_string) select 'before form reqs at " & Hour(Now()) & Minute(Now()) & Second(Now()) & "'"
           SQLStmtDebug.CommandType = 1
           'response.Write "sql = " & SQLStmtNewMd.CommandText
           Set SQLStmtDebug.ActiveConnection = conn
           SQLStmtDebug.CommandTimeout = 45 'Timeout per Command
           rsDebug.Open SQLStmtDebug  
           end if 

                Set SQLStmtEpisodesCheck = Server.CreateObject("ADODB.Command")
  	            Set rsEpisodesCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtEpisodesCheck.CommandText = "exec update_client_form_reqs_check " & form_ClientID
  	            SQLStmtEpisodesCheck.CommandType = 1
  	            Set SQLStmtEpisodesCheck.ActiveConnection = conn
  	            SQLStmtEpisodesCheck.CommandTimeout = 90 'Timeout per Command
  	            rsEpisodesCheck.Open SQLStmtEpisodesCheck
            
                'DELETE TEMP FILE
                Set fs=Server.CreateObject("Scripting.FileSystemObject")
                if fs.FileExists(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl") then
                     fs.DeleteFile(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl")
                end if
                set fs=nothing
        
            end if 'END OF CHECK FOR SAME FORM WITH SAME HASH
        
          

         if form_FormType = "MEDREV" and form_HCEF = "1" then
            Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_LinkedFormID
         elseif form_FormType = "AFCMPOC" and foundNextPageVal = 1 and next_pageFormID <>""  then
            Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FormID & "&np=" & next_pageFormID

         elseif form_GoalID <> "" and (form_FormType = "DHSP") and foundSubmitForm = 1 and foundFixFlag=0  then

            Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?ft=DHSP_DT&pid=" & form_ProgramID & "&sid=" & form_DHSPStaffID & "&cid=" & form_ClientID & "&lfid=" & form_FormID & "&gn=" & CInt(form_GoalID) + 1 & "&back_id=" & form_FormID
         elseif form_GoalID <> "" and (form_FormType = "DHSP_DT") and foundSubmitForm=1 and foundFixFlag=0  then
            Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?ft=DHSP_DT&pid=" & form_ProgramID & "&sid=" & form_DHSPStaffID & "&cid=" & form_ClientID & "&lfid=" & form_ParentFormID & "&gn=" & CInt(form_GoalID) + 1 & "&back_id=" & form_FormID
         elseif form_GoalID <> "" and (form_FormType = "DHSP_DT" or form_FormType = "DHSP") and foundSubmitForm=0 and foundFixFlag=0 and fix_FormID ="" and form_DataLock = "" then

          
              Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FormID
         elseif foundFixFlag=1 and (form_FormType = "DHSP_DT" or form_FormType = "DHSP") and fix_FormID ="" then
        
                   if form_FowardID<>"" then
                        Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                    Set rs3 = Server.CreateObject ("ADODB.Recordset")
	                    SQLStmt3.CommandText = "exec get_form_info " & form_FormID
                        SQLStmt3.CommandType = 1
                        Set SQLStmt3.ActiveConnection = conn
                        SQLStmt3.CommandTimeout = 45 'Timeout per Command
          
                        rs3.Open SQLStmt3
        	
                        Form_Type = rs3("Form_Type")
                        Linked_Form_ID = rs3("Linked_Form_ID")
                        status = rs3("Status")
    	                Main_Program = rs3("Main_Program")
    	                Main_Staff = rs3("Main_Staff")

                       if status="Finalized" then
                         Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?pathtype=copy&uid=" & form_FormID & "&if=Fix&pid=" & Main_Program & "&sid=" & Main_Staff & "&ft=" & Form_Type & "&lid=" & Linked_Form_ID & "&cid=" & form_ClientID
                       else
                            Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FormID & "&if=Fix"
                       end if
    	       
                    else
                      Response.Redirect "http://" & url_org_name & ":9080/samples/WebformProxy?uid=" & form_FormID & "&if=Fix"
                   end if

         elseif fix_FormID <>"" and form_FormID<>""  and (form_FormType = "DHSP_DT" or form_FormType = "DHSP") then

             if form_BackID <> "" then
               Set SQLStmt4 = Server.CreateObject("ADODB.Command")
  	            Set rs4 = Server.CreateObject("ADODB.Recordset")
	            SQLStmt4.CommandText = "exec update_dhsp_form_status " & fix_FormID & "," & form_FormID & "," & form_BackID
                SQLStmt4.CommandType = 1
                Set SQLStmt4.ActiveConnection = conn
                SQLStmt4.CommandTimeout = 45 'Timeout per Command
             
                rs4.Open SQLStmt4



             else
                 Set SQLStmt4 = Server.CreateObject("ADODB.Command")
  	            Set rs4 = Server.CreateObject("ADODB.Recordset")
	            SQLStmt4.CommandText = "exec update_dhsp_form_status_no_children " & fix_FormID & "," & form_FormID 
                SQLStmt4.CommandType = 1
                Set SQLStmt4.ActiveConnection = conn
                SQLStmt4.CommandTimeout = 45 'Timeout per Command
             
                rs4.Open SQLStmt4

            end if




                        Set SQLStmtDHSP = Server.CreateObject("ADODB.Command")
  	                    Set rsDHSP = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtDHSP.CommandText = "exec get_objective_for_provider_support " & form_ClientID
  	                    SQLStmtDHSP.CommandType = 1
  	                    Set SQLStmtDHSP.ActiveConnection = conn
  	                    SQLStmtDHSP.CommandTimeout = 45 'Timeout per Command
  	                
  	            
  	                    rsDHSP.Open SQLStmtDHSP
      	            
  	                
  	                    Do Until rsDHSP.EOF
  	                
  	                
  	                
  	                      Set SQLStmtGoalsObjs = Server.CreateObject("ADODB.Command")
  	                        Set rsGoalsObjs = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtGoalsObjs.CommandText = "exec generate_goals_objectives_check_dhsp " & rsDHSP("parent_id") & ",'" & Replace(rsDHSP("service_need"),"'","''") & "','"  &  Replace(rsDHSP("goal"),"'","''") & "','" &  Replace(rsDHSP("objective_name"),"'","''") & "'"
  	                        SQLStmtGoalsObjs.CommandType = 1
  	                        Set SQLStmtGoalsObjs.ActiveConnection = conn
  	                        SQLStmtGoalsObjs.CommandTimeout = 45 'Timeout per Command
  	                    
  	                   
  	                        rsGoalsObjs.Open SQLStmtGoalsObjs
  	                  
      	             
  	                        rsDHSP.MoveNext  	            
      	            
  	                    Loop





         end if
    end if 


           
  
%>

<script type="text/javascript">


  
    if('<%=Request.QueryString("tf")%>' != '')

    {

        if(opener && typeof opener.document != 'undefined')
        {
          

	        // exists
	        if('<%=form_ClientID%>' != '' && '<%=form_ClientID%>' != '-1')
	        {               
 	            if('<%=form_FormType%>' == 'MEDREV' && '<%=form_HCEF%>' != '1')
	            {
	                window.opener.location.href = "view_client_meds.asp?cid=" + '<%=form_ClientID%>';
	            }
               else if('<%=form_FormType%>' == 'HCEF' )
	            {
                  window.opener.location.href = "login_refresh2.asp?cid=" + '<%=form_ClientID%>';
                
	               
	            }
              else
	            {
                
           
                                      if ('<%=form_FormType%>' == 'ISP' && '<%=form_FormStatus %>' == '2') {
                                          window.opener.location.href = 'http://' + <%=url_org_name %> + ':9080/samples/WebformProxy?uid=' + '<%=form_FormID%>';
                                          window.open('','_self');
                                          window.close();
	                  
                                        } else {

                                                              if ('<%=form_DataLock  %>' != '') {
                                                                 if('<%=form_FormType%>' == 'HCEF' )
	                                                                  {
                                                                       window.opener.location.href = "login_refresh2.asp?cid=" + '<%=form_ClientID%>';
	                                                                 } else {
                                                                     

                                                                      <%  if InStr(precheck_form_type,"LOCATION_") <> 0 or InStr(precheck_form_type,"_LOCATION") <> 0 THEN %>
       
                                                                        window.opener.location.href = "index_locations.asp?lid=" + '<%=form_LocationID%>';

                                                                       <%  elseif InStr(precheck_form_type,"STAFF_") <> 0 or InStr(precheck_form_type,"_STAFF") <> 0 THEN %>
       
                                                                        window.opener.location.reload();
                                                                        
                                                                     <% else %>
                                                                          window.opener.location.href = "index.asp?cid=" + '<%=form_ClientID%>';
                                                                   <% end if %>

                                                                      }
                                                              

                                                               }   
                                                                   else {


                                                                 if (opener.document == '[object Document]') {
                                                               
                                                                    window.close();
                                                                 } else {
                                              
                                                                    <%  if InStr(precheck_form_type,"LOCATION_") <> 0 or InStr(precheck_form_type,"_LOCATION") <> 0 THEN %>
       
                                                                        window.opener.location.href = "index_locations.asp?lid=" + '<%=form_LocationID%>';
                                                                    <%  elseif InStr(precheck_form_type,"STAFF_") <> 0 or InStr(precheck_form_type,"_STAFF") <> 0 THEN %>
       
                                                                       window.opener.location.reload();

                                                                     <% else %>
                                                                        window.opener.location.href = "index.asp?cid=" + '<%=form_ClientID%>';
                                                                   <% end if %>

                                                                  

                                                               }
                                                                 
                                                               

                                                              }
                     
                     
                                        }
               
                  

                  

	            }
	        }
	        else
	        {


	            if('<%=form_FormType%>' == 'AF' || '<%=form_FormType%>' == 'AFS' || '<%=form_FormType%>' == 'CONTRACTS')
	            {
	                 window.opener.location.reload();
	            }
	             /*
                else if('<%=checked_form_client%>' == '' || '<%=checked_form_client%>' == '-1')
	            {
	                window.opener.location.reload();
	            }
                */
	            else
	            {

                   <%  if InStr(precheck_form_type,"LOCATION_") <> 0 or InStr(precheck_form_type,"_LOCATION") <> 0 THEN %>
       
                    window.opener.location.href = "index_locations.asp?lid=" + '<%=form_LocationID%>';

                  
                   <%  elseif InStr(precheck_form_type,"STAFF_") <> 0 or InStr(precheck_form_type,"_STAFF") <> 0 THEN %>
       
                       window.opener.location.reload();
                                                                        
                   <% else %>
                     window.opener.location.href = "index.asp?cid=" + '<%=checked_form_client%>';
                  <% end if %>


	              
             
	            }
	            
	        
	        }   

         //  alert("test2");
	        window.close();
        }
        else
        {
            window.open('','_self');
            window.close();
        }                 
    }
    else
    {

           
        window.location.href = "index.asp?cid=" + '<%=form_ClientID%>';
    }
</script>

<%
    else

      '  response.Write "do not submit from trans history"
  
    %>
    <script type="text/javascript">
    window.close();
    </script>
    <%
    end if
%>
</head>
<body>

</body>
</html>