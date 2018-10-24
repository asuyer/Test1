
<% 






    
               if precheck_form_value <> "" and form_FormID <> "" and form_FormID <> "-1" THEN
            
                   Set SQLStmtAgencyCheck = Server.CreateObject("ADODB.Command")
  	               Set rsAgencyCheck = Server.CreateObject ("ADODB.Recordset")
  	               SQLStmtAgencyCheck.CommandText = "select case when exists(select 1 from Forms_Master_Preload where Unique_Form_ID=" & form_FormID & " and Form_Type=" & precheck_form_value & ") then 'YES' else 'NO' end as Correct_Agency" 
  	               SQLStmtAgencyCheck.CommandType = 1
  	               Set SQLStmtAgencyCheck.ActiveConnection = conn
  	               SQLStmtAgencyCheck.CommandTimeout = 45 'Timeout per Command
          
           
  	               rsAgencyCheck.Open SQLStmtAgencyCheck
               
                   if rsAgencyCheck("Correct_Agency") = "NO" then
                        response.write "Refresh dashboard of the Agency that this form came from before submitting."
                        response.end
                   end if

              end if 




 Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	    SQLStmt2.CommandText = "exec does_staff_form_exist " & rsI("form_value")
  	    SQLStmt2.CommandType = 1
  	    Set SQLStmt2.ActiveConnection = conn
  	    SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	    'response.Write "sql = " & SQLStmt2.CommandText
  	    rs2.Open SQLStmt2
          	
        'IF FORM ID IS FOUND IN SYSTEM        
  	    IF rs2("form_id_exists") = "Yes" THEN
      
      	    
  	        'response.Write "yes exists found"
      	           
            'CHECK STATUS OF EXISTING FORM
            Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	        Set rs3 = Server.CreateObject ("ADODB.Recordset")
	        SQLStmt3.CommandText = "exec get_group_form_info " & rsI("form_value")
            SQLStmt3.CommandType = 1
            Set SQLStmt3.ActiveConnection = conn
            SQLStmt3.CommandTimeout = 45 'Timeout per Command
            'response.Write "sql = " & SQLStmt3.CommandText
            rs3.Open SQLStmt3
        	
    	    checked_form_id = rs3("unique_form_id")
    	    checked_form_group = rs3("Group_Id")
        	    
            'IF EXISTING FORM IS FINALIZED       
  	        IF rs3("Status") = "Finalized" THEN
      	        
  	            if rsI("mime_value") <> "" and rsI("mime_value") <> "NONE" THEN
  	                'CHECK DATA LOCK HASH QUERY GOES HERE
                    Set SQLStmtD = Server.CreateObject("ADODB.Command")
  	                Set rsD = Server.CreateObject ("ADODB.Recordset")
	                SQLStmtD.CommandText = "exec check_datalock_hash " & rsI("form_value") & ",'" &  rsI("mime_value") & "'"
                    SQLStmtD.CommandType = 1
                    Set SQLStmtD.ActiveConnection = conn
                    SQLStmtD.CommandTimeout = 45 'Timeout per Command
                    'response.Write "sql = " & SQLStmtD.CommandText
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
                        old_form_FormID = rsI("form_value") 
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
                    old_form_FormID = rsI("form_value") 
                    'response.Write "NO MIME NEW ID IS " & form_FormID
                    foundFormID = 1
                    foundCopyForm = 1
  	            end if           
           end if
       end if
       
       if foundSameFormHashMatch = 1 THEN 'THIS MEANS THE SAME FORM WAS SUBMITTED WITH NO SIGNED CHANGES ON IT, PIF AND ESP COMPS ALLOW UNSIGNED CHANGES SO SAVE FORM CONTENT ONLY
       
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
  	        SQLStmtI.CommandText = "exec update_existing_group_form_content " & checked_form_id & ",'" & new_xml_Contents & "','" & Session("user_name") & "'"
  	        SQLStmtI.CommandType = 1
  	        Set SQLStmtI.ActiveConnection = conn
  	        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	        rsI.Open SQLStmtI
  	              
       else
            '****************************************************************************
                    
            'PARSE FORM FOR FORM_TYPE, FORM_ID, SIGNATURES
            Set fs = CreateObject("Scripting.FileSystemObject") 
  	        fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
  	        Set wfile = fs.OpenTextFile(fileToOpen) 
            '**********USE TO LOOP THROUGH FILE TO FIND THINGS (MAY BE DESTRUCTIVE TO SIGNATURES)
            foundClientName = 0
            foundClientID = 0    
            foundLinkedFormID = 0 
            foundFormType = 0
            foundDataLock = 0
                                        
            form_ClientName = ""
            form_GroupID = ""
            form_LinkedFormID = ""
            form_FormType = ""
            form_DataLock = ""
            
            foundAssessedNeed1 = 0
            foundAssessedNeed2 = 0
            foundAssessedNeed3 = 0
            foundAssessedNeed4 = 0
            foundAssessedNeed5 = 0
            foundAssessedNeed6 = 0
                    
            need_cleanup_done = 0
              
            update_prescribers = 0 
             
            isDataLockDone = 0
                            
            temp_DataLock = ""
            
            individuals_count = 0
            
            individuals_section_loc_code = ""
            individuals_section_date_of_service = ""
            cur_drill_end = ""
            cur_total_time = ""
                            
            do while not wfile.AtEndOfStream 
                        
                singleline=wfile.readline 
                'response.Write "line = " & singleline & "<br>"
                                           
                if InStr(singleline,"<form_id>") and foundFormID = 0 THEN
                                    
                    foundFormID = 1
                    form_id_tag_start = InStr(singleline,"<form_id>")
                    form_id_tag_end = InStr(singleline,"</form_id>")
                    total_tag_length = form_id_tag_end - form_id_tag_start
                    form_FormID = Mid(singleline,(form_id_tag_start+9),(total_tag_length-9))
                                        
                end if
                
                if InStr(singleline,"<group_id>") and foundGroupID = 0 THEN
                                    
                    foundGroupID = 1
                    form_id_tag_start = InStr(singleline,"<group_id>")
                    form_id_tag_end = InStr(singleline,"</group_id>")
                    total_tag_length = form_id_tag_end - form_id_tag_start
                    form_GroupID = Mid(singleline,(form_id_tag_start+10),(total_tag_length-10))
                                        
                end if
                
                if InStr(singleline,"<loc_code>") and individuals_section_loc_code = "" THEN
                                    
                    form_id_tag_start = InStr(singleline,"<loc_code>")
                    form_id_tag_end = InStr(singleline,"</loc_code>")
                    total_tag_length = form_id_tag_end - form_id_tag_start
                    individuals_section_loc_code = Mid(singleline,(form_id_tag_start+10),(total_tag_length-10))
                                        
                end if
                
                if InStr(singleline,"<date_of_service>") and individuals_section_date_of_service = "" THEN
                                    
                    form_id_tag_start = InStr(singleline,"<date_of_service>")
                    form_id_tag_end = InStr(singleline,"</date_of_service>")
                    total_tag_length = form_id_tag_end - form_id_tag_start
                    individuals_section_date_of_service = Mid(singleline,(form_id_tag_start+17),(total_tag_length-17))
                                        
                end if
                

                 if InStr(singleline,"<drill_end>") <> 0 and cur_drill_end = "" THEN
                        form_id_tag_start = InStr(singleline,"<drill_end>")
                        form_id_tag_end = InStr(singleline,"</drill_end>")
                        total_tag_length = form_id_tag_end - form_id_tag_start
                        cur_drill_end = Mid(singleline,(form_id_tag_start+11),(total_tag_length-11))
                  end if    

                  if InStr(singleline,"<total_time>") <> 0 and cur_total_time = "" THEN
                        form_id_tag_start = InStr(singleline,"<total_time>")
                        form_id_tag_end = InStr(singleline,"</total_time>")
                        total_tag_length = form_id_tag_end - form_id_tag_start
                        cur_total_time = Mid(singleline,(form_id_tag_start+12),(total_tag_length-12))
                  end if 

                if InStr(singleline,"<linked_form_id>") and foundLinkedFormID = 0 THEN
                
                    foundLinkedFormID = 1
                    linked_form_id_tag_start = InStr(singleline,"<linked_form_id>")
                    linked_form_id_tag_end = InStr(singleline,"</linked_form_id>")
                    total_tag_length = linked_form_id_tag_end - linked_form_id_tag_start
                    form_LinkedFormID = Mid(singleline,(linked_form_id_tag_start+16),(total_tag_length-16))
                    
                end if
                
                if InStr(singleline,"<form_title>") and foundFormType = 0 THEN
                
                    foundFormType = 1
                    form_name_tag_start = InStr(singleline,"<form_title>")
                    form_name_tag_end = InStr(singleline,"</form_title>")
                    total_tag_length = form_name_tag_end - form_name_tag_start
                    form_FormType = Mid(singleline,(form_name_tag_start+12),(total_tag_length-12))
                    
                end if
            
                if InStr(singleline,"<button sid=""DataLock"">") and foundDataLock = 0 THEN
                
                    foundDataLock = 1
                    form_DataLock = temp_mime_lock
                    
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
                
                if InStr(singleline,"<client_id>") <> 0 and InStr(singleline, "<client_id></client_id>") = 0 THEN
                    individuals_count = individuals_count + 1
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
            
            if form_FormID = "" THEN
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
  	            SQLStmtI.CommandText = "exec insert_update_group_forms_master '" & form_FormType & "'," & form_GroupID & "," & form_FormID & ",0,'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Add'," & form_LinkedFormID
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            rsI.Open SQLStmtI
                
                'INSERT INTO FORMS SIG COMPLETED
                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "exec insert_update_group_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Add'"
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            rsI.Open SQLStmtI
          	    
            ELSE                
                if foundCopyForm = 1 THEN
                    idToFind = old_form_FormID
                else
                    idToFind = form_FormID
                end if
                
                'LOOK FOR FORM ID IN SYSTEM
                Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	            Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmt2.CommandText = "exec does_group_form_exist " & idToFind 
  	            SQLStmt2.CommandType = 1
  	            Set SQLStmt2.ActiveConnection = conn
  	            SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	            'response.Write "sql = " & SQLStmt2.CommandText
  	            rs2.Open SQLStmt2
              	
      	        'IF FORM ID IS FOUND IN SYSTEM        
  	            IF rs2("form_id_exists") = "Yes" THEN
          	              	    
                    'UPDATE/CHANGE TO FORM
                    form_Action = "Edit"
                    
                    'CHECK STATUS OF EXISTING FORM
                    Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	                Set rs3 = Server.CreateObject ("ADODB.Recordset")
	                SQLStmt3.CommandText = "exec get_group_form_info " & idToFind
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
  	                        SQLStmtI.CommandText = "exec insert_update_group_forms_master '" & form_FormType & "'," & form_GroupID & "," & form_FormID & "," & old_form_FormID & ",'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','" & new_form_child_status & "','Add'," & form_LinkedFormID
  	                        SQLStmtI.CommandType = 1
  	                        'response.Write "sql = " & SQLStmtI.CommandText
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        rsI.Open SQLStmtI
          	                
  	                        'FIX FILE ID IN FILE
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "update group_forms_master set File_Content = REPLACE(File_Content,'<form_id>" & old_form_FormID & "</form_id>','<form_id>" & form_FormID & "</form_id>') where unique_form_id = " & form_FormID
  	                        SQLStmtI.CommandType = 1
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        rsI.Open SQLStmtI  	                
                            
                            'INSERT INTO FORMS SIG COMPLETE FOR NEW UID
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "exec insert_update_group_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Add'"
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
  	                    SQLStmtI.CommandText = "exec insert_update_group_forms_master '" & form_FormType & "'," & form_GroupID & "," & form_FormID & ",0,'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Update'," & form_LinkedFormID
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    rsI.Open SQLStmtI
                        
                        'IF NEW FORM HAS DATA LOCK HASH UPDATE STATUS TO FINALIZED
                        IF form_DataLock <> "" THEN
                        
                            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                        Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                        SQLStmtI.CommandText = "update group_forms_master set status = 'Finalized' where unique_form_id = " & form_FormID
  	                        SQLStmtI.CommandType = 1
  	                        Set SQLStmtI.ActiveConnection = conn
  	                        SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                        rsI.Open SQLStmtI
          	                
                        END IF
                                            
                        'UPDATE FORMS SIG COMPLETED
                        Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtI.CommandText = "exec insert_update_group_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Update'"
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    'response.Write "updating datalock with sql = " & SQLStmtI.CommandText
  	                    rsI.Open SQLStmtI
          	                
                    END IF
                
                ELSE
                                
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
  	                SQLStmtI.CommandText = "select case when EXISTS(select 1 from forms_master_preload where unique_form_id =" & form_FormID & ") THEN 'Yes' ELSE 'No' End as preload_found" 
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
  	                SQLStmtI.CommandText = "exec insert_update_group_forms_master '" & form_FormType & "'," & form_GroupID & "," & form_FormID & ",0,'" & new_file_Contents & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Add'," & form_LinkedFormID
  	                SQLStmtI.CommandType = 1
  	                Set SQLStmtI.ActiveConnection = conn
  	                SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                'response.Write "sql = " & SQLStmtI.CommandText
  	                rsI.Open SQLStmtI
                    
                    'INSERT INTO FORMS SIG COMPLETED
                    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtI.CommandText = "exec insert_update_group_signatures_completed_datalock " & form_FormID & ",'" & form_DataLock & "', 'Add'"
  	                SQLStmtI.CommandType = 1
  	                Set SQLStmtI.ActiveConnection = conn
  	                SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                rsI.Open SQLStmtI
          	        
  	                'IF NEW FORM HAS DATA LOCK CHANGE STATUS TO FINALIZED
  	                  IF form_DataLock <> "" THEN
                        
                        Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	                    Set rsI = Server.CreateObject ("ADODB.Recordset")
  	                    SQLStmtI.CommandText = "update group_forms_master set status = 'Finalized' where unique_form_id = " & form_FormID
  	                    SQLStmtI.CommandType = 1
  	                    Set SQLStmtI.ActiveConnection = conn
  	                    SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	                    rsI.Open SQLStmtI
          	                
                     END IF
          	                
                END IF
                
            END IF
                      
            
            'individuals_count = non-empty client ids found above
            
            '********************************
            '********************************
            '********************************
            'NOW GENERATE INDIVIDUAL FORMS
            individuals_form_count = 0        
            individual_forms_generated = 0 
            
            'response.Write "there are " & individuals_count & " clients to generate<br/>"
            'DO A LOOP/FORM FOR EACH INDIVIDUAL FOUND ABOVE WITH NON EMPTY CLIENT ID
            Do while individuals_form_count < individuals_count
            
                new_form_string = ""                
                individual_generated = 0
                found_sig_item_end = 0
                individuals_skipped = 0
                
                'FIND NEXT UID
                Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	            Set rs3 = Server.CreateObject ("ADODB.Recordset")
                SQLStmt3.CommandText = "exec get_next_uid"
                SQLStmt3.CommandType = 1
                Set SQLStmt3.ActiveConnection = conn
                SQLStmt3.CommandTimeout = 45 'Timeout per Command
                rs3.Open SQLStmt3
                                          	     
                individual_FormID = rs3("next_uid")
            
                'PARSE FORM FOR FORM_TYPE, FORM_ID, SIGNATURES
                Set fs = CreateObject("Scripting.FileSystemObject") 
  	            fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
  	            Set wfile = fs.OpenTextFile(fileToOpen) 
                '**********USE TO LOOP THROUGH FILE TO FIND THINGS (MAY BE DESTRUCTIVE TO SIGNATURES)
                                            
                do while not wfile.AtEndOfStream 
                            
                    singleline=wfile.readline 
                    'response.Write "line = " & singleline & "<br>"
                    

                
            
                 

                    if InStr(singleline,"<individual_person>") <> 0 and ( (individual_forms_generated = 0 and individuals_form_count = 0 and individual_generated <> 1) or (individuals_skipped = individual_forms_generated and individual_generated <> 1 ) ) THEN       
                        'KEEP THIS GUY IN THE STRING
                        'response.Write "1: individual_forms_generated = " & individual_forms_generated & ", individuals_form_count = " & individuals_form_count & ", individual_generated = " & individual_generated & ", individuals_skipped " & individuals_skipped & "<br />"
                        
                        new_form_string = new_form_string & singleline & chr(13) & chr(10) 
                        
                        individual_generated = 1
                        individual_forms_generated = individual_forms_generated + 1
                                              
                        
                    elseif InStr(singleline,"<individual_person>") <> 0 THEN 
                        
                        'response.Write "2: individual_forms_generated = " & individual_forms_generated & ", individuals_form_count = " & individuals_form_count & ", individual_generated = " & individual_generated & ", individuals_skipped " & individuals_skipped & "<br />"
                        
                        individuals_skipped = individuals_skipped + 1
                                              
                        Do Until InStr(singleline,"</individual_person>") <> 0 
                            singleline=wfile.readline 
                        Loop
                        
                        'write out end of individuals section table
                        'new_form_string = new_form_string & singleline & chr(13) & chr(10)     
                    
                    elseif InStr(singleline,"<client_id>") <> 0 and InStr(singleline, "<client_id></client_id>") = 0 THEN
                        form_id_tag_start = InStr(singleline,"<client_id>")
                        form_id_tag_end = InStr(singleline,"</client_id>")
                        total_tag_length = form_id_tag_end - form_id_tag_start
                        cur_individual_client_id = Mid(singleline,(form_id_tag_start+11),(total_tag_length-11))
                        
                        new_form_string = new_form_string & singleline & chr(13) & chr(10)                         
                    
                    elseif InStr(singleline, "<form_id>") <> 0 THEN
                                                
                        new_form_string = new_form_string & "<form_id>" & individual_FormID & "</form_id>" & chr(13) & chr(10) 
                    
                    
                    elseif InStr(singleline, "<loc_code></loc_code>") <> 0 THEN
                                                
                        new_form_string = new_form_string & "<loc_code>" & individuals_section_loc_code & "</loc_code>" & chr(13) & chr(10) 
                    
                    elseif InStr(singleline, "<date_of_service></date_of_service>") <> 0 THEN
                                                
                        new_form_string = new_form_string & "<date_of_service>" & individuals_section_date_of_service & "</date_of_service>" & chr(13) & chr(10)     
                    
                    elseif InStr(singleline, "<drill_end_client>") <> 0 THEN
                                                
                        new_form_string = new_form_string & "<drill_end_client>" & cur_drill_end & "</drill_end_client>" & chr(13) & chr(10) 

                   elseif InStr(singleline, "<total_time_client>") <> 0 THEN
                                                
                        new_form_string = new_form_string & "<total_time_client>" & cur_total_time & "</total_time_client>" & chr(13) & chr(10) 

                    elseif InStr(singleline,"<signature sid=") THEN
            
                        singleline=wfile.readline
                    
                        'FOUND START OF SIGNATURE HASH ITEM LOOP AND DO NOT WRITE TO NEW FILE UNTIL AFTER END IS FOUND
                        Do while found_sig_item_end < 2
                            if InStr(singleline,"</signature>") THEN
                                found_sig_item_end = found_sig_item_end + 1
                            end if
                            
                            if found_sig_item_end <> 2 THEN
                                singleline=wfile.readline
                            end if
                        loop
                
                    else
                        new_form_string = new_form_string & singleline & chr(13) & chr(10)                           
                                    
                    end if
                                                  
                loop
                
                new_form_string = REPLACE(new_form_string, "process_submit_group.aspx", "process_submit.aspx")
                new_form_string = Replace(new_form_string,"'","''")


                
                
                'UPDATE RECORD IN FORMS MASTER
                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "exec insert_update_forms_master '" & form_FormType & "'," & cur_individual_client_id & "," & individual_FormID & ",0,'" & new_form_string & "','" & Session("user_name") & "','" & Session("user_name") & "','In-Process','Add',-1"
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            'Response.Write "sql = " & SQLStmtI.CommandText
  	            rsI.Open SQLStmtI 
  	            
  	            Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "update forms_master set Main_Program = g.Responsible_Program, Main_Staff = g.Organizer from groups_master g where g.group_id =  " & form_GroupID & " and forms_master.unique_form_id = " & individual_FormID
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	          '  Response.Write "sql = " & SQLStmtI.CommandText
   ' response.end
  	            rsI.Open SQLStmtI 



                'INSERT INTO TRANS HISTORY
                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "exec insert_transaction_history " & individual_FormID & ",'" & Session("user_name") & "','Create',''"
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            rsI.Open SQLStmtI


  	            
  	            'INSERT INTO FORMS SIG COMPLETED
                Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	            Set rsI = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtI.CommandText = "exec insert_update_signatures_completed_datalock " & individual_FormID & ",'', 'Add'"
  	            SQLStmtI.CommandType = 1
  	            Set SQLStmtI.ActiveConnection = conn
  	            SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	            rsI.Open SQLStmtI
                
                 Set SQLStmtEpisodesCheck = Server.CreateObject("ADODB.Command")
  	            Set rsEpisodesCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtEpisodesCheck.CommandText = "exec update_client_form_counts " & cur_individual_client_id
  	            SQLStmtEpisodesCheck.CommandType = 1
  	            Set SQLStmtEpisodesCheck.ActiveConnection = conn
  	            SQLStmtEpisodesCheck.CommandTimeout = 45 'Timeout per Command
  	            rsEpisodesCheck.Open SQLStmtEpisodesCheck

                Set SQLStmtEpisodesCheck = Server.CreateObject("ADODB.Command")
  	            Set rsEpisodesCheck = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtEpisodesCheck.CommandText = "exec update_client_form_reqs_check " & cur_individual_client_id
  	            SQLStmtEpisodesCheck.CommandType = 1
  	            Set SQLStmtEpisodesCheck.ActiveConnection = conn
  	            SQLStmtEpisodesCheck.CommandTimeout = 45 'Timeout per Command
  	            rsEpisodesCheck.Open SQLStmtEpisodesCheck
                
                wfile.close 
                Set wfile=nothing 
                Set fs=nothing
                
                individuals_form_count = individuals_form_count + 1
                
            Loop
            
           
               
           'if update_prescribers = 1 then
               'Check for new prescribing MDs and add to DB
           '     Set SQLStmtNewMd = Server.CreateObject("ADODB.Command")
           '     Set rsNewMd = Server.CreateObject ("ADODB.Recordset")
           '     SQLStmtNewMd.CommandText = "exec insert_prescriber " & form_GroupID & "," & form_FormID
           '     SQLStmtNewMd.CommandType = 1
                'response.Write "sql = " & SQLStmtNewMd.CommandText
           '     Set SQLStmtNewMd.ActiveConnection = conn
           '     SQLStmtNewMd.CommandTimeout = 45 'Timeout per Command
           '     rsNewMd.Open SQLStmtNewMd   
           'end if          
 	
	        '***CHECK FORM FOR ALERT RULES
            'Set SQLStmtAlertCheck = Server.CreateObject("ADODB.Command")
  	        'Set rsAlertCheck = Server.CreateObject ("ADODB.Recordset")
  	        'SQLStmtAlertCheck.CommandText = "exec alert_rules_check " & form_FormID
  	        'SQLStmtAlertCheck.CommandType = 1
  	        'Set SQLStmtAlertCheck.ActiveConnection = conn
  	        'SQLStmtAlertCheck.CommandTimeout = 45 'Timeout per Command
  	        'rsAlertCheck.Open SQLStmtAlertCheck
            
            '***CHECK FORM FOR MEDICATIONS UPDATES
            'Set SQLStmtMedicationsCheck = Server.CreateObject("ADODB.Command")
  	        'Set rsMedicationsCheck = Server.CreateObject ("ADODB.Recordset")
  	        'SQLStmtMedicationsCheck.CommandText = "exec update_medications_check " & form_FormID
  	        'SQLStmtMedicationsCheck.CommandType = 1
  	        'Set SQLStmtMedicationsCheck.ActiveConnection = conn
  	        'SQLStmtMedicationsCheck.CommandTimeout = 45 'Timeout per Command
  	        'rsMedicationsCheck.Open SQLStmtMedicationsCheck
  	        
  	        '***CHECK FORM FOR ALLERGIES UPDATES
            'Set SQLStmtAllergiesCheck = Server.CreateObject("ADODB.Command")
  	        'Set rsAllergiesCheck = Server.CreateObject ("ADODB.Recordset")
  	        'SQLStmtAllergiesCheck.CommandText = "exec update_allergies_check " & form_FormID
  	        'SQLStmtAllergiesCheck.CommandType = 1
  	        'Set SQLStmtAllergiesCheck.ActiveConnection = conn
  	        'SQLStmtAllergiesCheck.CommandTimeout = 45 'Timeout per Command
  	        'Response.Write "sql = " & SQLStmtAllergiesCheck.CommandText
  	        'rsAllergiesCheck.Open SQLStmtAllergiesCheck
  	        
  	        '***CHECK FORM NEW IAP GOAL/OBJECTIVE SECTIONS	  
            'Set SQLStmtGoalObjCheck = Server.CreateObject("ADODB.Command")
  	        'Set rsGoalObjCheck = Server.CreateObject ("ADODB.Recordset")
  	        'SQLStmtGoalObjCheck.CommandText = "exec update_group_goals_objectives_check " & form_FormID
  	        'SQLStmtGoalObjCheck.CommandType = 1
  	        'Set SQLStmtGoalObjCheck.ActiveConnection = conn
  	        'SQLStmtGoalObjCheck.CommandTimeout = 45 'Timeout per Command
  	        'Response.Write "sql = " & SQLStmtEventsCheck.CommandText
  	        'rsGoalObjCheck.Open SQLStmtGoalObjCheck
  	        
  	        
  	        '***CHECK FORM NEW IAP GOAL/OBJECTIVE SECTIONS	  
            'Set SQLStmtGoalObjCheck = Server.CreateObject("ADODB.Command")
  	        'Set rsGoalObjCheck = Server.CreateObject ("ADODB.Recordset")
  	        'SQLStmtGoalObjCheck.CommandText = "exec generate_individuals_forms " & form_FormID
  	        'SQLStmtGoalObjCheck.CommandType = 1
  	        'Set SQLStmtGoalObjCheck.ActiveConnection = conn
  	        'SQLStmtGoalObjCheck.CommandTimeout = 45 'Timeout per Command
  	        'Response.Write "sql = " & SQLStmtEventsCheck.CommandText
  	        'rsGoalObjCheck.Open SQLStmtGoalObjCheck
            
           'DELETE TEMP FILE
            Set fs=Server.CreateObject("Scripting.FileSystemObject")
            if fs.FileExists(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl") then
                 fs.DeleteFile(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl")
            end if
            set fs=nothing
        
        end if 'END OF CHECK FOR SAME FORM WITH SAME HASH
        
        conn.Close()
        Set conn = Nothing
%>
