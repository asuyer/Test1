<%
    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
    Set rs2 = Server.CreateObject ("ADODB.Recordset")
  	SQLStmt2.CommandText =  "exec get_staff_info_by_username '" & Session("user_name") & "'"
  	SQLStmt2.CommandType = 1
  	Set SQLStmt2.ActiveConnection = conn
  	SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	'response.write "SQL = " & SQLStmt2.CommandText
  	rs2.Open SQLStmt2
  	
  	cur_staff_id = rs2("staff_id")
    cur_staff_name = rs2("First_Name") & " " & rs2("Last_Name")
    local_form_save = rs2("Local_Form_Save")


'**************GET FORM TYPE/DESC FOR THE FORM TYPE PASSED IN OR FOR THE UID IF THAT IS PASSED IN 
Set SQLStmt2 = Server.CreateObject("ADODB.Command")
Set rs2 = Server.CreateObject ("ADODB.Recordset")

if Request.QueryString("ft") <> "" THEN
  	SQLStmt2.CommandText = "select Form_Description from Form_Types_Master where Form_Type = '" & Request.QueryString("ft") & "'"
else
    SQLStmt2.CommandText = "select Form_Description from Form_Types_Master where Form_Type = (Select form_type from location_forms_master where unique_form_id = " & Request.QueryString("uid") & ")"
end if

  	SQLStmt2.CommandType = 1
  	Set SQLStmt2.ActiveConnection = conn
  	SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	'response.write "SQL = " & SQLStmt2.CommandText
  	rs2.Open SQLStmt2
  	
  '	title_form_desc = rs2("Form_Description")

'**************NAME FOR TEMP FILE 
New_name = Session("user_name") & "_" & Year(Date()) & Month(Date()) & Day(Date()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now())

'**************GET EXISTING FORM BY UID IF IT IS PASSED IN
if Request.QueryString("uid") <> "" THEN    
        
    '**************GET FORM INFO FOR THE UID THAT WAS PASSED IN      
    Set SQLStmt2 = Server.CreateObject("ADODB.Command")
    Set rs2 = Server.CreateObject ("ADODB.Recordset")
    SQLStmt2.CommandText = "exec get_location_form_info " & Request.QueryString("uid")
    SQLStmt2.CommandType = 1
    Set SQLStmt2.ActiveConnection = conn
    SQLStmt2.CommandTimeout = 45 'Timeout per Command
    rs2.Open SQLStmt2
  	    
    cur_form_type = rs2("Form_Type")
    cur_form_location = rs2("Location_ID")
    blob_val = rs2("File_Content")
    form_created_date = rs2("Create_Date")
    footer_creat_date = rs2("Footer_Create_Date")
    form_created_by_name = rs2("Create_User_Name")
    cur_form_status = rs2("Status")
    resp_prog_id = rs2("Main_Program")
    resp_prog_name = rs2("Resp_Prog_Name")
    cur_main_staff_id = rs2("Main_Staff")
    resp_staff_name = rs2("Resp_Staff_Name")
    cur_resp_staff_ext_info = rs2("resp_staff_staff_number")'rs2("resp_staff_external_info")
    
      Set SQLStmtPre = Server.CreateObject("ADODB.Command")
        Set rsPre = Server.CreateObject ("ADODB.Recordset")
  	    SQLStmtPre.CommandText = "insert into Forms_Master_Preload(unique_form_id, client_id, form_type, linked_form_id, program_id, staff_id, entry_date) select " & Request.QueryString("uid") & "," & cur_form_client & ",'" & cur_form_type & "'," & cur_form_linked_form_id & "," & cur_main_prog_id & "," & cur_main_staff_id & ", getDate() where " & Request.QueryString("uid") & " not in(select unique_form_id from Forms_Master_Preload where Unique_Form_ID=" & Request.QueryString("uid") & ")"
  	    SQLStmtPre.CommandType = 1
  	    Set SQLStmtPre.ActiveConnection = conn
  	    SQLStmtPre.CommandTimeout = 45 'Timeout per Command



    '******START OF ESP RESP STAFF FILL
    if (cur_form_type = "ESPABU" or cur_form_type = "ESPABUA") and cur_form_status <> "Finalized" THEN
    
        Set SQLStmtRespStaff = Server.CreateObject("ADODB.Command")
        Set rsRespStaff = Server.CreateObject ("ADODB.Recordset")
        SQLStmtRespStaff.CommandText = "update forms_master set main_staff = " & cur_staff_id & " where unique_form_id = " & Request.QueryString("uid")
        SQLStmtRespStaff.CommandType = 1
        Set SQLStmtRespStaff.ActiveConnection = conn
        SQLStmtRespStaff.CommandTimeout = 45 'Timeout per Command
        rsRespStaff.Open SQLStmtRespStaff
        
        cur_resp_staff_ext_info = cur_external_staff_id
    end if
    '******END OF ESP RESP STAFF FILL
    
    new_blob_Contents = Replace(blob_val,"'","''")
  	      	    
  	'**************GET FORM ACCESS FOR THE STAFF USER AND THE UID THAT WAS PASSED IN      	    
    Set SQLStmtA = Server.CreateObject("ADODB.Command")
    Set rsA = Server.CreateObject ("ADODB.Recordset")
    SQLStmtA.CommandText = "exec get_staff_location_form_access " & Request.QueryString("uid") & ",'" & Session("user_name") & "'"
    SQLStmtA.CommandType = 1
    Set SQLStmtA.ActiveConnection = conn
    SQLStmtA.CommandTimeout = 45 'Timeout per Command
    rsA.Open SQLStmtA
  	    
    form_access_level = rsA("Access_Level")
         
  	'**************GET SIGNATURE INFO FOR THE UID THAT WAS PASSED IN   
  	cur_person_signer = ""
  	cur_person_signer_name = ""
  	cur_person_sign_date = ""
  	cur_parent_signer = ""
  	cur_parent_sign_date = ""
  	cur_parent_signer_name = ""
  	cur_md_signer = ""
  	cur_md_signer_name = ""
  	cur_md_sign_date = ""
  	cur_md_sign_position = ""
  	cur_provider_signer = ""
  	cur_provider_signer_name = ""
  	cur_provider_sign_date = ""
  	cur_provider_sign_position = ""
  	cur_supervisor_signer = ""
  	cur_supervisor_signer_name = ""
  	cur_supervisor_sign_date = ""
  	cur_supervisor_sign_position = ""
    
    '**************RECORD THE VIEW ACTION IN TRANSACTION HISTORY        
    Set SQLStmtI = Server.CreateObject("ADODB.Command")
  	Set rsI = Server.CreateObject ("ADODB.Recordset")
  	SQLStmtI.CommandText = "exec insert_transaction_history " & Request.QueryString("uid") & ",'" & Session("user_name") & "','View',''"
  	SQLStmtI.CommandType = 1
  	Set SQLStmtI.ActiveConnection = conn
  	SQLStmtI.CommandTimeout = 45 'Timeout per Command
  	rsI.Open SQLStmtI
    
    'FIND TEMPLATE FILE, MAKE COPY OF IT, AND WRITE THIS CLIENT_ID, CLIENT_NAME, and NEXT UNIQUE_FORM_ID INTO IT
    Set fs2 = CreateObject("Scripting.FileSystemObject")                     
    Set wfile2 = fs2.CreateTextFile(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl", True)
           
    wfile2.Write(blob_val)
            
    Set fs = CreateObject("Scripting.FileSystemObject") 
    fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
  	Set wfile = fs.OpenTextFile(fileToOpen) 
  	        
  	Set fs3 = CreateObject("Scripting.FileSystemObject") 
    Set wfile3 = fs.CreateTextFile(form_root_path & "web_root\temp_forms\" & New_name & "_01.xfdl", True)
            
    '**********USE TO LOOP THROUGH FILE TO FIND THINGS (DESTRUCTIVE TO SIGNATURES WITHOUT CHAR RETURN/LINE FEED)
    foundUFV = 0
    foundProviderSigner = 0
    foundProviderName = 0 
    foundSupervisorSigner = 0
    foundSupervisorName = 0
    foundProviderSignDate = 0
    foundProviderDate = 0
    foundSupervisorDate = 0
    foundSupervisorSignDate = 0
    foundMDSigner = 0
    foundMDName = 0
    foundMDSignDate = 0
    foundMDDate = 0
    foundPersonSigner = 0
    foundPersonName = 0
    foundPersonSignDate = 0
    foundPersonDate = 0
    foundParentSigner = 0
    foundParentName = 0
    foundParentSignDate = 0
    foundParentDate = 0
    foundFooter = 0
    goalSectionsFound = 0
    found_line_before_goals = 0
    found_line_before_needs = 0
    found_here_before = 0
            
    cur_goal_id = -1
    cur_obj_id = -1
   
    do while not wfile.AtEndOfStream 
            
        singleline=wfile.readline 
        'response.Write singleline        
        
        'JSK; Section for creating pre authorized insurance table on PIF
       if InStr(singleline,"<provider_signer></provider_signer>") and foundProvSigner = 0 THEN
           foundProviderSigner = 1
           wfile3.Write(Replace(singleline, "<provider_signer></provider_signer>", "<provider_signer>" & cur_provider_signer & "</provider_signer>") & chr(13) & chr(10))  
            
        elseif InStr(singleline,"<provider_signer_cred></provider_signer_cred>") THEN
           cur_prov_signer = cur_provider_signer
           if cur_provider_sign_position <> "" THEN
               cur_prov_signer = cur_prov_signer & " - " & cur_provider_sign_position
           end if
           wfile3.Write(Replace(singleline, "<provider_signer_cred></provider_signer_cred>", "<provider_signer_cred>" & cur_prov_signer & "</provider_signer_cred>") & chr(13) & chr(10))   
        
        elseif InStr(singleline,"<provider_name></provider_name>") and foundProviderName = 0 THEN
           foundProviderName = 1
           wfile3.Write(Replace(singleline, "<provider_name></provider_name>", "<provider_name>" & cur_provider_signer_name & "</provider_name>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<provider_name_cred></provider_name_cred>") THEN
                               
            cur_prov_name = cur_provider_signer_name
            
            if cur_provider_sign_position <> "" THEN
                cur_prov_name = cur_prov_name & " - " & cur_provider_sign_position
            end if
            
            wfile3.Write(Replace(singleline, "<provider_name_cred></provider_name_cred>", "<provider_name_cred>" & cur_prov_name & "</provider_name_cred>") & chr(13) & chr(10))
        
        elseif InStr(singleline,"<provider_sign_date></provider_sign_date>") and foundProviderSignDate = 0 THEN
           foundProviderSignDate = 1
           wfile3.Write(Replace(singleline, "<provider_sign_date></provider_sign_date>", "<provider_sign_date>" & cur_provider_sign_date & "</provider_sign_date>") & chr(13) & chr(10))
        elseif InStr(singleline,"<provider_date></provider_date>") and foundProviderDate = 0 THEN
           foundProviderDate = 1
           wfile3.Write(Replace(singleline, "<provider_date></provider_date>", "<provider_date>" & cur_provider_sign_date & "</provider_date>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<person_signer></person_signer>") and foundPersonSigner = 0 THEN
           foundPersonSigner = 1
           wfile3.Write(Replace(singleline, "<person_signer></person_signer>", "<person_signer>" & cur_person_signer & "</person_signer>") & chr(13) & chr(10))   
            
        elseif InStr(singleline,"<person_name></person_name>") and foundPersonName = 0 THEN
           foundPersonName = 1
           wfile3.Write(Replace(singleline, "<person_name></person_name>", "<person_name>" & cur_person_signer_name & "</person_name>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<person_sign_date></person_sign_date>") and foundPersonSignDate = 0 THEN
           foundPersonSignDate = 1
           wfile3.Write(Replace(singleline, "<person_sign_date></person_sign_date>", "<person_sign_date>" & cur_person_sign_date & "</person_sign_date>") & chr(13) & chr(10))   
            
        elseif InStr(singleline,"<person_date></person_date>") and foundPersonDate = 0 THEN
           foundPersonDate = 1
           wfile3.Write(Replace(singleline, "<person_date></person_date>", "<person_date>" & cur_person_sign_date & "</person_date>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<md_signer></md_signer>") and foundMDSigner = 0 THEN
           foundMDSigner = 1
           wfile3.Write(Replace(singleline, "<md_signer></md_signer>", "<md_signer>" & cur_md_signer & "</md_signer>") & chr(13) & chr(10))   
            
        elseif InStr(singleline,"<md_signer_cred></md_signer_cred>") THEN
           cur_md_signer = cur_md_signer 
           if cur_md_sign_position <> "" THEN
               cur_md_signer = cur_md_signer & " - " & cur_md_sign_position 
           end if
           wfile3.Write(Replace(singleline, "<md_signer_cred></md_signer_cred>", "<md_signer_cred>" & cur_md_signer & "</md_signer_cred>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<md_name></md_name>") and foundMDName = 0 THEN
           foundMDName = 1
           wfile3.Write(Replace(singleline, "<md_name></md_name>", "<md_name>" & cur_md_signer_name & "</md_name>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<md_name_cred></md_name_cred>") THEN
           cur_md_name = cur_md_signer_name
           if cur_md_sign_position <> "" THEN
               cur_md_name = cur_md_name & " - " & cur_md_sign_position
           end if
           wfile3.Write(Replace(singleline, "<md_name_cred></md_name_cred>", "<md_name_cred>" & cur_md_name & "</md_name_cred>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<md_sign_date></md_sign_date>") and foundMDSignDate = 0 THEN
           foundMDSignDate = 1
           wfile3.Write(Replace(singleline, "<md_sign_date></md_sign_date>", "<md_sign_date>" & cur_md_sign_date & "</md_sign_date>") & chr(13) & chr(10))   
            
        elseif InStr(singleline,"<md_date></md_date>") and foundMDDate = 0 THEN
           foundMDDate = 1
           wfile3.Write(Replace(singleline, "<md_date></md_date>", "<md_date>" & cur_md_sign_date & "</md_date>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<supervisor_signer></supervisor_signer>") and foundSupervisorSigner = 0 THEN
           foundSupervisorSigner = 1
           wfile3.Write(Replace(singleline, "<supervisor_signer></supervisor_signer>", "<supervisor_signer>" & cur_supervisor_signer & "</supervisor_signer>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<supervisor_signer_cred></supervisor_signer_cred>") THEN
            cur_sup_signer = cur_supervisor_signer
            if cur_supervisor_sign_position <> "" THEN
                cur_sup_signer = cur_sup_signer & " - " & cur_supervisor_sign_position 
            end if
            wfile3.Write(Replace(singleline, "<supervisor_signer_cred></supervisor_signer_cred>", "<supervisor_signer_cred>" & cur_sup_signer & "</supervisor_signer_cred>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<supervisor_date></supervisor_date>") and foundSupervisorDate = 0 THEN
            foundSupervisorDate = 1
            wfile3.Write(Replace(singleline, "<supervisor_date></supervisor_date>", "<supervisor_date>" & cur_supervisor_sign_date & "</supervisor_date>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<supervisor_sign_date></supervisor_sign_date>") and foundSupervisorSignDate = 0 THEN
            foundSupervisorSignDate = 1
            wfile3.Write(Replace(singleline, "<supervisor_sign_date></supervisor_sign_date>", "<supervisor_sign_date>" & cur_supervisor_sign_date & "</supervisor_sign_date>") & chr(13) & chr(10))       
            
        elseif InStr(singleline,"<supervisor_name></supervisor_name>") and foundSupervisorName = 0 THEN
            foundSupervisorName = 1
            wfile3.Write(Replace(singleline, "<supervisor_name></supervisor_name>", "<supervisor_name>" & cur_supervisor_signer_name & "</supervisor_name>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<supervisor_name_cred></supervisor_name_cred>") THEN
            cur_sup_name = cur_supervisor_signer_name
            if cur_supervisor_sign_position <> "" THEN
                cur_sup_name = cur_sup_name & " - " & cur_supervisor_sign_position
            end if
            wfile3.Write(Replace(singleline, "<supervisor_name_cred></supervisor_name_cred>", "<supervisor_name_cred>" & cur_sup_name & "</supervisor_name_cred>") & chr(13) & chr(10))                   
        elseif InStr(singleline,"<parent_signer></parent_signer>") and foundParentSigner = 0 THEN
            foundParentSigner = 1
            wfile3.Write(Replace(singleline, "<parent_signer></parent_signer>", "<parent_signer>" & cur_parent_signer & "</parent_signer>") & chr(13) & chr(10))   
            
        elseif InStr(singleline,"<parent_name></parent_name>") and foundParentName = 0 THEN
            foundParentName = 1               
            wfile3.Write(Replace(singleline, "<parent_name></parent_name>", "<parent_name>" & cur_parent_signer_name & "</parent_name>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<parent_sign_date></parent_sign_date>") and foundParentSignDate = 0 THEN
            foundParentSignDate = 1
            wfile3.Write(Replace(singleline, "<parent_sign_date></parent_sign_date>", "<parent_sign_date>" & cur_parent_sign_date & "</parent_sign_date>") & chr(13) & chr(10))   
            
        elseif InStr(singleline,"<parent_date></parent_date>") and foundParentDate = 0 THEN
            foundParentDate = 1
            wfile3.Write(Replace(singleline, "<parent_date></parent_date>", "<parent_date>" & cur_parent_sign_date & "</parent_date>") & chr(13) & chr(10))
              
        elseif InStr(singleline,"<allergies></allergies>") <> 0 and cur_client_allergies <> "" and cur_form_status <> "Finalized" THEN
                wfile3.Write("<allergies>" & cur_client_allergies & "</allergies>" & chr(13) & chr(10))  
               
        elseif InStr(singleline, "<is_esp_pif></is_esp_pif>") <> 0 THEN
                wfile3.Write("<is_esp_pif>" & is_esp_pif & "</is_esp_pif>" & chr(13) & chr(10))
                
        elseif InStr(singleline,"<signer_lock_name>") <> 0 and cur_form_status <> "Finalized" THEN
            wfile3.Write("<signer_lock_name>" & cur_staff_name & "</signer_lock_name>" & chr(13) & chr(10))  
        
        elseif InStr(singleline,"<mandatorycolor>") <> 0 THEN
            wfile3.write(singleline & chr(13) & chr(10) )
            has_mandatory_color_set = 1        
        
        elseif InStr(singleline,"</ufv_settings>") <> 0 and cur_form_status <> "Finalized" and has_mandatory_color_set <> 1 THEN
        'response.Write "redo colors"
                wfile3.Write(Replace(singleline, "</ufv_settings>", "<mandatorycolor>" & mandatory_form_field_color & "</mandatorycolor>" & chr(13) & chr(10) & "<errorcolor>" & error_field_color & "</errorcolor>" & chr(13) & chr(10) & "</ufv_settings>" & chr(13) & chr(10) )) 
        
        elseif InStr(singleline,"<currentprogram>") and cur_form_status <> "Finalized" THEN
               wfile3.Write( "<currentprogram>" & resp_prog_name & "</currentprogram>" & chr(13) & chr(10))           
        
        elseif InStr(singleline,"<organization_name1>") and cur_form_type = "PPPN" and cur_form_status <> "Finalized" THEN            
            wfile3.Write("<organization_name1>"& rsC("DOB") & "</organization_name1>") & chr(13) & chr(10)   
                
        elseif InStr(singleline,"<ssn>") and cur_form_status <> "Finalized" THEN
            client_ssn_tag_start = InStr(singleline,"<ssn>")
            client_ssn_tag_end = InStr(singleline,"</ssn>")
            total_tag_length = client_ssn_tag_end - client_ssn_tag_start
            old_form_ssn = Mid(singleline,(client_ssn_tag_start+5),(total_tag_length-5))
            wfile3.Write(Replace(singleline, "<ssn>"& old_form_ssn & "</ssn>", "<ssn>"& rsC("SSN") & "</ssn>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<gender>") <> 0 and cur_form_status <> "Finalized" THEN
            wfile3.Write("<gender>"& rsC("gender") & "</gender>" & chr(13) & chr(10) )
            
        elseif InStr(singleline,"<admit_date>") and cur_form_status <> "Finalized" THEN
            client_admit_date_tag_start = InStr(singleline,"<admit_date>")
            client_admit_date_tag_end = InStr(singleline,"</admit_date>")
            total_tag_length = client_admit_date_tag_end - client_admit_date_tag_start
            old_form_admit_date = Mid(singleline,(client_admit_date_tag_start+10),(total_tag_length-10))
            wfile3.Write(Replace(singleline, "<admit_date>"& old_form_admit_date & "</admit_date>", "<admit_date>"& rsC("Registration_Date") & "</admit_date>") & chr(13) & chr(10))
            
        elseif InStr(singleline,"<modifiable>on</modifiable>") and foundUFV = 0  and form_access_level = "V" THEN
            foundUFV = 1
            wfile3.Write(Replace(singleline, "<modifiable>on</modifiable>", "<modifiable>off</modifiable>") & chr(13) & chr(10)) 
            
        elseif InStr(singleline,"<modifiable>off</modifiable>") and foundUFV = 0  and form_access_level = "E" THEN
            foundUFV = 1
            wfile3.Write(Replace(singleline, "<modifiable>off</modifiable>", "<modifiable>on</modifiable>") & chr(13) & chr(10)) 
        
        elseif (InStr(singleline,"<xforms:instance id=""Axis_1_List"" xmlns="""">") <> 0 or InStr(singleline,"<xforms:instance xmlns="""" id=""Axis_1_List"">") <> 0) and cur_form_status <> "Finalized" THEN
                wfile3.Write (singleline & chr(13) & chr(10))
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10))

                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "get_diag_codes 1"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	            
  	            Do Until rsV.EOF
        
                     wfile3.Write ("<choice value=""" & rsV("diag_id") & """>" & rsv("description") & "~" & rsV("diag_value") & "</choice>" & chr(13) & chr(10))
                     '<choice value="319">ABC-XYZ</choice>
                                
                rsV.MoveNext
                Loop
        
  	   elseif InStr(singleline,"<popup") <> 0 and cur_form_status <> "Finalized" and cur_form_type <> "PIA" THEN 'THIS IS FOR GENERATED PIA forms
                cur_sid = ""
                
                'POSSIBLY DYNAMIC FIND SID AND LOOKUP IN CODEMAP
                popup_tag_start = InStr(singleline,"<popup sid=""")
                popup_tag_end = InStr(singleline,""">")
                total_tag_length = popup_tag_end - popup_tag_start
                cur_sid = Mid(singleline,(popup_tag_start+12),(total_tag_length-12))
                
                if InStrRev(cur_sid, "_") THEN
                    undloc = InStrRev(cur_sid,"_")
                    numcheck = Mid(cur_sid,undloc+1)
                    if isNumeric(numcheck) Then
                        stripped_sid = Mid(cur_sid, 1, undloc-1)
                       ' response.Write"found underscore, base value = " & stripped_sid
                        cur_sid = stripped_sid
                    end if
                end if     
                   
  	            if cur_sid = "Loc_Code" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select location_id, description from location_master order by description" 
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("location_id")
  	                    cur_desc = rsV("description")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_desc & " (" & cur_name & ")</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                else
                    wfile3.Write (singleline & chr(13) & chr(10))
                end if
                 	        
  	      elseif foundDynamicPopup = 1 and InStr(singleline,"<xforms:label></xforms:label>") <> 0 and cur_form_status <> "Finalized" THEN                           
  	                wfile3.Write (singleline & chr(13) & chr(10))
  	                
  	                'WRITE OUT NEW VALUES
  	                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_value_list_for_sid '" & cur_sid & "'" 
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
  	                
  	                Do Until rsV.EOF
  	                    
  	                   'special cases of values to skip here
  	                    if rsV("short_desc") = "Ed which is ESP Provider or Sub" and cur_sid = "Eval_Location" and Request.QueryString("ft") = "ESPACA" THEN
  	                        'skip it
  	                    else
  	                        wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                            if responsible_prog_in_cbfs = 1 and rsV("short_desc") = "Residential DMH" and cur_sid = "Referral_Source" THEN
                            wfile3.Write ("<xforms:label>DMH</xforms:label>" & chr(13) & chr(10))
                            else
                            wfile3.Write ("<xforms:label>" & rsV("short_desc") & "</xforms:label>" & chr(13) & chr(10))
                            end if
                            wfile3.Write ("<xforms:value>" & rsV("code_name") & "</xforms:value>" & chr(13) & chr(10))
                            wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                            wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                            wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                            wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                            wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                            wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                            wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                            wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                            wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                            wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                            wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                            wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                        end if
                        
                    rsV.MoveNext
                    Loop
                    
                    'SKIP THROUGH OLD VALUES
                    Do Until InStr(singleline,"</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                                        
                    foundDynamicPopup = 0
        
        '******************TEMP PRODCEDURE CODE SECTION*************************************
        elseif InStr(singleline,"<!-- NO SHOW RELATED RULES BEGIN -->") <> 0 and cur_form_status <> "Finalized" THEN
            wfile3.Write (singleline & chr(13) & chr(10))
            singleline=wfile.readline 
            
            'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
            Do Until InStr(singleline, "<!-- NO SHOW RELATED RULES END -->") <> 0
                if InStr(singleline,"id=""no_show_procedure_rules""") <> 0 THEN
                    singleline = wfile.readline
                else                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline = wfile.readline
                end if
            Loop
            
            wfile3.Write (singleline & chr(13) & chr(10))
        '***********************************************************************************
        
        
        elseif InStr(singleline,"</menu>") THEN
                   'if InStr(singleline, "</save>") THEN
                   'else
                   'end if
                   
                   if local_form_save = "0" THEN
                        wfile3.Write( "<save>hidden</save></menu>" & chr(13) & chr(10))            
                   else
                        wfile3.Write( "<save>on</save></menu>" & chr(13) & chr(10)) 
                   end if
        
        
        '################################################ 	                
  	 elseif InStr(singleline,"<dr_first_meds_plain_text>") <> 0 and cur_form_status <> "Finalized" THEN
                    
                    'FIND END OF OLD FIELD AND STRIP IT OUT
                    if InStr(singleline,"</dr_first_meds_plain_text>") = 0 THEN
                         Do Until found_end_of_dr_first_plain_text = 1
                            if InStr(singleline,"</dr_first_meds_plain_text>") <> 0 THEN
                                found_end_of_dr_first_plain_text = 1
                            else                                                 
                                singleline = wfile.readline
                            end if
                        Loop
                    end if
                    
                    med_string_start = "<dr_first_meds_plain_text>"
                    med_string_mid = ""
                    med_string_end = "</dr_first_meds_plain_text>" & chr(13) & chr(10)                    
                                        
                    efs_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & cur_form_client & ",'All'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                
  	                    if rsMeds("stop_date") = "" and rsMeds("med_type") = "Dr_First" THEN
                           'clean_stop_date = REPLACE(rsMeds("stop_date"),"-","/")
                            clean_start_date = REPLACE(rsMeds("start_date"),"-","/")
                            clean_fill_date = REPLACE(rsMeds("fill_date"),"-","/")
      	                    
      	                        efs_meds_count = efs_meds_count + 1
      	                
  	                            med_string_mid = med_string_mid & REPLACE(Replace(efs_meds_count & ": " & rsMeds("brand_name") & " (" & rsMeds("generic_name") & "), Strength:" & rsMeds("strength") & " " & rsMeds("strength_unit") & ", Quantity:" & rsMeds("quantity") & " " & rsMeds("quantity_unit") & ", Dose:" & rsMeds("dose") & " " & rsMeds("dose_unit") & ", Frequency:" & rsMeds("dose_timing") & ", Additional Freq:" & rsMeds("dose_other") & ", Route:" & rsMeds("action") & " " & rsMeds("route") & ", Instructions:" & rsMeds("patient_notes") & ", Prescribed By:" & rsMeds("prov_first_name") & " " & rsMeds("prov_last_name") & "  NID: " & rsMeds("prov_npi") & ", Refills:" & rsMeds("refills") & ", Fill Date:" & clean_fill_date & ", Start Date: " & clean_start_date & ", Stop Reason:" & rsMeds("stop_reason") & ", Comments:" & rsMeds("comments") & chr(13) & chr(10) & chr(13) & chr(10),"&","&amp;"),"<","&lt;") 
      	                                    
                        end if   
                        
                                      
  	                rsMeds.MoveNext
  	                Loop
  	                            
  	                wfile3.Write (med_string_start & med_string_mid & med_string_end)
  	                
 '################################################  	                
  	 elseif InStr(singleline,"<other_meds_plain_text>") <> 0 and cur_form_status <> "Finalized" THEN
                    
                    'FIND END OF OLD FIELD AND STRIP IT OUT
                    if InStr(singleline,"</other_meds_plain_text>") = 0 THEN
                         Do Until found_end_of_other_meds_plain_text = 1
                            if InStr(singleline,"</other_meds_plain_text>") <> 0 THEN
                                found_end_of_other_meds_plain_text = 1
                            else                                                 
                                singleline = wfile.readline
                            end if
                        Loop
                    end if
                    
                    med_string_start = "<other_meds_plain_text>"
                    med_string_mid = ""
                    med_string_end = "</other_meds_plain_text>" & chr(13) & chr(10)                    
                                        
                    efs_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & cur_form_client & ",'All'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                
  	                    if (rsMeds("stop_date") = "" and rsMeds("med_type") = "Other_Orders") or rsMeds("med_type") = "Current" THEN
                           'clean_stop_date = REPLACE(rsMeds("stop_date"),"-","/")
                            clean_start_date = REPLACE(rsMeds("start_date"),"-","/")
      	                    
      	                    efs_meds_count = efs_meds_count + 1
      	                    if rsMeds("brand_name") = "Other" THEN
      	                        cur_brand_gen_string = rsMeds("brand_name") & " (" & rsMeds("brand_name_other")
      	                    else
      	                       cur_brand_gen_string = rsMeds("brand_name") & " (" & rsMeds("generic_name")
      	                    end if 
      	                    
  	                        med_string_mid = med_string_mid & REPLACE(Replace(efs_meds_count & ": " & cur_brand_gen_string & "), Strength:" & rsMeds("strength") & " " & rsMeds("strength_unit") & ", Quantity:" & rsMeds("quantity") & " " & rsMeds("quantity_unit") & ", Dose:" & rsMeds("dose") & " " & rsMeds("dose_unit") & ", Frequency:" & rsMeds("dose_timing") & ", Route:" & rsMeds("action") & " " & rsMeds("route") & ", Instructions:" & rsMeds("patient_notes") & ", Prescribed By:" & rsMeds("prov_first_name") & " " & rsMeds("prov_last_name") & "  NID: " & rsMeds("prov_npi") & ", Refills:" & rsMeds("refills") & ", Type:" & Replace(rsMeds("med_type"),"_"," ") & ", Start Date: " & clean_start_date & chr(13) & chr(10) & chr(13) & chr(10),"&","&amp;"),"<","&lt;")   
      	                                    
                        end if               
  	                rsMeds.MoveNext
  	                Loop
  	                            
  	                wfile3.Write (med_string_start & med_string_mid & med_string_end)  	   
        elseif InStr(singleline,"<xforms:instance id=""popupList"" xmlns="""">") <> 0 and Request.QueryString("ft") = "CHTF" THEN
            found_no_show_instance = 1 
                    
        else
            wfile3.write(singleline & chr(13) & chr(10))
            
        end if
            
        loop 
        
        wfile.close 
        Set wfile=nothing 
        Set fs=nothing  
        Set wfile2=nothing
        Set fs2=nothing
        Set fs3=nothing
        Set wfile3=nothing
            
        Set fs = CreateObject("Scripting.FileSystemObject") 
  	    fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & "_01.xfdl"
  	    Set wfile = fs.OpenTextFile(fileToOpen) 
  	        
        blob_val = wfile.ReadAll
            
        Set wfile=nothing 
        Set fs=nothing 
%>
<%               
    else '*************************LOOK FOR A TEMPLATE OF A SPECIFIC FORM TYPE
  	    
  	    if Request.QueryString("ft") <> "" THEN  	    
  	        Set SQLStmt2 = Server.CreateObject("ADODB.Command")
  	        Set rs2 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt2.CommandText = "select (select max(Local_Form_Save) from Roles_Master where Role_ID IN(select role_id from staff_roles_assign where staff_id = (select staff_id from staff_master where user_name = '" & Session("user_name") & "'))) as Local_Form_Save, Template_File_Content, (select cast(datepart(day,getdate()) as varchar(2)) + ' ' + substring(convERT(varchar(50), getdate(), 113),3,10) + cast( case when datepart(hour,getdate()) > 12 then datepart(hour,getdate()) - 12 else datepart(hour,getdate()) end as varchar(2)) + ':' + cast(datepart(minute,getdate()) as varchar(2)) + case when datepart(hour,getdate()) > 12 then ' PM' else ' AM' end) as form_create_date from Form_Types_Master where Form_Type = '" & Request.QueryString("ft") & "'"
  	        SQLStmt2.CommandType = 1
  	        Set SQLStmt2.ActiveConnection = conn
  	        'response.Write "sql = " & SQLStmt2.CommandText
  	        SQLStmt2.CommandTimeout = 45 'Timeout per Command
  	        rs2.Open SQLStmt2
  	        
  	        local_form_save = rs2("Local_Form_Save")
  	        blob_val = rs2("Template_File_Content")
  	        form_create_date = rs2("form_create_date")
  	                	           	        
  	        'FIND NEXT UNIQUE FORM ID
  	        Set SQLStmt3 = Server.CreateObject("ADODB.Command")
  	        Set rs3 = Server.CreateObject ("ADODB.Recordset")

  	        SQLStmt3.CommandText = "exec get_next_uid"
  	        SQLStmt3.CommandType = 1
  	        Set SQLStmt3.ActiveConnection = conn
  	        SQLStmt3.CommandTimeout = 45 'Timeout per Command
  	        rs3.Open SQLStmt3
  	        
  	        next_uid = rs3("next_uid")
  	         	        
  	        if next_uid <> "" THEN
                'CREATE RECORD IN PRELOAD TABLE FOR THE NEW ADD   
                if Request.QueryString("lfid") = "" AND Request.QueryString("iapid") = "" THEN
                    pre_lfid = -1
                elseif Request.QueryString("lfid") <> "" THEN
                    pre_lfid = Request.QueryString("lfid")
                else
                    pre_lfid = Request.QueryString("iapid")    
                end if
                
                pre_gid = Request.QueryString("gid")
                
                pre_pid = -1
                pre_sid = -1
                pre_svcid = -1
                pre_cid = -1              
                           
                Set SQLStmtPre = Server.CreateObject("ADODB.Command")
                Set rsPre = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtPre.CommandText = "insert into Forms_Master_Preload(unique_form_id, client_id, form_type, linked_form_id, program_id, staff_id, entry_date) select " & next_uid & "," & pre_cid & ",'" & Request.QueryString("ft") & "'," & pre_lfid & "," & pre_pid & "," & pre_sid & ", getDate()" 
  	            SQLStmtPre.CommandType = 1
  	            Set SQLStmtPre.ActiveConnection = conn
  	            SQLStmtPre.CommandTimeout = 45 'Timeout per Command
 	            rsPre.Open SQLStmtPre

            end if
  	          	        
  	        'FIND TEMPLATE FILE, MAKE COPY OF IT, AND WRITE THIS CLIENT_ID, CLIENT_NAME, and NEXT UNIQUE_FORM_ID INTO IT
  	        Set fs2 = CreateObject("Scripting.FileSystemObject") 
            Set wfile2 = fs2.CreateTextFile(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl", True)
            
            wfile2.Write(blob_val)
            
            Set fs = CreateObject("Scripting.FileSystemObject") 
  	        fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & ".xfdl"
  	        Set wfile = fs.OpenTextFile(fileToOpen) 
  	        
  	        Set fs3 = CreateObject("Scripting.FileSystemObject") 
            Set wfile3 = fs.CreateTextFile(form_root_path & "web_root\temp_forms\" & New_name & "_01.xfdl", True)
            
            '**********USE TO LOOP THROUGH FILE TO FIND THINGS (DESTRUCTIVE TO SIGNATURES WITHOUT CHAR RETURN/LINE FEED)
            foundClientName = 0
            foundClientFName = 0
            foundClientMName = 0
            foundClientLName = 0
            foundClientID = 0
            foundFormID = 0
            foundLinkedFormID = 0
            foundProvSigner = 0
            foundRecNum = 0
            foundClientDOB = 0
            foundClientSSN = 0
            foundClientProgram = 0
            
            counter=0
            model_end_found = 0
            
            'GET STAFF INFO
            Set SQLStmtGroup = Server.CreateObject("ADODB.Command")
            Set rsGroup = Server.CreateObject ("ADODB.Recordset")
  	        SQLStmtGroup.CommandText = "get_location_info " & Request.QueryString("lid")
  	        SQLStmtGroup.CommandType = 1
  	        Set SQLStmtGroup.ActiveConnection = conn
  	        SQLStmtGroup.CommandTimeout = 45 'Timeout per Command
 	        rsGroup.Open SQLStmtGroup
 	        
 	        cur_location_name = rsGroup("Location_Name")
 	        cur_location_address = rsGroup("address")
 	        cur_location_city = rsGroup("city")
 	        cur_location_state = rsGroup("state")
 	        cur_location_zip = rsGroup("zip")
 	        cur_location_client_count = rsGroup("location_client_count")
 	        cur_location_staff_count = rsGroup("location_staff_count")
           
            '*************************MODEL SECTION*********************************   
            do while not wfile.AtEndOfStream and model_end_found = 0
            counter=counter+1
            singleline=wfile.readline 

            if InStr(singleline,"</xformsmodels>") <> 0 then
             wfile3.Write(singleline & chr(13) & chr(10)) 
             model_end_found = 1
             
            '******************************
                                 
            'THESE COME FROM THE CLIENT RECORD LEVEL   
            elseif InStr(singleline,"<client_name></client_name>") and foundClientName = 0 THEN
               foundClientName = 1
               wfile3.Write(Replace(singleline, "<client_name></client_name>", "<client_name>" & rs2("client_name") & "</client_name>") & chr(13) & chr(10)) 
            
            elseif InStr(singleline,"<location_name></location_name>") and foundClientName = 0 THEN
               wfile3.Write("<location_name>" & cur_location_name & "</location_name>" & chr(13) & chr(10))                               
            
             elseif InStr(singleline,"<location_address></location_address>") and foundClientName = 0 THEN
               wfile3.Write("<location_address>" & cur_location_address & "</location_address>" & chr(13) & chr(10))                  
            
            elseif InStr(singleline,"<location_city></location_city>") and foundClientName = 0 THEN
               wfile3.Write("<location_city>" & cur_location_city & "</location_city>" & chr(13) & chr(10))    
            elseif InStr(singleline,"<location_state></location_state>") and foundClientName = 0 THEN
               wfile3.Write("<location_state>" & cur_location_state & "</location_state>" & chr(13) & chr(10))    
            elseif InStr(singleline,"<location_zip></location_zip>") and foundClientName = 0 THEN
               wfile3.Write("<location_zip>" & cur_location_zip & "</location_zip>" & chr(13) & chr(10))  
            elseif InStr(singleline,"<clients_enrolled></clients_enrolled>") and foundClientName = 0 THEN
               wfile3.Write("<clients_enrolled>" & cur_location_client_count & "</clients_enrolled>" & chr(13) & chr(10)) 
            elseif InStr(singleline,"<staff_assigned></staff_assigned>") and foundClientName = 0 THEN
               wfile3.Write("<staff_assigned>" & cur_location_staff_count & "</staff_assigned>" & chr(13) & chr(10))        
               
                 
            elseif InStr(singleline,"<contact_phone></contact_phone>") THEN
                  wfile3.Write("<contact_phone>" & PIF_contact_phone & "</contact_phone>" & chr(13) & chr(10))        
                                   
            elseif InStr(singleline,"<contact_phone_ext></contact_phone_ext>") THEN
                  wfile3.Write("<contact_phone_ext>" & PIF_contact_phone_ext & "</contact_phone_ext>" & chr(13) & chr(10))      
            elseif InStr(singleline,"<legal_guardian></legal_guardian>") THEN
                  wfile3.Write("<legal_guardian>" & PIF_legal_guardian & "</legal_guardian>" & chr(13) & chr(10))        
            elseif InStr(singleline,"<guardian_phone></guardian_phone>") THEN
                  wfile3.Write("<guardian_phone>" & PIF_guardian_phone & "</guardian_phone>" & chr(13) & chr(10))
            elseif InStr(singleline,"<guardian_phone_ext></guardian_phone_ext>") THEN
                  wfile3.Write("<guardian_phone_ext>" & PIF_guardian_phone_ext & "</guardian_phone_ext>" & chr(13) & chr(10))      
            elseif InStr(singleline,"<emergency_contact></emergency_contact>") THEN
                  wfile3.Write("<emergency_contact>" & PIF_emergency_contact & "</emergency_contact>" & chr(13) & chr(10))        
            elseif InStr(singleline,"<emergency_contact_phone></emergency_contact_phone>") THEN
                  wfile3.Write("<emergency_contact_phone>" & PIF_emergency_contact_phone & "</emergency_contact_phone>" & chr(13) & chr(10))
            elseif InStr(singleline,"<emergency_contact_phone_ext></emergency_contact_phone_ext>") THEN
                  wfile3.Write("<emergency_contact_phone_ext>" & PIF_emergency_contact_phone_ext & "</emergency_contact_phone_ext>" & chr(13) & chr(10))      
           
            elseif InStr(singleline, "<is_esp_pif></is_esp_pif>") <> 0 THEN
                wfile3.Write("<is_esp_pif>" & is_esp_pif & "</is_esp_pif>" & chr(13) & chr(10))
                  
          elseif InStr(singleline,"<form_id></form_id>") and foundFormID = 0 THEN
               foundFormID = 1
               wfile3.Write(Replace(singleline, "<form_id></form_id>", "<form_id>" & next_uid & "</form_id>") & chr(13) & chr(10)) 
               
           elseif InStr(singleline,"<form_location_id></form_location_id>") THEN
               wfile3.Write("<form_location_id>" & Request.QueryString("lid") & "</form_location_id>" & chr(13) & chr(10)) 
                           
            elseif InStr(singleline,"<form_div></form_div>") <> 0 THEN
                 
                 Set SQLStmtDiv = Server.CreateObject("ADODB.Command")
                 Set rsDiv = Server.CreateObject ("ADODB.Recordset")                 
                 SQLStmtDiv.CommandText = "select division from program_master where program_id = " & Request.QueryString("pid")
  	             SQLStmtDiv.CommandType = 1
  	             Set SQLStmtDiv.ActiveConnection = conn
  	             'response.write "SQL = " & SQLStmtDiv.CommandText
  	             rsDiv.Open SQLStmtDiv
  	             
  	             form_div = rsDiv("division")  
  	             
  	             wfile3.Write("<form_div>" & form_div & "</form_div>" & chr(13) & chr(10)) 
  	                     
            elseif InStr(singleline,"<iap_dated></iap_dated>") THEN
               wfile3.Write(Replace(singleline, "<iap_dated></iap_dated>", "<iap_dated>" & iap_dated & "</iap_dated>") & chr(13) & chr(10))
            
            elseif InStr(singleline,"<encounter_id></encounter_id>") THEN
               wfile3.Write(Replace(singleline, "<encounter_id></encounter_id>", "<encounter_id>" & next_eid & "</encounter_id>") & chr(13) & chr(10)) 
                          
               wfile3.Write(Replace(singleline, "<rec_num></rec_num>", "<rec_num>" & form_rec_num_id & "</rec_num>") & chr(13) & chr(10)) 
            
        
             elseif InStr(singleline,"<allergies></allergies>") <> 0 and cur_client_allergies <> "" THEN 'Request.QueryString("ft") = "MEDREV" THEN
                wfile3.Write("<allergies>" & cur_client_allergies & "</allergies>" & chr(13) & chr(10))           

            elseif InStr(singleline,"<linked_form_id></linked_form_id>") and foundLinkedFormID = 0 and Request.QueryString("lfid") <> "" THEN
               foundLinkedFormID = 1
               wfile3.Write(Replace(singleline, "<linked_form_id></linked_form_id>", "<linked_form_id>" & Request.QueryString("lfid") & "</linked_form_id>") & chr(13) & chr(10))      
          
            elseif InStr(singleline,"<organization_name></organization_name>") THEN
               wfile3.Write(Replace(singleline, "<organization_name></organization_name>", "<organization_name>" & org_name & "</organization_name>") & chr(13) & chr(10))
            
            elseif InStr(singleline,"<gender></gender>") <> 0 then

               wfile3.Write(Replace(singleline, "<gender></gender>", "<gender>" & rs2("gender") & "</gender>") & chr(13) & chr(10))
            
            elseif InStr(singleline,"<admit_date></admit_date>") <> 0 then
                wfile3.Write(Replace(singleline, "<admit_date></admit_date>", "<admit_date>" & rs2("admit_date") & "</admit_date>") & chr(13) & chr(10))
         
            elseif InStr(singleline,"<dr_first_allergies>") <> 0 THEN 'ALL OTHER ALLERGIES HERE
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    'wfile3.Write (singleline & chr(13) & chr(10))
                    
                    medrev_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_allergies_for_client_by_type " & Request.QueryString("cid") & ",'Dr_First'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                  	                  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                    cur_allergy_kind = rsMeds("rcopia_drug_id")
  	                    'cur_allergy_status = rsMeds("active")
  	                  	
  	                  	if cur_allergy_kind = "" THEN
  	                  	    cur_allergy_kind = "Other"
  	                  	else
  	                  	    cur_allergy_kind = "Medication"
  	                  	end if                     
  	                  	                     
  	                    if rsMeds("active") = "Active" or rsMeds("active") = "1" THEN
  	                        cur_active = "Active"
  	                
  	                        wfile3.Write ("<row>" & chr(13) & chr(10))
  	                            wfile3.Write ("<allergy_type>Dr_First</allergy_type>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_name>" & REPLACE(rsMeds("name"),"&","&amp;") & "</allergy_name>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_reaction>" & REPLACE(rsMeds("reaction"),"&","&amp;") & "</allergy_reaction>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_kind>" & cur_allergy_kind & "</allergy_kind>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_onset_date>" & rsMeds("onsetdate") & "</allergy_onset_date>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_status>" & cur_active & "</allergy_status>" & chr(13) & chr(10))
     
                            wfile3.Write ("</row>" & chr(13) & chr(10))
                                                 
                            medrev_meds_count = medrev_meds_count + 1   
                        end if
                        
  	                rsMeds.MoveNext
  	                Loop
  	                
  	                if medrev_meds_count = 0 THEN
  	                   ' wfile3.Write ("<row>" & chr(13) & chr(10))
  	                end if
  	                
  	                Do Until InStr(singleline, "</dr_first_allergies>") <> 0
  	                    singleline=wfile.readline 
  	                Loop    
  	                
  	                'write out end of table2
  	                wfile3.Write (singleline & chr(13) & chr(10))
            
            elseif InStr(singleline,"<emr_allergies>") <> 0 THEN 'ALL OTHER ALLERGIES HERE
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    'wfile3.Write (singleline & chr(13) & chr(10))
                    
                    medrev_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_allergies_for_client_by_type " & Request.QueryString("cid") & ",'EMR'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                    if rsMeds("active") = "Active" or rsMeds("active") = "1" THEN
  	                        cur_active = "Active"
  	                     	                
  	                        wfile3.Write ("<row>" & chr(13) & chr(10))
  	                            wfile3.Write ("<client_allergy_id>" & rsMeds("client_allergy_id") & "</client_allergy_id>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_type1>EMR</allergy_type1>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_name1>" & REPLACE(rsMeds("name"),"&","&amp;") & "</allergy_name1>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_reaction1>" & REPLACE(rsMeds("reaction"),"&","&amp;") & "</allergy_reaction1>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_kind1>" & rsMeds("allergy_kind") & "</allergy_kind1>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_onset_date1>" & rsMeds("onsetdate") & "</allergy_onset_date1>" & chr(13) & chr(10))
                                wfile3.Write ("<allergy_status1>" & cur_active & "</allergy_status1>" & chr(13) & chr(10))
     
                            wfile3.Write ("</row>" & chr(13) & chr(10))
                        
                             medrev_meds_count = medrev_meds_count + 1   
                        end if                     
                        
  	                rsMeds.MoveNext
  	                Loop
  	                
  	                if medrev_meds_count = 0 THEN
  	                    wfile3.Write ("<row>" & chr(13) & chr(10))
  	                         wfile3.Write ("<client_allergy_id></client_allergy_id>" & chr(13) & chr(10))
                            wfile3.Write ("<allergy_type1>EMR</allergy_type1>" & chr(13) & chr(10))
                            wfile3.Write ("<allergy_name1></allergy_name1>" & chr(13) & chr(10))
                            wfile3.Write ("<allergy_reaction1></allergy_reaction1>" & chr(13) & chr(10))
                            wfile3.Write ("<allergy_kind1></allergy_kind1>" & chr(13) & chr(10))
                            wfile3.Write ("<allergy_onset_date1></allergy_onset_date1>" & chr(13) & chr(10))
                            wfile3.Write ("<allergy_status1></allergy_status1>" & chr(13) & chr(10))
                        wfile3.Write ("</row>" & chr(13) & chr(10))
  	                end if
  	                
  	                Do Until InStr(singleline, "</emr_allergies>") <> 0
  	                    singleline=wfile.readline 
  	                Loop    
  	                
  	                'write out end of table2
  	                wfile3.Write (singleline & chr(13) & chr(10))
            
            elseif InStr(singleline,"<dr_first_meds_table>") <> 0 THEN 'DR FIRST MEDS HERE
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    'wfile3.Write (singleline & chr(13) & chr(10))
                    
                    medrev_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'Dr_First'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtRules.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                    wfile3.Write ("<row>" & chr(13) & chr(10))  	                    
  	                        wfile3.Write ("<medication_type>Dr_First</medication_type>" & chr(13) & chr(10))
  	                        wfile3.Write ("<start_date>" & rsMeds("start_date") & "</start_date>" & chr(13) & chr(10))
  	                        wfile3.Write ("<dc_date>" & rsMeds("stop_date") & "</dc_date>" & chr(13) & chr(10))
  	                        wfile3.Write ("<med_brand>" & REPLACE(rsMeds("brand_name"),"&","&amp;") & "</med_brand>" & chr(13) & chr(10))
  	                        wfile3.Write ("<med_generic>" & REPLACE(rsMeds("generic_name"),"&","&amp;") & "</med_generic>" & chr(13) & chr(10))
  	                        wfile3.Write ("<strength>" & rsMeds("strength") & "</strength>" & chr(13) & chr(10))
  	                        wfile3.Write ("<amount>" & rsMeds("quantity") & " " & rsMeds("quantity_unit") & "</amount>" & chr(13) & chr(10))
  	                        wfile3.Write ("<dose>" & REPLACE(rsMeds("dose"),"&","&amp;") & " " & rsMeds("dose_unit") & "</dose>" & chr(13) & chr(10)) 
  	                        wfile3.Write ("<additional>" & REPLACE(rsMeds("dose_other"),"&","&amp;") & "</additional>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medroutecode>" & rsMeds("action" ) & " " & rsMeds("route") & "</medroutecode>" & chr(13) & chr(10))
  	                        wfile3.Write ("<prescribe_md>" & rsMeds("prov_first_name") & " " & rsMeds("prov_last_name") & "  NID: " & rsMeds("prov_npi") &"</prescribe_md>" & chr(13) & chr(10))
  	                        wfile3.Write ("<instructions>" & REPLACE(rsMeds("comments"),"&","&amp;") & "</instructions>" & chr(13) & chr(10))  	                    
  	                        wfile3.Write ("<reason>" & REPLACE(rsMeds("stop_reason"),"&","&amp;") & "</reason>" & chr(13) & chr(10))
  	                        wfile3.Write ("<hmea_res></hmea_res>" & chr(13) & chr(10))
  	                        wfile3.Write ("<hmea_day></hmea_day>" & chr(13) & chr(10))
  	                        wfile3.Write ("<psych_treatment></psych_treatment>" & chr(13) & chr(10))
  	                        wfile3.Write ("<last_presc_id>" & rsMeds("last_prescription_id") & "</last_presc_id>" & chr(13) & chr(10))
  	                        wfile3.Write ("<rcopia_id>" & rsMeds("rcopia_drug_id") & "</rcopia_id>" & chr(13) & chr(10))
  	                        wfile3.Write ("<rcid>" & rsMeds("rcid") & "</rcid>" & chr(13) & chr(10))
  	                        wfile3.Write ("<pharm_instructions>" & REPLACE(rsMeds("other_notes"),"&","&amp;") & "</pharm_instructions>" & chr(13) & chr(10))
  	                        wfile3.Write ("<patient_instructions>" & REPLACE(rsMeds("patient_notes"),"&","&amp;") & "</patient_instructions>" & chr(13) & chr(10))
  	                        wfile3.Write ("<duration>" & rsMeds("duration") & "</duration>" & chr(13) & chr(10))
  	                        wfile3.Write ("<refills>" & rsMeds("refills") & "</refills>" & chr(13) & chr(10))
  	                        wfile3.Write ("<substitutions_permitted>" & rsMeds("subsitution_permitted") & "</substitutions_permitted>" & chr(13) & chr(10))
  	                        wfile3.Write ("<fill_date>" & rsMeds("fill_date") & "</fill_date>" & chr(13) & chr(10))
  	                        wfile3.Write ("<schedule>" & rsMeds("schedule") & "</schedule>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medfrqcode>" & rsMeds("dose_timing") & "</medfrqcode>" & chr(13) & chr(10))
                        wfile3.Write ("</row>" & chr(13) & chr(10))
                                                 
                        medrev_meds_count = medrev_meds_count + 1   
                        
  	                rsMeds.MoveNext
  	                Loop  	           
  	                
  	                Do Until InStr(singleline, "</dr_first_meds_table>") <> 0
  	                    singleline=wfile.readline 
  	                Loop    
  	                
  	                'write out end of table2
  	                wfile3.Write (singleline & chr(13) & chr(10))
            
            elseif InStr(singleline,"<emr_meds_table>") <> 0 THEN 'ALL OTHER MEDS HERE
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    'wfile3.Write (singleline & chr(13) & chr(10))
                    
                    medrev_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'All_Other'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                    wfile3.Write ("<row>" & chr(13) & chr(10))
  	                        wfile3.Write ("<client_med_id>" & rsMeds("client_med_id") & "</client_med_id>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medication_type1>" & rsMeds("med_type") & "</medication_type1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<start_date1>" & rsMeds("start_date") & "</start_date1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<dc_date1>" & rsMeds("stop_date") & "</dc_date1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<med_brand1>" & REPLACE(rsMeds("brand_name"),"&","&amp;") & "</med_brand1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<med_brand_other1>" & REPLACE(rsMeds("brand_name_other"),"&","&amp;") & "</med_brand_other1>" & chr(13) & chr(10))  	                    
  	                        wfile3.Write ("<med_generic1>" & REPLACE(rsMeds("generic_name"),"&","&amp;") & "</med_generic1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<strength1>" & rsMeds("strength") & "</strength1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medunitcode_4>" & rsMeds("strength_unit") & "</medunitcode_4>" & chr(13) & chr(10))
  	                        wfile3.Write ("<amount1>" & rsMeds("quantity") & "</amount1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medunitcode_5>" & rsMeds("quantity_unit") & "</medunitcode_5>" & chr(13) & chr(10))
  	                        wfile3.Write ("<dose1>" & REPLACE(rsMeds("dose"),"&","&amp;") & "</dose1>" & chr(13) & chr(10)) 
  	                        wfile3.Write ("<medunitcode_6>" & rsMeds("dose_unit") & "</medunitcode_6>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medfrqcode_2>" & rsMeds("dose_timing") & "</medfrqcode_2>" & chr(13) & chr(10))
  	                        wfile3.Write ("<frequency_input1>" & rsMeds("dose_other") & "</frequency_input1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medroutecode_2>" & rsMeds("route") & "</medroutecode_2>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_prescribing_md_new></cme_prescribing_md_new>" & chr(13) & chr(10))
  	                        wfile3.Write ("<additional1>" & REPLACE(rsMeds("dose_other"),"&","&amp;") & "</additional1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<instructions1>" & REPLACE(rsMeds("other_notes"),"&","&amp;") & "</instructions1>" & chr(13) & chr(10))  	        
  	                        wfile3.Write ("<hmea_res>" & rsMeds("is_res") & "</hmea_res>" & chr(13) & chr(10))
  	                        wfile3.Write ("<hmea_day>" & rsMeds("is_day") & "</hmea_day>" & chr(13) & chr(10)) 
  	                        wfile3.Write ("<cme_substitutions_permitted>" & rsMeds("subsitution_permitted") & "</cme_substitutions_permitted>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_duration>" & rsMeds("duration") & "</cme_duration>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_instructions>" & rsMeds("comments") & "</cme_instructions>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_patient_instructions>" & rsMeds("patient_notes") & "</cme_patient_instructions>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_reason>" & REPLACE(rsMeds("stop_reason"),"&","&amp;") & "</cme_reason>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_prescribing_md>" & rsMeds("prov_first_name") & "</cme_prescribing_md>" & chr(13) & chr(10))
                            wfile3.Write ("<refills1>" & rsMeds("refills") & "</refills1>" & chr(13) & chr(10))
                            wfile3.Write ("<cme_new_ic>" & rsMeds("new_informed_consent") & "</cme_new_ic>" & chr(13) & chr(10))
                            wfile3.Write ("<cme_type>" & rsMeds("new_or_renew") & "</cme_type>" & chr(13) & chr(10))
  	                        wfile3.Write ("<psych_treatment1>" & rsMeds("is_psych") & "</psych_treatment1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<rationale_condition>" & rsMeds("rationale_condition") & "</rationale_condition>" & chr(13) & chr(10))
  	                        wfile3.Write ("<side_effects>" & rsMeds("side_effects") & "</side_effects>" & chr(13) & chr(10))
     
                        wfile3.Write ("</row>" & chr(13) & chr(10))
                                                 
                        medrev_meds_count = medrev_meds_count + 1   
                        
  	                rsMeds.MoveNext
  	                Loop
  	                
  	                if medrev_meds_count = 0 THEN
  	                    wfile3.Write ("<row>" & chr(13) & chr(10))
  	                        wfile3.Write ("<client_med_id></client_med_id>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medication_type1></medication_type1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<start_date1></start_date1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<dc_date1></dc_date1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<med_brand1></med_brand1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<med_brand_other1></med_brand_other1>" & chr(13) & chr(10))  	                    
  	                        wfile3.Write ("<med_generic1></med_generic1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<strength1></strength1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medunitcode_4></medunitcode_4>" & chr(13) & chr(10))
  	                        wfile3.Write ("<amount1></amount1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medunitcode_5></medunitcode_5>" & chr(13) & chr(10))
  	                        wfile3.Write ("<dose1></dose1>" & chr(13) & chr(10)) 
  	                        wfile3.Write ("<medunitcode_6></medunitcode_6>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medfrqcode_2></medfrqcode_2>" & chr(13) & chr(10))
  	                        wfile3.Write ("<frequency_input1></frequency_input1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<medroutecode_2></medroutecode_2>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_prescribing_md_new></cme_prescribing_md_new>" & chr(13) & chr(10))
  	                        wfile3.Write ("<additional1></additional1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<instructions1></instructions1>" & chr(13) & chr(10))  	        
  	                        wfile3.Write ("<hmea_res></hmea_res>" & chr(13) & chr(10))
  	                        wfile3.Write ("<hmea_day></hmea_day>" & chr(13) & chr(10)) 
  	                        wfile3.Write ("<cme_substitutions_permitted></cme_substitutions_permitted>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_duration></cme_duration>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_instructions></cme_instructions>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_patient_instructions></cme_patient_instructions>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_reason></cme_reason>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_prescribing_md></cme_prescribing_md>" & chr(13) & chr(10))
                            wfile3.Write ("<refills1></refills1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<cme_rc></cme_rc>" & chr(13) & chr(10))
                            wfile3.Write ("<cme_new></cme_new>" & chr(13) & chr(10))
                            wfile3.Write ("<cme_new_ic></cme_new_ic>" & chr(13) & chr(10))
                            wfile3.Write ("<cme_type></cme_type>" & chr(13) & chr(10))
  	                        wfile3.Write ("<psych_treatment1></psych_treatment1>" & chr(13) & chr(10))
  	                        wfile3.Write ("<rationale_condition></rationale_condition>" & chr(13) & chr(10))
  	                        wfile3.Write ("<side_effects></side_effects>" & chr(13) & chr(10))
                        wfile3.Write ("</row>" & chr(13) & chr(10))
  	                end if
  	                
  	                Do Until InStr(singleline, "</emr_meds_table>") <> 0
  	                    singleline=wfile.readline 
  	                Loop    
  	                
  	                'write out end of table2
  	                wfile3.Write (singleline & chr(13) & chr(10))
  	        
  	        '******************** CURRENT PAST MEDS TABLE *************************************
            elseif InStr(singleline,"<current_past_meds_table>") <> 0 THEN 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    'wfile3.Write (singleline & chr(13) & chr(10))
                    
                    medrev_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
                        SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'CP'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                    wfile3.Write ("<row>" & chr(13) & chr(10))
  	                    
  	                        wfile3.Write ("<client_med_id>" & rsMeds("client_med_id") & "</client_med_id>" & chr(13) & chr(10))
  	                        wfile3.Write ("<current_past_med>" & rsMeds("med_type") & "</current_past_med>" & chr(13) & chr(10))
                            wfile3.Write ("<med_brand>" & REPLACE(rsMeds("brand_name"),"&","&amp;") & "</med_brand>" & chr(13) & chr(10))
                            wfile3.Write ("<med_brand_other>" & REPLACE(rsMeds("brand_name_other"),"&","&amp;") & "</med_brand_other>" & chr(13) & chr(10))
                            wfile3.Write ("<med_generic>" & REPLACE(rsMeds("generic_name"),"&","&amp;") & "</med_generic>" & chr(13) & chr(10))
                            wfile3.Write ("<dose>" & REPLACE(rsMeds("dose"),"&","&amp;") & "</dose>" & chr(13) & chr(10))
                            wfile3.Write ("<medunitcode>" & rsMeds("dose_unit") & "</medunitcode>" & chr(13) & chr(10))
                            wfile3.Write ("<medroutecode>" & rsMeds("route") & "</medroutecode>" & chr(13) & chr(10))
                            wfile3.Write ("<medfrqcode>" & rsMeds("dose_timing") & "</medfrqcode>" & chr(13) & chr(10))
                            wfile3.Write ("<prescribing_md>" & rsMeds("prov_first_name") & "</prescribing_md>" & chr(13) & chr(10))
                            wfile3.Write ("<prescribing_md_new></prescribing_md_new>" & chr(13) & chr(10))
                            wfile3.Write ("<rationale_condition>" & rsMeds("rationale_condition") & "</rationale_condition>" & chr(13) & chr(10))
                            wfile3.Write ("<side_effects>" & rsMeds("side_effects") & "</side_effects>" & chr(13) & chr(10)) 	 
                            
                            wfile3.Write ("<completed></completed>" & chr(13) & chr(10)) 	                 
                
                        wfile3.Write ("</row>" & chr(13) & chr(10))
                                                 
                        medrev_meds_count = medrev_meds_count + 1   
                        
  	                rsMeds.MoveNext
  	                Loop
  	                
  	                if medrev_meds_count = 0 THEN
  	                    wfile3.Write ("<row>" & chr(13) & chr(10))
  	                    
  	                        wfile3.Write ("<client_med_id></client_med_id>" & chr(13) & chr(10))
  	                        wfile3.Write ("<current_past_med></current_past_med>" & chr(13) & chr(10))
                            wfile3.Write ("<med_brand></med_brand>" & chr(13) & chr(10))
                            wfile3.Write ("<med_brand_other></med_brand_other>" & chr(13) & chr(10))
                            wfile3.Write ("<med_generic></med_generic>" & chr(13) & chr(10))
                            wfile3.Write ("<dose></dose>" & chr(13) & chr(10))
                            wfile3.Write ("<medunitcode></medunitcode>" & chr(13) & chr(10))
                            wfile3.Write ("<medroutecode></medroutecode>" & chr(13) & chr(10))
                            wfile3.Write ("<medfrqcode></medfrqcode>" & chr(13) & chr(10))                            
                            wfile3.Write ("<prescribing_md></prescribing_md>" & chr(13) & chr(10))
                            wfile3.Write ("<prescribing_md_new></prescribing_md_new>" & chr(13) & chr(10))
                            wfile3.Write ("<rationale_condition></rationale_condition>" & chr(13) & chr(10))
                            wfile3.Write ("<side_effects></side_effects>" & chr(13) & chr(10)) 	  
                            
                            wfile3.Write ("<completed></completed>" & chr(13) & chr(10))
                                 
                        wfile3.Write ("</row>" & chr(13) & chr(10))
  	                end if
  	                
  	                Do Until InStr(singleline, "</current_past_meds_table>") <> 0
  	                    singleline=wfile.readline 
  	                Loop    
  	                
  	                'write out end of table2
  	                wfile3.Write (singleline & chr(13) & chr(10))
  	        
 	                
  	        elseif InStr(singleline,"<other_meds_table>") <> 0 THEN 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    'wfile3.Write (singleline & chr(13) & chr(10))
                    
                    medrev_meds_count = 0
                    
                    Set SQLStmtMeds2 = Server.CreateObject("ADODB.Command")
                    Set rsMeds2 = Server.CreateObject ("ADODB.Recordset")

  	                SQLStmtMeds2.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'All' "
  	                SQLStmtMeds2.CommandType = 1
  	                Set SQLStmtMeds2.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds2.CommandText
  	                SQLStmtMeds2.CommandTimeout = 45 'Timeout per Command
  	                rsMeds2.Open SQLStmtMeds2

  	                'write out each row
  	                Do Until rsMeds2.EOF
  	                
  	                    new_or_renew = rsMeds2("new_or_renew")
                        renew_check = ""
                        new_check = ""
                        dc_check = ""
                        med_status = ""
                        
                        IF rsMeds2("stop_date") <> "" THEN
                            dc_check = "true"
                            Med_Status = "Discontinued"
                        ELSEIF new_or_renew = "New" THEN
                            new_check = "true"
                            med_status = "New_Adjusted"
                        ELSEIF new_or_renew = "Renew" THEN
                            renew_check = "true"
                            med_status = "Refill"
                        END IF
                        
                        if rsMeds2("stop_date") <> "" THEN
                            clean_stop_date = REPLACE(rsMeds2("stop_date"),"-","/")
                        else
                            clean_stop_date = "Now"
                        end if
                        
                        clean_start_date = REPLACE(rsMeds2("start_date"),"-","/")
  	                
  	                    if rsMeds2("med_type") = "Dr_First" or rsMeds2("med_type") = "Other_Orders" THEN
  	                        clean_date_range = clean_start_date & "-" & clean_stop_date
  	                    else
  	                        clean_date_range = ""
  	                    end if
  	                    
  	                     wfile3.Write ("<row>" & chr(13) & chr(10))
  	                    
  	                            wfile3.Write ("<client_med_id>" & rsMeds2("client_med_id") & "</client_med_id>" & chr(13) & chr(10))
                                wfile3.Write ("<med_brand>" & REPLACE(rsMeds2("brand_name"),"&","&amp;") & " (" & REPLACE(rsMeds2("generic_name"),"&","&amp;") & ")" & chr(13) & chr(10) & "*" & rsMeds2("med_type") & "* " & clean_date_range & "</med_brand>" & chr(13) & chr(10))
                                wfile3.Write ("<med_status>" & med_status & "</med_status>" & chr(13) & chr(10))
                                wfile3.Write ("<rationale_condition>" & REPLACE(rsMeds2("comments"),"&","&amp;") & "</rationale_condition>" & chr(13) & chr(10))
                                wfile3.Write ("<dose_route_freq>" & REPLACE(rsMeds2("dose"),"&","&amp;") & " " & rsMeds2("dose_unit") & " / " & rsMeds2("route") & " / " & rsMeds2("dose_timing") & "</dose_route_freq>" & chr(13) & chr(10))
                                wfile3.Write ("<amount_refills>" & rsMeds2("quantity") & " " & rsMeds2("quantity_unit") & " / " & rsMeds2("refills") & "</amount_refills>" & chr(13) & chr(10))
                                wfile3.Write ("<med_plan></med_plan>" & chr(13) & chr(10))
                                wfile3.Write ("<prescribing_md>" & rsMeds2("prov_first_name") & "</prescribing_md>" & chr(13) & chr(10))
                                wfile3.Write ("<med_renew>" & renew_check & "</med_renew>" & chr(13) & chr(10))
                                wfile3.Write ("<med_new>" & new_check & "</med_new>" & chr(13) & chr(10))
                                wfile3.Write ("<med_dc>" & dc_check & "</med_dc>" & chr(13) & chr(10))
                                wfile3.Write ("<med_dose>" & REPLACE(rsMeds2("dose"),"&","&amp;") & " " & rsMeds2("dose_unit") & "</med_dose>" & chr(13) & chr(10))
                                wfile3.Write ("<med_route>" & rsMeds2("route") & "</med_route>" & chr(13) & chr(10))
                                wfile3.Write ("<med_freq>" & rsMeds2("dose_timing") & "</med_freq>" & chr(13) & chr(10))
                                wfile3.Write ("<med_days>" & rsMeds2("duration") & "</med_days>" & chr(13) & chr(10))
                                wfile3.Write ("<med_qty>" & rsMeds2("quantity") & "</med_qty>" & chr(13) & chr(10))   
                                wfile3.Write ("<med_refills>" & rsMeds2("refills") & "</med_refills>" & chr(13) & chr(10))   
                                wfile3.Write ("<med_consent>" & rsMeds2("new_informed_consent") & "</med_consent>" & chr(13) & chr(10)) 
  	                        
  	                        if Request.QueryString("ft") = "ESPDPSC" THEN
  	                            wfile3.Write ("<med_instructions>" & REPLACE(rsMeds2("patient_notes"),"&","&amp;") & "</med_instructions>" & chr(13) & chr(10))  	        
  	                            wfile3.Write ("<completed></completed>" & chr(13) & chr(10))  	        
  	                        end if 
  	                        
  	                        if Request.QueryString("ft") = "HCPMO" THEN
  	                            wfile3.Write ("<medication_ordered>" & REPLACE(rsMeds2("brand_name"),"&","&amp;") & " (" & REPLACE(rsMeds2("generic_name"),"&","&amp;") & ")" & chr(13) & chr(10) & "*" & rsMeds2("med_type") & "*" & "</medication_ordered>" & chr(13) & chr(10))  	        
  	                            wfile3.Write ("<medication_ordered_dose>" & REPLACE(rsMeds2("dose"),"&","&amp;") & " " & rsMeds2("dose_unit") & "</medication_ordered_dose>" & chr(13) & chr(10))  	        
  	                            wfile3.Write ("<medication_ordered_freq>" & rsMeds2("dose_timing") & "</medication_ordered_freq>" & chr(13) & chr(10))  	        
  	                            wfile3.Write ("<medication_ordered_hcp>" & rsMeds2("prov_first_name") & "</medication_ordered_hcp>" & chr(13) & chr(10))  	        
	        
  	                        end if 
  	                        
  	                        if Request.QueryString("ft") = "ICP" THEN
  	                            wfile3.Write ("<prescriber>" & rsMeds2("prov_first_name") & "</prescriber>" & chr(13) & chr(10))
  	                            wfile3.Write ("<med>" & REPLACE(rsMeds2("brand_name"),"&","&amp;") & " (" & REPLACE(rsMeds2("generic_name"),"&","&amp;") & ")" & chr(13) & chr(10) & "*" & rsMeds2("med_type") & "*" & "</med>" & chr(13) & chr(10))
  	                            wfile3.Write ("<symptoms></symptoms>" & chr(13) & chr(10))                                  
  	                        end if
  	                        'end if  	                    
  	                                   
                        wfile3.Write ("</row>" & chr(13) & chr(10))
                                                 
                        medrev_meds_count = medrev_meds_count + 1   
                        
  	                rsMeds2.MoveNext
  	                Loop	               
  	                
  	                Do Until InStr(singleline, "</other_meds_table>") <> 0
  	                'response.Write "single line = " & singleline & "<br />"
  	                    singleline=wfile.readline 
  	                Loop    
  	                
  	                'write out end of other_meds_table
  	                wfile3.Write (singleline & chr(13) & chr(10))
            
            elseif InStr(singleline,"<plain_text_meds_block></plain_text_meds_block>") <> 0 THEN
                    
                    med_string_start = "<plain_text_meds_block>"
                    med_string_mid = ""
                    med_string_end = "</plain_text_meds_block>" & chr(13) & chr(10)                    
                                        
                    efs_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'All'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                
  	                    if rsMeds("stop_date") <> "" THEN
                            clean_stop_date = REPLACE(rsMeds("stop_date"),"-","/")
                        else
                            clean_stop_date = "Now"
                        end if
                        
                        clean_start_date = REPLACE(rsMeds("start_date"),"-","/")
  	                
  	                    if rsMeds("med_type") = "Dr_First" or rsMeds("med_type") = "Other_Orders" THEN
  	                        clean_date_range = clean_start_date & " - " & clean_stop_date
  	                    else
  	                        clean_date_range = ""
  	                    end if
  	                
  	                    if rsMeds("med_type") = "Dr_First" or rsMeds("med_type") = "Other_Orders" THEN
  	                    
      	                    efs_meds_count = efs_meds_count + 1
  	                
  	                        med_string_mid = med_string_mid & REPLACE(Replace(efs_meds_count & ": " & rsMeds("brand_name") & " (" & rsMeds("generic_name") & "), Strength:" & rsMeds("strength") & " " & rsMeds("strength_unit") & ", Quantity:" & rsMeds("quantity") & " " & rsMeds("quantity_unit") & ", Dose:" & rsMeds("dose") & " " & rsMeds("dose_unit") & ", Frequency:" & rsMeds("dose_timing") & ", Route:" & rsMeds("action") & " " & rsMeds("route") & ", Instructions:" & rsMeds("patient_notes") & ", Prescribed By:" & rsMeds("prov_first_name") & " " & rsMeds("prov_last_name") & "  NID: " & rsMeds("prov_npi") & chr(13) & chr(10) & "*" & rsMeds("med_type") & "*, " & clean_date_range & chr(13) & chr(10) & chr(13) & chr(10),"&","&amp;"),"<","&lt;")   
  	                                    
                        end if                
  	                rsMeds.MoveNext
  	                Loop
  	                            
  	                wfile3.Write (med_string_start & med_string_mid & med_string_end)  
  	                             
            elseif InStr(singleline,"<dr_first_meds_plain_text></dr_first_meds_plain_text>") <> 0 THEN
                    
                    med_string_start = "<dr_first_meds_plain_text>"
                    med_string_mid = ""
                    med_string_end = "</dr_first_meds_plain_text>" & chr(13) & chr(10)                    
                                        
                    efs_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'All'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                
  	                    if rsMeds("stop_date") = "" and rsMeds("med_type") = "Dr_First" THEN
                           'clean_stop_date = REPLACE(rsMeds("stop_date"),"-","/")
                            clean_start_date = REPLACE(rsMeds("start_date"),"-","/")
                            clean_fill_date = REPLACE(rsMeds("fill_date"),"-","/")
      	                    
      	                        efs_meds_count = efs_meds_count + 1
      	                
  	                            med_string_mid = med_string_mid & REPLACE(Replace(efs_meds_count & ": " & rsMeds("brand_name") & " (" & rsMeds("generic_name") & "), Strength:" & rsMeds("strength") & " " & rsMeds("strength_unit") & ", Quantity:" & rsMeds("quantity") & " " & rsMeds("quantity_unit") & ", Dose:" & rsMeds("dose") & " " & rsMeds("dose_unit") & ", Frequency:" & rsMeds("dose_timing") & ", Additional Freq:" & rsMeds("dose_other") & ", Route:" & rsMeds("action") & " " & rsMeds("route") & ", Instructions:" & rsMeds("patient_notes") & ", Prescribed By:" & rsMeds("prov_first_name") & " " & rsMeds("prov_last_name") & "  NID: " & rsMeds("prov_npi") & ", Refills:" & rsMeds("refills") & ", Fill Date:" & clean_fill_date & ", Start Date: " & clean_start_date & ", Stop Reason:" & rsMeds("stop_reason") & ", Comments:" & rsMeds("comments") & chr(13) & chr(10) & chr(13) & chr(10),"&","&amp;"),"<","&lt;")   
      	                                    
                        end if               
  	                rsMeds.MoveNext
  	                Loop
  	                            
  	                wfile3.Write (med_string_start & med_string_mid & med_string_end)
  	                
     	                
  	    elseif InStr(singleline,"<other_meds_plain_text></other_meds_plain_text>") <> 0 THEN
                    
                    med_string_start = "<other_meds_plain_text>"
                    med_string_mid = ""
                    med_string_end = "</other_meds_plain_text>" & chr(13) & chr(10)                    
                                        
                    efs_meds_count = 0
                    
                    Set SQLStmtMeds = Server.CreateObject("ADODB.Command")
                    Set rsMeds = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtMeds.CommandText = "exec get_current_meds_for_client_by_type " & Request.QueryString("cid") & ",'All'"
  	                SQLStmtMeds.CommandType = 1
  	                Set SQLStmtMeds.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtMeds.CommandText
  	                SQLStmtMeds.CommandTimeout = 45 'Timeout per Command
  	                rsMeds.Open SQLStmtMeds
  	                
  	                'write out each row
  	                Do Until rsMeds.EOF
  	                
  	                    if (rsMeds("stop_date") = "" and rsMeds("med_type") = "Other_Orders") or rsMeds("med_type") = "Current" THEN
                            clean_start_date = REPLACE(rsMeds("start_date"),"-","/")
      	                    
      	                    efs_meds_count = efs_meds_count + 1
      	                    if rsMeds("brand_name") = "Other" THEN
      	                        cur_brand_gen_string = rsMeds("brand_name") & " (" & rsMeds("brand_name_other")
      	                    else
      	                       cur_brand_gen_string = rsMeds("brand_name") & " (" & rsMeds("generic_name")
      	                    end if 

                            med_string_mid = med_string_mid & REPLACE(Replace(efs_meds_count & ": " & cur_brand_gen_string & "), Strength:" & rsMeds("strength") & " " & rsMeds("strength_unit") & ", Quantity:" & rsMeds("quantity") & " " & rsMeds("quantity_unit") & ", Dose:" & rsMeds("dose") & " " & rsMeds("dose_unit") & ", Frequency:" & rsMeds("dose_timing") & ", Route:" & rsMeds("action") & " " & rsMeds("route") & ", Instructions:" & rsMeds("patient_notes") & ", Prescribed By:" & rsMeds("prov_first_name") & " " & rsMeds("prov_last_name") & "  NID: " & rsMeds("prov_npi") & ", Refills:" & rsMeds("refills") & ", Type:" & Replace(rsMeds("med_type"),"_"," ") & ", Start Date: " & clean_start_date & chr(13) & chr(10) & chr(13) & chr(10),"&","&amp;"),"<","&lt;")   
      	                                    
                        end if               
  	                rsMeds.MoveNext
  	                Loop
  	                            
  	                wfile3.Write (med_string_start & med_string_mid & med_string_end)  	                
  	                          
            'END OF MEDS SECTIONS
              
            elseif InStr(singleline,"<currentprogram>") THEN
               wfile3.Write( "<currentprogram>" & cur_resp_program_name & "</currentprogram>" & chr(13) & chr(10))    
                 
        

             elseif InStr(singleline,"<completed_by></completed_by>") <> 0  THEN 
                wfile3.Write("<completed_by>" & cur_staff_name & "</completed_by>" & chr(13) & chr(10))   

            
            elseif InStr(singleline,"<provider_number></provider_number>") <> 0 and svc_staff <> "" THEN 
                wfile3.Write("<provider_number>" & svc_staff & "</provider_number>" & chr(13) & chr(10))    
            
            elseif InStr(singleline,"<provider_number></provider_number>") <> 0 and svc_staff = "" THEN
                wfile3.Write("<provider_number>" & cur_resp_staff_ext_info & "</provider_number>" & chr(13) & chr(10))
            
            elseif InStr(singleline,"<procedure_code></procedure_code>") <> 0 and svc_code <> "" THEN
                wfile3.Write("<procedure_code>" & svc_code & "</procedure_code>" & chr(13) & chr(10))    
            
            elseif InStr(singleline,"<date_of_service></date_of_service>") <> 0 and svc_start_date_time_clean <> "" THEN
                                
                wfile3.Write("<date_of_service>" & LEFT(svc_start_date_time_clean,10) & "</date_of_service>" & chr(13) & chr(10))      
            
            
            
            elseif InStr(singleline,"<start_time></start_time>") <> 0 and svc_start_date_time <> "" THEN
                wfile3.Write("<start_time>" & LEFT(RIGHT(svc_start_date_time,8),5) & ":00</start_time>" & chr(13) & chr(10)) 
                
            elseif InStr(singleline,"<loc_code></loc_code>") and svc_location <> "" THEN
               wfile3.Write("<loc_code>" & svc_location & "</loc_code>" & chr(13) & chr(10))     
            
           elseif InStr(singleline, "<diag_code></diag_code>") <> 0 and parent_ca_diag <> "" THEN 
                wfile3.Write "<diag_code>" & parent_ca_diag & "</diag_code>" & chr(13) & chr(10)
             
           elseif InStr(singleline,"<signer_lock_name>") <> 0 THEN
                
                wfile3.Write("<signer_lock_name>" & cur_staff_name & "</signer_lock_name>" & chr(13) & chr(10))
                                                                
             elseif InStr(singleline, "<!--INSERT MED BRAND GENERIC RULES HERE-->") THEN
                Set SQLStmtRules = Server.CreateObject("ADODB.Command")
                Set rsRules = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtRules.CommandText = "select medication_id, medication_brand_name, medication_generic_name from medications_master where is_active = 1"
  	            SQLStmtRules.CommandType = 1
  	            Set SQLStmtRules.ActiveConnection = conn
  	            'response.Write "sql = " & SQLStmtRules.CommandText
  	            SQLStmtRules.CommandTimeout = 45 'Timeout per Command
  	            rsRules.Open SQLStmtRules
  	            
  	            'med_rule_counter = 1
  	            
                 'response.Write "in rule check and rule type = " & rsRules("Rule_Type") & " and statement = x" & rsRules("Full_Rule_Statement") & "x"
                 page = "page2"
                 med_table_name = "current_past_meds_table"
                 med_generic_field_name = "med_generic"
                 med_brand_field_name = "med_brand"
                 
                 if Request.QueryString("ft") = "MEDREV" then
                     page = "page1"
                     med_table_name = "emr_meds_table"
                     med_generic_field_name = "med_generic1"
                     med_brand_field_name = "med_brand1"
                 end if
                 
                 if Request.QueryString("ft") = "ACA" then
                     page = "page4"
                 elseif Request.QueryString("ft") = "ESPACA" then
                    page = "page4"
                    med_table_name = "blank_current_past_meds_table"
                 elseif Request.QueryString("ft") = "ESPCCA" then
                    page = "page4"
                    med_table_name = "blank_current_past_meds_table"
                 elseif Request.QueryString("ft") = "CCA" then
                     page = "page7"
                 elseif Request.QueryString("ft") = "CTPV" then
                    page = "page1"
                 elseif Request.QueryString("ft") = "ICP" then
                    page = "page1"
                    med_table_name = "blank_current_past_meds_table"
                 elseif Request.QueryString("ft") = "PCSU" then
                    page = "page1"
                 elseif Request.QueryString("ft") = "PHARMPN" then
                    page = "page2"
                 elseif Request.QueryString("ft") = "TDSP" then
                    page = "page2"
                 elseif Request.QueryString("ft") = "MEDADD" then
                    page = "page1"
                    med_table_name = "current_past_meds_table"
                 end if
                 
                    wfile3.Write ("<xforms:bind id=""Med_Generic"" nodeset=""instance('Generated')/" & page & "/" & med_table_name & "/row/" & med_generic_field_name & """ calculate=""if(../" & med_brand_field_name & " = 'Other', " & med_generic_field_name & " , ")
              
                Do Until rsRules.EOF
                    wfile3.Write ("if(../" & med_brand_field_name & " = '" & rsRules("medication_id") & "','" & rsRules("medication_generic_name") & "',")      
                    med_rule_counter = med_rule_counter + 1           
                rsRules.MoveNext
                Loop
                
                wfile3.Write ("''")
                
                Do Until med_rule_counter = 0
                    wfile3.Write (")")                   
                    med_rule_counter = med_rule_counter - 1
                Loop
                     
                wfile3.Write (")"" readonly=""boolean-from-string(if(../" & med_brand_field_name & " = 'Other', 'false','true'))""></xforms:bind>" & chr(13) & chr(10))
                

            '******************TEMP PRODCEDURE CODE SECTION*************************************
            elseif InStr(singleline,"<!-- NO SHOW RELATED RULES BEGIN -->") <> 0 THEN
                wfile3.Write (singleline & chr(13) & chr(10))
                singleline=wfile.readline 
                
                'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                Do Until InStr(singleline, "<!-- NO SHOW RELATED RULES END -->") <> 0
                    if InStr(singleline,"id=""no_show_procedure_rules""") <> 0 THEN
                        singleline = wfile.readline
                    else                    
                        wfile3.Write (singleline & chr(13) & chr(10))
                        singleline = wfile.readline
                    end if
                Loop
                
                wfile3.Write (singleline & chr(13) & chr(10))
        '***********************************************************************************


            elseif InStr(singleline, "<!-- INSERT DYNAMIC FORM RULES HERE -->") THEN
            
                Set SQLStmtRules = Server.CreateObject("ADODB.Command")
                Set rsRules = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtRules.CommandText = "get_dynamic_rules_for_form '" & Request.QueryString("ft") & "'"
  	            SQLStmtRules.CommandType = 1
  	            Set SQLStmtRules.ActiveConnection = conn
  	            'response.Write "sql = " & SQLStmtRules.CommandText
  	            SQLStmtRules.CommandTimeout = 45 'Timeout per Command
  	            rsRules.Open SQLStmtRules
                
                dynamic_rule_counter = 1
                
                Do Until rsRules.EOF
                 'response.Write "in rule check and rule type = " & rsRules("Rule_Type") & " and statement = x" & rsRules("Full_Rule_Statement") & "x"
                  
                        if rsRules("Rule_Type") = "relevant" and rsRules("Full_Rule_Statement") = "" THEN
               
                            wfile3.Write ("<xforms:bind id=""" & Request.QueryString("ft") & "_" & dynamic_rule_counter & """ nodeset=""instance('Generated')/page" & rsRules("Form_Page") & "/" & rsRules("Target_SID") & """ relevant=""boolean-from-string(if(../" & rsRules("Driver_SID") & " = '" & rsRules("Driver_Value") & "','true','false'))""></xforms:bind>" & chr(13) & chr(10))
                        elseif rsRules("Rule_Type") = "calculate" and rsRules("Full_Rule_Statement") = "" THEN
                        
                            wfile3.Write ("<xforms:bind id=""" & Request.QueryString("ft") & "_" & dynamic_rule_counter & """ nodeset=""instance('Generated')/page" & rsRules("Form_Page") & "/" & rsRules("Target_SID") & """ calculate=""if(../" & rsRules("Driver_SID") & " = '" & rsRules("Driver_Value") & "','" & rsRules("Result_Value") & "', ../" & rsRules("Target_SID") & ")"" readonly=""boolean-from-string(if(../" & rsRules("Driver_SID") & " = '" & rsRules("Driver_Value") & "','true', 'false'))""></xforms:bind>" & chr(13) & chr(10))

                        else
                        
                            'wfile3.Write( rsRules("Full_Rule_Statement") & 
                        
                        end if
                
                        dynamic_rule_counter = dynamic_rule_counter + 1
                        
                rsRules.MoveNext
                Loop 
            
            elseif InStr(singleline,"<xforms:instance id=""popupList"" xmlns="""">") <> 0 and Request.QueryString("ft") <> "CHTF" THEN
                wfile3.Write (singleline & chr(13) & chr(10)) 'instane line
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10)) 'data line
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10)) 'location line

                'WRITE OUT ALL PERSON PRESENT ONES FIRST
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "select service_code_id, service_code + '-' + REPLACE(REPLACE(service_desc,'&','&amp;'),'<','&lt;') as service_desc from billing_procedure_service_codes where [end] IS NULL and service_code NOT IN('160','161','162','163','164','165') and service_type_id = 1 order by service_desc"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV      	            
      	        
      	        wfile3.Write ("<choice value=""" & "Person_Present" & """ label=""" & "Person Present" & """>" & chr(13) & chr(10))
      	            
  	            Do Until rsV.EOF
                     'Write out program level choice
                     wfile3.Write ("<choice value=""" & rsV("service_code_id") & """ label=""" & rsV("service_desc") & """></choice>" & chr(13) & chr(10))                   
                                                   
                rsV.MoveNext
                Loop
                
                wfile3.Write ("</choice>" & chr(13) & chr(10))
                
                'WRITE OUT ALL PERSON PRESENT ONES FIRST
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "select service_code_id, service_code + '-' + REPLACE(REPLACE(service_desc,'&','&amp;'),'<','&lt;') as service_desc from billing_procedure_service_codes where [end] IS NULL and service_code NOT IN('160','161','162','163','164','165') and service_type_id = 1 order by service_desc"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV      	            
      	        
      	        wfile3.Write ("<choice value=""" & "Family_Present" & """ label=""" & "Family Present" & """>" & chr(13) & chr(10))
      	            
  	            Do Until rsV.EOF
                     'Write out program level choice
                     wfile3.Write ("<choice value=""" & rsV("service_code_id") & """ label=""" & rsV("service_desc") & """></choice>" & chr(13) & chr(10))                   
                                                   
                rsV.MoveNext
                Loop
                
                wfile3.Write ("</choice>" & chr(13) & chr(10))
                
                'WRITE OUT PERSON NO SHOW ITEMS NEXT
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "select service_code_id, service_code + '-' + REPLACE(REPLACE(service_desc,'&','&amp;'),'<','&lt;') as service_desc from billing_procedure_service_codes where [end] IS NULL and service_code IN('160','161','162','163','164','165') and service_type_id = 1 order by service_desc"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV      	            
      	        
      	        wfile3.Write ("<choice value=""" & "Person_No_Show" & """ label=""" & "Person No Show" & """>" & chr(13) & chr(10))
      	            
  	            Do Until rsV.EOF
                     'Write out program level choice
                     wfile3.Write ("<choice value=""" & rsV("service_code_id") & """ label=""" & rsV("service_desc") & """></choice>" & chr(13) & chr(10))                             
                                                   
                rsV.MoveNext
                Loop
                
                wfile3.Write ("</choice>" & chr(13) & chr(10))
                
                'WRITE OUT PERSON CANCELLED ITEMS NEXT
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "select service_code_id, service_code + '-' + REPLACE(REPLACE(service_desc,'&','&amp;'),'<','&lt;') as service_desc from billing_procedure_service_codes where [end] IS NULL and service_code IN('160','161','162','163','164','165') and service_type_id = 1 order by service_desc"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV      	            
      	        
      	        wfile3.Write ("<choice value=""" & "Person_Cancelled" & """ label=""" & "Person Cancelled" & """>" & chr(13) & chr(10))
      	            
  	            Do Until rsV.EOF
                     'Write out program level choice
                     wfile3.Write ("<choice value=""" & rsV("service_code_id") & """ label=""" & rsV("service_desc") & """></choice>" & chr(13) & chr(10))                              
                                                   
                rsV.MoveNext
                Loop
                
                wfile3.Write ("</choice>" & chr(13) & chr(10))
                
                'WRITE OUT PROVIDER CANCELLED ITEMS NEXT
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "select service_code_id, service_code + '-' + REPLACE(REPLACE(service_desc,'&','&amp;'),'<','&lt;') as service_desc from billing_procedure_service_codes where [end] IS NULL and service_code IN('160','161','162','163','164','165') and service_type_id = 1 order by service_desc"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV      	            
      	        
      	        wfile3.Write ("<choice value=""" & "Provider_Cancelled" & """ label=""" & "Provider Cancelled" & """>" & chr(13) & chr(10))
      	            
  	            Do Until rsV.EOF
                     'Write out program level choice
                     wfile3.Write ("<choice value=""" & rsV("service_code_id") & """ label=""" & rsV("service_desc") & """></choice>" & chr(13) & chr(10))                              
                                                   
                rsV.MoveNext
                Loop
                
                wfile3.Write ("</choice>" & chr(13) & chr(10))
                            
            elseif InStr(singleline,"<xforms:instance id=""Axis_1_List"" xmlns="""">") THEN
                wfile3.Write (singleline & chr(13) & chr(10))
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10))

                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "get_diag_codes 1"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	            
  	            Do Until rsV.EOF
        
                     wfile3.Write ("<choice value=""" & rsV("diag_id") & """>" & rsv("description") & "~" & rsV("diag_value") & "</choice>" & chr(13) & chr(10))
                     '<choice value="319">ABC-XYZ</choice>
                                
                rsV.MoveNext
                Loop
                
            elseif InStr(singleline,"<xforms:instance id=""Axis_2_List"" xmlns="""">") THEN
                wfile3.Write (singleline & chr(13) & chr(10))
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10))
                
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "get_diag_codes 2"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	            
  	            Do Until rsV.EOF
        
                     wfile3.Write ("<choice value=""" & rsV("diag_id") & """>" & rsv("description") & "~" & rsV("diag_value") & "</choice>" & chr(13) & chr(10))
                                
                rsV.MoveNext
                Loop
            
            elseif InStr(singleline,"<xforms:instance id=""Axis_3_List"" xmlns="""">") THEN
                wfile3.Write (singleline & chr(13) & chr(10))
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10))
                
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "get_diag_codes 3"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	            
  	            Do Until rsV.EOF
        
                     wfile3.Write ("<choice value=""" & rsV("diag_id") & """>" & rsv("description") & "~" & rsV("diag_value") & "</choice>" & chr(13) & chr(10))
                                
                rsV.MoveNext
                Loop
           
           elseif InStr(singleline,"<xforms:instance id=""Axis_4_List"" xmlns="""">") THEN
                wfile3.Write (singleline & chr(13) & chr(10))
                singleline=wfile.readline 
                wfile3.Write (singleline & chr(13) & chr(10))
                
                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "get_diag_codes 4"
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	             
  	            Do Until rsV.EOF
        
                     wfile3.Write ("<choice value=""" & rsV("diag_id") & """>" & rsv("description") & "~" & rsV("diag_value") & "</choice>" & chr(13) & chr(10))
                                
                rsV.MoveNext
                Loop         
                
            else
                wfile3.write(singleline & chr(13) & chr(10))
            end if
            
        loop               
        '********************* END OF MODEL SECTION ****************************              
          
        '**************************LAYOUT SECTION****************************
        do while not wfile.AtEndOfStream and model_end_found = 1
            counter=counter+1
            singleline=wfile.readline
              
            if InStr(singleline,"<data sid=""agency_logo_placeholder1"">") <> 0 THEN
                      'response.write "in it"
                        wfile3.Write(singleline & chr(13) & chr(10)) 
                        singleline=wfile.readline 
                        wfile3.Write(singleline & chr(13) & chr(10)) 
                        singleline=wfile.readline 
                        wfile3.Write(singleline & chr(13) & chr(10)) 
                        singleline=wfile.readline 

                        Do Until InStr(singleline, "</data>") <> 0
                            singleline = wfile.readline
                 
                        Loop
                
                        wfile3.Write ("<mimedata encoding=""base64"">" & logo_image_hash & "</mimedata>" & chr(13) & chr(10))
                    wfile3.Write ("</data>" & chr(13) & chr(10))    

            elseif InStr(singleline,"<mimedata encoding=") <> 0 and  Request.QueryString("ft") = "LOCATION_RECORD" THEN ' EFS and INTAKE need a picture

  	            Set SQLStmtComp = Server.CreateObject("ADODB.Command")
  	            Set rsComp = Server.CreateObject ("ADODB.Recordset")

  	            'SQLStmtComp.CommandText = "exec get_picture " &  Request.QueryString("cid")
  	            
  	                 
  	            SQLStmtComp.CommandText = "select CASE WHEN qs.picture IS NULL THEN 0 ELSE 1 END as ""exists"", " &_ 
" CAST('<mimedata encoding=""base64"">' + cast(N'' as xml).value('xs:base64Binary(xs:hexBinary(sql:column(""qs.picture"")))', 'varchar(max)') + '</mimedata>' as varchar(max)) as picture " &_
" from location_master as qs where location_id = " &  Request.QueryString("lid")

  	            SQLStmtComp.CommandType = 1
  	            Set SQLStmtComp.ActiveConnection = conn
  	            SQLStmtComp.CommandTimeout = 45 'Timeout per Command
  	            ' response.Write "sql = " & SQLStmtComp.CommandText

  	            rsComp.Open SQLStmtComp

                       if (rsComp("exists") = "1") and found_photo_at_admission1 = 1 THEN

                            ' now write out tables and loop through till end of table6 not above should really come from CBATIPE change once rob has it working

                          wfile3.Write(rsComp("picture") & chr(13) & chr(10) )
                          
                          found_photo_at_admission1 = 2

                          Do Until InStr(singleline,"</mimedata>") <> 0 or wfile.AtEndOfStream  
                               singleline=wfile.readline
                           Loop
  	   
                      else 
                        ' dont have data so write out table2 and go on
                        wfile3.Write(singleline & chr(13) & chr(10)) 
    
                     end if ' exists
                     
            elseif InStr(singleline,"<data sid=""photo_at_admission1"">") <> 0 THEN
                found_photo_at_admission1 = 1
                wfile3.Write(singleline & chr(13) & chr(10))
            elseif InStr(singleline,"</ufv_settings>") THEN
                wfile3.Write(Replace(singleline, "</ufv_settings>", "<mandatorycolor>" & mandatory_form_field_color & "</mandatorycolor>" & chr(13) & chr(10) & "<errorcolor>" & error_field_color & "</errorcolor>" & chr(13) & chr(10) & "</ufv_settings>" & chr(13) & chr(10) ))   
            
            elseif InStr(singleline,"</printsettings>") THEN
                    wfile3.Write(Replace(singleline, "</printsettings>", "<footer>" & chr(13) & chr(10) _
                       & "<left compute=""'&#xA;" & chr(13) & chr(10) _
                       & "PRINT FOOTER &#xA;" & chr(13) & chr(10) _
                       & "__________________________________________ &#xA;" & chr(13) & chr(10) _
                       & "Form Name: " & title_form_desc & " &#xA;" & chr(13) & chr(10) _
                       & "Group Name: " & REPLACE(cur_group_name,"'","\'") & " &#xA;" & chr(13) & chr(10) _
                       & "Created/Printed On:" & form_create_date & "/' +. viewer.printDate() +. ' ' +. viewer.printTime()"">" & chr(13) & chr(10) _
                       & "</left>" & chr(13) & chr(10) _
                       & "<center>" & chr(13) & chr(10) _
                       & " " & chr(13) & chr(10) _
                       & "   __________________________________________" & chr(13) & chr(10) _
                       & "   Responsible Program: " & REPLACE(cur_group_program_name,"'","\'") & chr(13) & chr(10) _
                       & "   Created By: " & REPLACE(cur_staff_name,"'","\'") & chr(13) & chr(10) _
                       & "   Printed By: " & REPLACE(cur_staff_name,"'","\'") & "</center>" & chr(13) & chr(10) _
                       & "<right compute=""'&#xA;" & chr(13) & chr(10) _
                       & " &#xA;" & chr(13) & chr(10) _
                       & "__________________________________________ &#xA;" & chr(13) & chr(10) _
                       & "Organizer: " & REPLACE(cur_group_organizer_name,"'","\'") & " &#xA;" & chr(13) & chr(10) _
                       & "Form Page: ' +. viewer.printFormPage() +. ' of ' +. viewer.printTotalFormPages() +. '" & " &#xA;" & chr(13) & chr(10) _
                       & "Printed Sheet: ' +. viewer.printPageSheet() +. ' of ' +. viewer.printTotalPageSheets() +. ' '"">" & chr(13) & chr(10) _
                       & "</right>" & chr(13) & chr(10) _
                       & "</footer>" & chr(13) & chr(10) & "</printsettings>" & chr(13) & chr(10)) )
                                                   
            elseif InStr(singleline,"</menu>") THEN
                   if local_form_save = "0" THEN
                        wfile3.Write(Replace(singleline, "</menu>", "<save>hidden</save></menu>") & chr(13) & chr(10))            
                   else
                        wfile3.Write(Replace(singleline, "</menu>", "<save>on</save></menu>") & chr(13) & chr(10)) 
                   end if
                        
            elseif InStr(singleline,"<popup") THEN
                cur_sid = ""
                
                'POSSIBLY DYNAMIC FIND SID AND LOOKUP IN CODEMAP
                popup_tag_start = InStr(singleline,"<popup sid=""")
                popup_tag_end = InStr(singleline,""">")
                total_tag_length = popup_tag_end - popup_tag_start
                cur_sid = Mid(singleline,(popup_tag_start+12),(total_tag_length-12))
                
                if InStrRev(cur_sid, "_") THEN
                    undloc = InStrRev(cur_sid,"_")
                    numcheck = Mid(cur_sid,undloc+1)
                    if isNumeric(numcheck) Then
                        stripped_sid = Mid(cur_sid, 1, undloc-1)
                       ' response.Write"found underscore, base value = " & stripped_sid
                        cur_sid = stripped_sid
                    end if
                end if           
                
                '************NEED SECTION FOR IAP*********************
                if cur_sid = "Med_Brand" or cur_sid = "Med_Brand1" THEN
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
                
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select medication_id, medication_brand_name from medications_master where is_active = 1 order by medication_brand_name"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
  	                
  	                Do Until rsV.EOF  	       
  	                    cur_id = rsV("medication_id")         
  	                    cur_desc = rsV("medication_brand_name")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_desc & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                                 wfile3.Write ("<xforms:label>Other</xforms:label>" & chr(13) & chr(10))
                                 wfile3.Write ("<xforms:value>Other</xforms:value>" & chr(13) & chr(10))
                                 wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                                    wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                                 wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                              wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                elseif cur_sid = "Linked_Assessed_Need" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_assessed_needs_for_client_and_form_for_iap " & Request.QueryString("cid") & "," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & "," & Request.QueryString("lfid")
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtV.CommandText
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_desc = rsV("Need_Desc")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & rsV("form_desc") & " - " & rsV("form_date") & " - " & cur_desc & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & rsV("Need_ID") & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                
                elseif cur_sid = "Linked_Assessed_Need2" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'REQUIREMENT CHECK IS CONFIG OPTION
                
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_assessed_needs_for_client_and_form_for_iap " & Request.QueryString("cid") & "," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & "," & Request.QueryString("lfid")
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF
  	                
  	                cur_desc = rsV("Need_Desc")
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & rsV("form_type") & " - " & rsV("form_date") & " - " & cur_desc & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & rsV("Need_ID") & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                
                elseif cur_sid = "Linked_Assessed_Need3" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'REQUIREMENT CHECK IS CONFIG OPTION
                
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_assessed_needs_for_client_and_form_for_iap " & Request.QueryString("cid") & "," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & "," & Request.QueryString("lfid")
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF
  	                
  	                cur_desc = rsV("Need_Desc")
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & rsV("form_type") & " - " & rsV("form_date") & " - " & cur_desc & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & rsV("Need_ID") & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                 '************NEED SECTION FOR BAP*********************
                elseif cur_sid = "CBHI_Linked_Assessed_Need"  and Request.QueryString("lfid") THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_cbhi_assessed_needs_for_client_and_form_for_bpa " & Request.QueryString("cid") & "," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & "," & Request.QueryString("lfid")
  	                SQLStmtV.CommandType = 1
  	                'response.Write "sql = " & SQLStmtV.CommandText
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_desc = rsV("Need_Desc")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & rsV("form_type") & " - " & rsV("form_date") & " - " & cur_desc & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & rsV("Need_ID") & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                '**********ALL ACTIVE PROGRAMS FOR CURRENT CLIENT**********************
                elseif cur_sid = "ClientProgram" THEN
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select Program_ID, Program_Name from program_master p where program_name != 'Intake' and is_active = 1 and exists(select 1 from client_program_assign where client_id = " & Request.QueryString("cid") & " and (end_date = '' or end_date IS NULL) and program_id = p.program_id)"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("Program_ID")
  	                    cur_name = rsV("Program_Name")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
           
                   '*********Locations for the Resp Program**********************
               elseif cur_sid = "Program_Location" THEN
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "<xforms:label>Select a Location</xforms:label>") <> 0
                        singleline = wfile.readline
                    Loop
                    
                    'WRITE OUT NEW VALUES               
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                'HILL REWRITE NEEDED FOR INTEGRATION*****
  	                SQLStmtV.CommandText = "select Location_id, Location_Desc from hill_Location_Master where location_id in(select location_id from hill_program_location where program_id = (select external_info from program_master where program_id = " & Request.QueryString("pid") & ") ) order by Location_Desc "
  	                'SQLStmtV.CommandText = "select Location_id, Location_Name from Location_Master where Location_id in(select location_id from program_location_assign where program_id = " & Request.QueryString("pid") & ")"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("Location_id")
  	                    cur_name = rsV("Location_Desc")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                
                    'WRITE OUT PLACE HOLDER START 
                    wfile3.Write("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                
                '**********ALL ACTIVE PROGRAMS**********************
                elseif cur_sid = "Program" THEN
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select Program_ID, Program_Name from program_master where is_active = 1"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("Program_ID")
  	                    cur_name = rsV("Program_Name")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                
                '********SECTION FOR INCIDENT FORM (SSMH)****************
                elseif cur_sid = "RU" THEN
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select Program_ID, RU_ID from program_master where is_active = 1 and RU_ID IS NOT NULL order by RU_ID"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("Program_ID")
  	                    cur_name = rsV("RU_ID")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                
                elseif cur_sid = "Staff_Involved" or cur_sid = "Staff_Involved1" THEN 'THIS IS FOR SSMH INCIDENT REPORT
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select staff_id, last_name, first_name from staff_master where is_active = 1"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("staff_id")
  	                    cur_name = rsV("last_name") & ", " & rsV("first_name")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                
                elseif cur_sid = "Therapist" or cur_sid = "Therapist_1" or cur_sid = "Intake_Coord" THEN 'THERAPIST LIST FOR 
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop                
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select staff_id, last_name, first_name from staff_master where is_active = 1 order by last_name, first_name"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("staff_id")
  	                    cur_name = rsV("last_name") & ", " & rsV("first_name")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                    
                elseif cur_sid = "NewStaff" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select staff_id, (last_name + ', ' + first_name) as staff_name from staff_master where staff_id IN(select staff_id from staff_program_assign where program_id = " & Request.QueryString("pid") & ") order by last_name, first_name"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("staff_id")
  	                    cur_name = rsV("staff_name") 
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                elseif cur_sid = "Add_Staff" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select staff_id, (last_name + ', ' + first_name) as staff_name from staff_master where staff_id IN(select cm.code_name from code_map cm, code_def cd where cd.form_sid = 'Add_Staff' and cd.code_type = cm.code_type) order by last_name, first_name"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("staff_id")
  	                    cur_name = rsV("staff_name") 
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                '******************START NEW HILL BS FIELDS
                elseif cur_sid = "Loc_Code" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select location_id, description from location_master order by description"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("location_id")
  	                    cur_desc = rsV("description")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_desc & " (" & cur_name & ")</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                elseif cur_sid = "Procedure_Code" and Request.QueryString("ft") <> "GPPN_GROUP" and Request.QueryString("ft") <> "NOTE_GROUP" THEN
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select service_code_id, service_code + '-' + REPLACE(REPLACE(service_desc,'&','&amp;'),'<','&lt;') as service_desc from billing_procedure_service_codes where [end] IS NULL and service_type_id = (select service_type_id from program_master where program_id = '" & Request.QueryString("pid") & "') order by service_desc"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("service_code_id")
  	                    cur_name = rsV("service_desc") 
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                    
              
                
                elseif cur_sid = "Eval_Location" THEN 'GET LOCATION LISTS FROM HILL TABL
                
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
             
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "select location_id, location_desc from hill_location_master order by Location_Desc"
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                'response.Write "sql = " & SQLStmtV.CommandText
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                Do Until rsV.EOF  	                
  	                    cur_id = rsV("location_id")
  	                    cur_name = rsV("location_desc")
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_name & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsV.MoveNext
                    Loop
                    
                elseif cur_sid = "Prescribing_MD" THEN
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                    singleline=wfile.readline 
                    wfile3.Write (singleline & chr(13) & chr(10))
                
                    'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop
                
                    Set SQLStmtPrescriber = Server.CreateObject("ADODB.Command")
  	                Set rsPres = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtPrescriber.CommandText = "select client_prescriber_id, prescriber_name from client_prescriber_master where client_id = " & Request.QueryString("cid") & " order by prescriber_name"
  	                SQLStmtPrescriber.CommandType = 1
  	                Set SQLStmtPrescriber.ActiveConnection = conn
  	                'response.Write("sql = " & SQLStmtPrescriber.CommandText)
  	                SQLStmtPrescriber.CommandTimeout = 45 'Timeout per Command
  	                rsPres.Open SQLStmtPrescriber
  	                
  	                Do Until rsPres.EOF  	
  	                    cur_client_prescriber_id = rsPres("client_prescriber_id")      
  	                    cur_prescriber = rsPres("prescriber_name")         
  	                    
  	                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:label>" & cur_prescriber & "</xforms:label>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:value>" & cur_client_prescriber_id & "</xforms:value>" & chr(13) & chr(10))
                        wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                        wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                        wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                        wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                        wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                        wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                        wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                        wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    rsPres.MoveNext
                    Loop
                        wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                                 wfile3.Write ("<xforms:label>New</xforms:label>" & chr(13) & chr(10))
                                 wfile3.Write ("<xforms:value>New</xforms:value>" & chr(13) & chr(10))
                                 wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                                    wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                                 wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                              wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write (singleline & chr(13) & chr(10))
                    
                '****************ALL OTHER CODE MAP LISTS SECTION******************
                else
                
                    Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_value_list_for_sid '" & cur_sid & "'" 
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
      	            
  	                if rsV.EOF THEN
  	                    wfile3.Write (singleline & chr(13) & chr(10))
  	                else
  	                    wfile3.Write (singleline & chr(13) & chr(10)) 
  	                    'SET FLAG AND LOOK FOR NEXT TAG OF <xforms:label></xforms:label> and end tag of </xforms:select1> and insert value list between
  	                    foundDynamicPopup = 1
  	                end if
  	            
  	            end if
  	        
  	        elseif foundDynamicPopup = 1 and InStr(singleline,"<xforms:label></xforms:label>") THEN                           
  	                wfile3.Write (singleline & chr(13) & chr(10))
  	                
  	                'FIND PLACE HOLDER AND SKIP ALL PREVIOUSLY GENERATED ITEMS IN THE DROPDOWN
                    Do Until InStr(singleline, "</xforms:select1>") <> 0
                        singleline = wfile.readline
                    Loop   
  	                
  	                Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	                Set rsV = Server.CreateObject ("ADODB.Recordset")
  	                SQLStmtV.CommandText = "get_value_list_for_sid '" & cur_sid & "'" 
  	                SQLStmtV.CommandType = 1
  	                Set SQLStmtV.ActiveConnection = conn
  	                SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	                rsV.Open SQLStmtV
  	                
  	                Do Until rsV.EOF
  	                    
  	                        wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                            wfile3.Write ("<xforms:label>" & rsV("short_desc") & "</xforms:label>" & chr(13) & chr(10))
                            wfile3.Write ("<xforms:value>" & rsV("code_name") & "</xforms:value>" & chr(13) & chr(10))
                            wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                            wfile3.Write ("<value compute=""label""></value>" & chr(13) & chr(10))
                            wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                            wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                            wfile3.Write ("<y>1</y>" & chr(13) & chr(10))
                            wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                            wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                            wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                            wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                            wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                            wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                            wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                                               
                    rsV.MoveNext
                    Loop
                    
                    wfile3.Write (singleline & chr(13) & chr(10))
                                        
                    foundDynamicPopup = 0
  	                
            elseif InStr(singleline,"<line sid=""NEEDS_BASE_LINE_1"">") THEN
  	            found_line_before_needs = 1
  	            wfile3.write(singleline & chr(13) & chr(10))
  	        
  	        elseif InStr(singleline,"</line>") and found_line_before_needs = 1 THEN
  	            wfile3.write(singleline & chr(13) & chr(10))
  	            
  	            need_counter = 1
  	            need_row_counter = 1
  	            
  	            if Request.QueryString("ft") = "CCAU" or Request.QueryString("ft") = "CBFSCCAU" THEN
  	                base_line_width = 881
  	            else
  	                base_line_width = 868
  	            end if
  	            
  	            Set SQLStmtV = Server.CreateObject("ADODB.Command")
  	            Set rsV = Server.CreateObject ("ADODB.Recordset")
  	            SQLStmtV.CommandText = "get_assessed_needs_for_client_and_form " & Request.QueryString("cid") & "," & iap_requires_finalized_needs & "," & iap_ignores_needs_older_than_one_year & "," & Request.QueryString("lfid")
  	            SQLStmtV.CommandType = 1
  	            Set SQLStmtV.ActiveConnection = conn
  	            response.Write "sql = " & SQLStmtV.CommandText
  	            SQLStmtV.CommandTimeout = 45 'Timeout per Command
  	            rsV.Open SQLStmtV
      	            
  	            Do Until rsV.EOF
  	            
  	            cur_parent_type = rsV("form_type")
  	            cur_parent_date = rsV("form_date")    
  	            cur_desc = rsV("Need_Desc")
  	            cur_status = rsV("Need_Status")
  	            
  	            wfile3.Write ("<field sid=""Previous_Prioritized_Needs_Desc_" & need_counter & """>" & chr(13) & chr(10))
                wfile3.Write ("<xforms:textarea ref=""instance('Generated')/page2/previous_prioritized_needs_desc_" & need_counter & """>" & chr(13) & chr(10))
                wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                wfile3.Write ("</xforms:textarea>" & chr(13) & chr(10))
                wfile3.Write ("<readonly>on</readonly>" & chr(13) & chr(10))
                wfile3.Write ("<border>off</border>" & chr(13) & chr(10))
                wfile3.Write ("<bgcolor>#EEEEEE</bgcolor>" & chr(13) & chr(10))
                wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("<below compute=""itemprevious""></below>" & chr(13) & chr(10))
                wfile3.Write ("<offsetx>12</offsetx>" & chr(13) & chr(10))
                wfile3.Write ("<width>567</width>" & chr(13) & chr(10))
                wfile3.Write ("<height compute=""toggle(value) == '1' or &#xA;" & chr(13) & chr(10))
                wfile3.Write ("viewer.measureHeight('pixels') > '38' ? viewer.measureHeight('pixels') > '38' ?  &#xA;" & chr(13) & chr(10))
                wfile3.Write ("viewer.measureHeight('pixels') : '38' : '38'"">38</height>" & chr(13) & chr(10))
                wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("<scrollvert>never</scrollvert>" & chr(13) & chr(10))
                wfile3.Write ("<scrollhoriz>wordwrap</scrollhoriz>" & chr(13) & chr(10))
                wfile3.Write ("</field>" & chr(13) & chr(10))
                
                if Request.QueryString("ft") = "CCAU" or Request.QueryString("ft") = "CBFSCCAU" THEN                
                    wfile3.Write ("<radiogroup sid=""Previous_Prioritized_Needs_Status_" & need_counter & """>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:select1 appearance=""full"" ref=""instance('Generated')/page2/previous_prioritized_needs_status_" & need_counter & """>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Active</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<x>10</x>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Person</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>20</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Family</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>22</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Deferred</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>28</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Referred</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>22</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:select1>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<width>54</width>" & chr(13) & chr(10))
                    wfile3.Write ("<after>Prioritized_Needs_Desc_" & need_counter & "</after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>96</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("<offsety>10</offsety>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("</radiogroup>" & chr(13) & chr(10))
                
                    
                     need_label = "LINE24"
                else
                    wfile3.Write ("<radiogroup sid=""Previous_Prioritized_Needs_Status_" & need_counter & """>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:select1 appearance=""full"" ref=""instance('Generated')/page2/previous_prioritized_needs_status_" & need_counter & """>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Active</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<x>1</x>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Person_Declined</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>38</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Deferred</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>38</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:label></xforms:label>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:value>Referred_Out</xforms:value>" & chr(13) & chr(10))
                    wfile3.Write ("<xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<after compute=""itemprevious""></after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>33</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                    wfile3.Write ("<size>9</size>" & chr(13) & chr(10))
                    wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                    wfile3.Write ("</labelfontinfo>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:extension>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:item>" & chr(13) & chr(10))
                    wfile3.Write ("</xforms:select1>" & chr(13) & chr(10))
                    wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("<width>54</width>" & chr(13) & chr(10))
                    wfile3.Write ("<after>Prioritized_Needs_Desc_" & need_counter & "</after>" & chr(13) & chr(10))
                    wfile3.Write ("<offsetx>88</offsetx>" & chr(13) & chr(10))
                    wfile3.Write ("<offsety>8</offsety>" & chr(13) & chr(10))
                    wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                    wfile3.Write ("</radiogroup>" & chr(13) & chr(10))
                
                    
                    need_label = "LINE35"
                end if
                
                wfile3.Write ("<line sid=""NEEDS_BASE_LINE_" & (need_counter+1) & """>" & chr(13) & chr(10))
                wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("<below>Previous_Prioritized_Needs_Desc_" & need_counter & "</below>" & chr(13) & chr(10))
                wfile3.Write ("<width>" & base_line_width & "</width>" & chr(13) & chr(10))
                wfile3.Write ("<height>1</height>" & chr(13) & chr(10))
                wfile3.Write ("<offsetx>-12</offsetx>" & chr(13) & chr(10))
                wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("</line>" & chr(13) & chr(10))
  	            
  	            need_counter = need_counter + 1
  	            
  	            rsV.MoveNext
  	            Loop
  	            
  	            if need_counter > 1 THEN
  	            wfile3.Write ("<label sid=""NEED_COUNT_LABEL"">" & chr(13) & chr(10))
                wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("<below>" & need_label & "</below>" & chr(13) & chr(10))
                wfile3.Write ("<offsetx>10</offsetx>" & chr(13) & chr(10))
                wfile3.Write ("<offsety>-5</offsety>" & chr(13) & chr(10))
                wfile3.Write ("<width>600</width>" & chr(13) & chr(10))
                wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("<value>First " & (need_counter-1) & " needs taken from " & cur_parent_type & "-" & cur_parent_date & "</value>" & chr(13) & chr(10))
                wfile3.Write ("<bgcolor>#EEEEEE</bgcolor>" & chr(13) & chr(10))
                wfile3.Write ("<fontinfo>" & chr(13) & chr(10))
                wfile3.Write ("<fontname>Arial</fontname>" & chr(13) & chr(10))
                wfile3.Write ("<size>8</size>" & chr(13) & chr(10))
                wfile3.Write ("<effect>bold</effect>" & chr(13) & chr(10))
                wfile3.Write ("</fontinfo>" & chr(13) & chr(10))
                wfile3.Write ("<visible>off</visible>" & chr(13) & chr(10))
                wfile3.Write ("<printvisible>off</printvisible>" & chr(13) & chr(10))
                wfile3.Write ("</label>" & chr(13) & chr(10))
  	            
  	            end if 
  	            
  	            wfile3.Write ("<line sid=""LINE52"">" & chr(13) & chr(10))
                wfile3.Write ("<itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("<below>NEEDS_BASE_LINE_" & need_counter & "</below>" & chr(13) & chr(10))
                wfile3.Write ("<offsetx>0</offsetx>" & chr(13) & chr(10))
                wfile3.Write ("<width>" & base_line_width & "</width>" & chr(13) & chr(10))
                wfile3.Write ("<height>1</height>" & chr(13) & chr(10))
                wfile3.Write ("<offsety>-7</offsety>" & chr(13) & chr(10))
                wfile3.Write ("</itemlocation>" & chr(13) & chr(10))
                wfile3.Write ("</line>" & chr(13) & chr(10))
  	            
  	            found_line_before_needs = 0  	            
                
            else
                wfile3.write(singleline & chr(13) & chr(10))
            end if
            
            loop 
            wfile.close 
            Set wfile=nothing 
            Set fs=nothing  
            Set wfile2=nothing
            Set fs2=nothing
            Set fs3=nothing
            Set wfile3=nothing
            
            Set fs = CreateObject("Scripting.FileSystemObject") 
  	        fileToOpen = form_root_path & "web_root\temp_forms\" & New_name & "_01.xfdl"
  	        Set wfile = fs.OpenTextFile(fileToOpen) 
  	        
            blob_val = wfile.ReadAll
            
            wfile.close
            Set wfile=nothing
            Set fs=nothing
            
  	        '*********************************************************************
      	end if
  	end if
  	
  	Set fs=Server.CreateObject("Scripting.FileSystemObject")
    if fs.FileExists(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl") then
        fs.DeleteFile(form_root_path & "web_root\temp_forms\" & New_name & ".xfdl")
    end if
    if fs.FileExists(form_root_path & "web_root\temp_forms\" & New_name & "_01.xfdl") then
        fs.DeleteFile(form_root_path & "web_root\temp_forms\" & New_name & "_01.xfdl")
    end if
    set fs=nothing
    
    conn.Close()
    Set conn = Nothing
%>