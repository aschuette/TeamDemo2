<!--#include file ="checkusr.asp"-->
<!--#include file ="vb.asp"-->
<!--#include file ="db.asp"-->
<%

Function IIf(condition,value1,value2)
	If condition Then IIf = value1 Else IIf = value2
End Function



'for each Item  in Request.form
'  Response.Write(item & "  " & Request(Item) & "<br>")
'next


id = request("id")
task = request("task")
pg = request("pg")
keyType = "person"
pgLoad = request("pgLoad")


if pgLoad = "" then pgLoad = now

doAudit = true

select case pg
case "names"
  pgname = "Names"
  bNew = true
  theTable = "person_names"
  theTableKey = "person_name_id"
  nullFields = "start_date|end_date"
  orderBy = "created_dt desc"
  pFields = "full_name|name_restriction|first_name|middle_name|last_name|name_suffix|nick_name|maiden_name|legal_name|phonetic_first_name|phonetic_last_name"
case "addresses", "emergency"
  bNew = true
  theTable = "person_address"
  theTableKey = "person_address_id"
  deleteField = "address_type"
  nullFields = "valid_from|valid_to"
  orderBy = "primary_flag desc, CASE WHEN active = 'Y' THEN 0 ELSE 1 END, address_type"
  if pg = "emergency" then
    pgname = "Emergency Contact"
    whereExtra = " AND address_type = 'Emergency contact' "
    noSave = "primary_flag"
  else
    pgname = "Addresses"
  end if
case "ethnicity"
  pgname = "Ethnicity"
  bNew = true
  theTable = "person_ethnicity"
  theTableKey = "person_ethnicity_id"
  deleteField = "ethnicity_id"
  orderBy = "primary_flag desc, ethnicity_id"
  pFields = "ethnic"
case "ethnicityNew"
  pgname = "Ethnicity [IPEDS]"
  bNew = false
  theTable = "person_ethnicity_new"
  theTableKey = "person_ethnicity_id"
  deleteField = "person_ethnicity_id"
  orderBy = "person_ethnicity_id"
  pFields = ""
  keyType = "person"
  'onlysSave ="person_id|hispanic_flag|amind_flag|asian_flag|black_flag|pacif_flag|white_flag"




case "citizenship"
  pgname = "Citizenship&nbsp;History"
  bNew = true
  theTable = "citizenship_history"
  theTableKey = "citizenship_history_id"
  orderBy = "created_dt"
  pFields = "citizenship|country_of_citizenship|country_of_origin|current_res_country|perm_res_country|visa_date|visa_notes|visa_number|visa_type"
case "family"
  pgname = "Family&nbsp;Info"
  bNew = true
  theTable = "family_info"
  theTableKey = "family_info_id"
  deleteField = "relation_type"
  orderBy = "relation_type, created_dt desc"
  pFields = ""
case "requirements"
  pgname = "Requirements"
  bNew = true
  theTable = "student_requirements"
  theTableKey = "student_requirement_id"
  doAudit = false
  deleteField = "requirement_id"
  nullFields = "completed_date|expires_date"
  orderBy = "(select requirement_name FROM requirements WHERE requirements.requirement_id = student_requirements.requirement_id)"
  if session("zone") <> "" then
    whereExtra = " AND student_requirements.requirement_id = (SELECT r.requirement_id FROM requirements r WHERE r.requirement_id = student_requirements.requirement_id AND (r.zone_name = '' OR r.zone_name = '" & session("zone") & "')) "
  end if
case "exams"
   pgname = "Exam&nbsp;Scores"
  bNew = true
  theTable = "student_exams"
  theTableKey = "student_exam_id"
  deleteField = "exam_id"
  nullFields = "application_date|exam_date|scores_recorded_date|reference_number|student_exam_doc_id"
  orderBy = "exam_date desc" '"(select exam_name from exams where exams.exam_id = student_exams.exam_id) desc"
case "other"
  pgname = "Other&nbsp;Info"
  theTable = "person"
  theTableKey = "person_id"
  onlySave = "harvard_id|ssn|date_of_birth|email_address|alt_email_address|marital|birthplace|dentpin|deceased|gender|archive_barcode|archive_box|archive_years|notes"
  bNew = false
case "ferpa"
  pgname = "FERPA&nbsp;Flags"
  bNew = false
  theTable = "person"
  theTableKey = "person_id"
  onlySave = "name_restriction|address_restriction|photo_restriction"
case "instance"
  pgname = "Student&nbsp;Instance"
  bNew = false
  theTable = "student_instance"
  theTableKey = "student_instance_id"
  nullFields = "actual_grad_date|expected_grad_date|matric_date|tuition_count|years_in_program|reg_status_date|current_class|year_of_study_id|cross_reg_school_id"
  noSave = "attempted_credits_calc_date|earned_credits_calc_date|study_away_credits|study_away_credits_calc_date|total_credits_attempted|total_credits_earned|transfer_credits|transfer_credits_calc_date"
  keyType = "instance"
case "instances"
  pgname = "Other&nbsp;Instances"
  bNew = false
  theTable = "student_instance"
  theTableKey = "student_instance_id"
  orderBy = "student_instance_id desc"
case "advisors"
  pgname = "Advisors"
  bNew = true
  theTable = "advisors"
  theTableKey = "advisor_id"
  nullFields = "start_date|end_date"
  deleteField = "role_id"
  keyType = "instance"
case "projects"
  pgname = "Projects"
  bNew = true
  theTable = "student_projects"
  theTableKey = "project_id"
  nullFields = "project_date"
  keyType = "instance"
case "reghist"
  pgname = "Registration&nbsp;History"
  bNew = true
  theTable = "student_reg_hist"
  theTableKey = "student_reg_hist_id"
  nullFields = "effective_start_date|effective_end_date"
  deleteField = "reg_status_id"
  orderBy = "effective_start_date desc, effective_end_date desc"
  keyType = "instance"
case "correspond"
  pgname = "Correspondences"
  bNew = true
  theTable = "correspondence"
  theTableKey = "correspondence_id"
  nullFields = "printed_date"
  deleteField = "letter_id"
  onlySave = "correspondence_id|student_instance_id|letter_id|status|printed_date"
  orderBy = "letter_id, printed_date"
  keyType = "instance"
case "correspond2"
  pgname = "Details"
  theTable = "correspondence"
  theTableKey = "correspondence_id"
  nullFields = "requested_date|scheduled_date|printed_date|copies|person_address_id"
  deleteField = "letter_id"
  orderBy = "letter_id, printed_date, requested_date"
  keyType = "instance"
  if request("cid") <> "" then
    whereExtra = " AND correspondence_id = " & checkstring(request("cid"),50)
  else
    whereExtra = " AND 1=0 "
  end if
case "degrees"
  pgname = "Degrees"
  bNew = true
  theTable = "student_degrees"
  theTableKey = "student_degree_id"
  nullFields = "start_date|end_date|completed_date|completed_month|completed_year|student_instance_id|degree_status_id|program_track_id|institution_id"
  deleteField = "degree_id"
  orderBy = "(case when (select d.degree_name from degrees d where d.degree_id = student_degrees.degree_id) = 'Non-degree' then -1000 else student_degree_id end) desc"
case "termbill"
  pgname = "Term&nbsp;Billing"
  bNew = true
  theTable = "term_billing"
  theTableKey = "term_billing_id"
  nullFields = "date_sent"
  deleteField = "possible_charge_id"
  orderBy = "(select t.start_date from terms as t, term_billing_charge_items as b where t.term_id = b.term_id and b.possible_charge_id = term_billing.possible_charge_id), date_sent, possible_charge_id"
  keyType = "instance"
  if session("zone") <> "" then
    whereExtra = " AND ( (select top 1 b.zone_name from term_billing_charge_items as b where b.possible_charge_id = term_billing.possible_charge_id) = '" & session("zone") & "') "
  end if
case "tracks"
  pgname = "Program&nbsp;Tracks"
  theTable = "student_program_track"
  theTableKey = "student_program_track_id"
  nullFields = "start_date|end_date|completed_date"
  orderBy = "active desc, start_date desc"
  keyType = "instance"
case else
  pgname = "Details"
  bNew = false
end select

set rs = Server.CreateObject("ADODB.RecordSet")

if task = "" then
  NoPageMargin = true
  PrintShell
  response.end
end if

if task = "upframe" then
  response.write("<html><body bgcolor=""F9F9F9"" onload=""parent.document.location='p_details.asp?pg=" & request("pg") & "&id=" & request("id") & "';""></body></html>")
  response.end
end if

name_restriction_style = ""
address_restriction_style = ""
sql = "select p.person_id, p.full_name, p.name_restriction, p.address_restriction from person as p, student_instance as i where p.person_id = i.person_id and i.student_instance_id = " & checkstring(id,50)
rs.open sql,conn,1,1
if not rs.eof then
  person_id = rs("person_id")
  full_name = rs("full_name")
  if rs("name_restriction") = "Y" then
    name_restriction_style = ";background-color:FFFF99"
  end if
  if rs("address_restriction") = "Y" then
    address_restriction_style = ";background-color:FFFF99"
  end if
end if
rs.close

rows = 0
if pg="ethnicityNew" then
	rows = 1
end if

cols = -1
if theTable <> "" then
  set rs = conn.OpenSchema(4, Array(dbName,empty,theTable,empty))
  colList = ""
  while not rs.eof
    colList = colList & rs("column_name") & "|"
    rs.movenext
  wend
  rs.close
  if colList <> "" then
    colList = left(colList,len(colList)-1)
    colArray = split(colList,"|")
    cols = ubound(colArray)
  end if
end if

if pFields <> "" then
  pFieldsArray = split(pFields,"|")
end if

noSave = "|" & noSave & "|" & theTableKey & "|created_by|created_dt|updated_by|updated_dt|"
if onlySave <> "" then onlySave = "|" & onlySave & "|"

if nullFields <> "" then nullFields = "|" & nullFields & "|"
nullpFields = "|date_of_birth|visa_date|"

select case task
case "edit"

  'added for aamcId
  if pg = "other" then
	aamcId = ""
	sqlA = "SELECT AMCAS_ID FROM Madris.dbo.Madris_Amcas_Lookup WHERE person_id = " & checkstring(person_id,50)
	rs.open sqlA,conn,1,1
   if not rs.eof then
		aamcId = rs("AMCAS_ID")
	end if
	rs.close
  end if


  if cols >= 0 then
    if orderBy <> "" then orderBy = " order by " & orderBy
    if keyType = "person" then
      sql = "select * from " & theTable & " where person_id = " & checkstring(person_id,50) & whereExtra & orderBy
    else
      sql = "select * from " & theTable & " where student_instance_id = " & checkstring(id,50) & whereExtra & orderBy
    end if
    rs.open sql,conn,1,1
    while not rs.eof
      rows = rows + 1
      for i = 0 to cols
        x = rs(colArray(i))
        if isnull(x) then x = ""
        execute(colArray(i) & "_" & rows & " = x")
        'execute(colArray(i) & "_" & rows & " = """ & replace(x,"""","""""") & """")
      next
      rs.movenext
    wend
    rs.close
  end if
  if pFields <> "" then
    sql = "select * from person where person_id = " & checkstring(person_id,50)
    rs.open sql,conn,1,1
    if not rs.eof then
      for i = 0 to ubound(pFieldsArray)
        x = rs(pFieldsArray(i))
        if isnull(x) then x = ""
        x = replace(x,"""","""""")
        execute(pFieldsArray(i) & "_0" & " = x")
      next
    end if
    rs.close
  end if
  if pg = "correspond2" and request("cid") = "" then
    rows = 1
  end if
  PrintDetails
case "add","refresh"
  if cols >= 0 then
    rows = toNum(request("rows"))
    for r = 1 to rows
      for i = 0 to cols
        x = request(colArray(i) & "_" & r)
        if isnull(x) then x = ""
        x = replace(x,"""","""""")
        execute(colArray(i) & "_" & r & " = x")
      next
    next
  end if
  if pFields <> "" then
    for i = 0 to ubound(pFieldsArray)
      x = request(pFieldsArray(i) & "_0")
      if isnull(x) then x = ""
      x = replace(x,"""","""""")
      execute(pFieldsArray(i) & "_0" & " = x")
    next
  end if
  for i = 0 to rows
    execute("hasChanged_" & i & " = """ & request("hasChanged_"&i) & """")
  next
  if task = "add" then
    rows = rows + 1
    execute("hasChanged_" & rows & " = 2")
    select case pg
    case "names", "citizenship"
      for i = 0 to ubound(pFieldsArray)
        x = eval(pFieldsArray(i) & "_0")
        execute(pFieldsArray(i) & "_" & rows & " = x")
      next
      execute("end_date_" & rows & " = """ & date & """")
      execute("notes_" & rows & " = visa_notes_0")
    case "ethnicity"
      execute("person_ethnicity_id_" & rows & " = ""-" & rows & """")
    case "addresses", "emergency"
      execute("person_address_id_" & rows & " = ""-" & rows & """")
      execute("valid_from_" & rows & " = ""NEW""")
    case else
    end select
    execute("person_id_" & rows & " = """ & person_id & """")
    execute("student_instance_id_" & rows & " = """ & id & """")
  end if
  if pg = "ethnicity" or pg = "addresses" then
    primary_flag = request("primary_flag")
    if primary_flag = "" or IsNull(primary_flag) then primary_flag = "N" end if
  end if
  PrintDetails

case "save"

  if cols >= 0 then
    rows = toNum(request("rows"))
    for r = 1 to rows
      for i = 0 to cols

        x = request(colArray(i) & "_" & r)

        if isnull(x) then x = ""
        if pg = "ethnicityNew" then
				'Response.write x & "- "
				if x = "" then x = "N"
				'Response.write x & "<BR>"
        end if
		  if pg = "instance" then
			if colArray(i) & "_" & r = "exclude_from_oasis_CPP_file_1" then
				if x = "" then x = "N"
			end if
		  end if

        execute(colArray(i) & "_" & r & " = x")
        'execute(colArray(i) & "_" & r & " = """ & replace(x,"""","""""") & """")
      next
    next
  end if
  if pFields <> "" then
    for i = 0 to ubound(pFieldsArray)
      x = request(pFieldsArray(i) & "_0")
      if isnull(x) then x = ""
      x = replace(x,"""","""""")
      execute(pFieldsArray(i) & "_0" & " = x")
    next
  end if
  for i = 0 to rows
    execute("hasChanged_" & i & " = """ & request("hasChanged_"&i) & """")
  next

  '*******************************************************************************
  if pg = "ethnicity" or pg = "addresses" then
    primary_flag = request("primary_flag")
    if primary_flag = "" or IsNull(primary_flag) then primary_flag = "N" end if
    for i = 1 to rows
      if eval(theTableKey & "_" & i) = primary_flag then
        execute("primary_flag_" & i & " = ""Y""")
      else
        execute("primary_flag_" & i & " = ""N""")
      end if
    next
    sql = "select " & theTableKey & " from " & theTable & " where primary_flag = 'Y' and person_id = " & checkstring(person_id,50)
    rs.open sql,conn,1,1
    if not rs.eof then
      if cstr(primary_flag) <> cstr(rs(theTableKey)) then
        for i = 1 to rows
          if cstr(eval(theTableKey & "_" & i)) = cstr(rs(theTableKey)) then
            execute("hasChanged_" & i & " = 1")
          end if
        next
      end if
    end if
    rs.close
  end if

  '*******************************************************************************
  if pg = "reghist" then
    for i = 1 to rows
      if eval("hasChanged_" & i) <> "" and eval("hasChanged_" & i) <> "0" and eval("reg_status_id_"&i) <> "" then
        sql = "select reg_status_name from reg_status where reg_status_id = " & checkstring(eval("reg_status_id_"&i),50)
        rs.open sql,conn,1,1
        if not rs.eof then execute("reg_status_name_"&i&" = """ & rs("reg_status_name") & """")
        rs.close
      end if
    next
  end if
  '*******************************************************************************

  if deleteField <> "" then
    for i = 1 to rows
      tempDoDelete = false
      if pg = "degrees" and eval(deleteField & "_" & i) = "-1" then tempDoDelete = true
      if pg <> "degrees" and (eval(deleteField & "_" & i) = "" or eval(deleteField & "_" & i) = "0") then tempDoDelete = true
      if tempDoDelete then
        if eval("hasChanged_" & i) = "1" then
          execute("hasChanged_" & i & " = 3")
        else
          execute("hasChanged_" & i & " = 0")
        end if
      end if
    next
  end if
  sql = ""
  sql = sql & " DECLARE @ErrorSave INT"
  sql = sql & " SET @ErrorSave = 0"
  sql = sql & " BEGIN TRANSACTION"
  if doAudit then
    if keyType = "person" then
      sql = sql & " IF EXISTS (SELECT * FROM " & theTable & " WHERE person_id = " & checkstring(person_id,50) & " AND updated_dt > " & checkstring(pgLoad,50) & ") SET @ErrorSave = 1"
    else
      sql = sql & " IF EXISTS (SELECT * FROM " & theTable & " WHERE student_instance_id = " & checkstring(id,50) & " AND updated_dt > " & checkstring(pgLoad,50) & ") SET @ErrorSave = 1"
    end if
  end if
  if pFields <> "" then
    sql = sql & " IF EXISTS (SELECT * FROM person WHERE person_id = " & checkstring(person_id,50) & " AND updated_dt > " & checkstring(pgLoad,50) & ") SET @ErrorSave = 1"
  end if
  sql = sql & " IF (@ErrorSave = 0)"
  sql = sql & " BEGIN"
  sql = sql & " SET @ErrorSave = @ErrorSave "
  '*******************************************************************************
  ' 2/11/2014 add saving of aamcId
  if Request.Form("aamcId") <> "" then
  'update if exists
	sql = sql & "UPDATE  Madris.dbo.Madris_Amcas_Lookup SET AMCAS_ID = CAST("	&   checkstring(Request.Form("aamcId"),50)   & " AS INT) WHERE person_id =  " & checkstring(person_id,50)
  'insert if doesn't exist
	sql = sql & " INSERT INTO Madris.dbo.Madris_Amcas_Lookup (person_id, harvard_id, AMCAS_ID, APPL_PERSON_ID) SELECT person_id, harvard_id, CAST(" &   checkstring(Request.Form("aamcId"),50)   & " AS INT), APPL_PERSON_ID = 0  FROM Madris.dbo.person  WHERE person_id =  " & checkstring(person_id,50) & " AND NOT EXISTS (SELECT person_id from  Madris.dbo.Madris_Amcas_Lookup WHERE person_id = " & checkstring(person_id,50) & ") " ' = (SELECT ISNULL(APPL_PERSON_ID,0) FROM Amcas_Prod.amcasys.APPLICANT_PERSON WHERE AAMC_ID = CAST(" &   checkstring(aamcId,50)   & " AS INT))

  end if
  for i = 1 to rows
	    if eval("hasChanged_"&i) = "1" then
      sql = sql & " UPDATE " & theTable & " SET "
      for j = 0 to cols
        if (instr(noSave,"|"&colArray(j)&"|") = 0) and ((onlySave = "") or (instr(onlySave,"|"&colArray(j)&"|") > 0)) then
          sql = sql & colArray(j) & " = " & checkstring2(eval(colArray(j)&"_"&i),(instr(nullFields,"|"&colArray(j)&"|") > 0)) & ", "
        end if
      next
      if doAudit then
        sql = sql & " updated_by = " & checkstring(session("user_id"),50) & ", updated_dt = " & checkstring(now,50)
      elseif cols > 0 then
        sql = left(sql,len(sql)-2)
      end if
      sql = sql & " WHERE " & theTableKey & " = " & checkstring(eval(theTableKey&"_"&i),50)
    elseif eval("hasChanged_"&i) = "2" then
      sql = sql & " INSERT INTO " & theTable & " ("
      for j = 0 to cols
        if (instr(noSave,"|"&colArray(j)&"|") = 0) and ((onlySave = "") or (instr(onlySave,"|"&colArray(j)&"|") > 0)) then
          sql = sql & colArray(j) & ", "
        end if
      next
      if doAudit then
        sql = sql & " created_by, created_dt, updated_by, updated_dt) VALUES ("
      else
        if cols > 0 then sql = left(sql,len(sql)-2)
        sql = sql & ") VALUES ("
      end if
      for j = 0 to cols
        if (instr(noSave,"|"&colArray(j)&"|") = 0) and ((onlySave = "") or (instr(onlySave,"|"&colArray(j)&"|") > 0)) then
          sql = sql & checkstring2(eval(colArray(j)&"_"&i),(instr(nullFields,"|"&colArray(j)&"|") > 0)) & ", "
        end if
      next
      if doAudit then
        sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ", "
        sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ")"
      else
        if cols > 0 then sql = left(sql,len(sql)-2)
        sql = sql & ")"
      end if
    elseif eval("hasChanged_"&i) = "3" then
      sql = sql & " DELETE FROM " & theTable & " WHERE " & theTableKey & " = " & checkstring(eval(theTableKey&"_"&i),50)
    end if
  next
  '*******************************************************************************
  if pFields <> "" and eval("hasChanged_0") = "1" then
    sql = sql & " UPDATE person SET "
    for i = 0 to ubound(pFieldsArray)
      sql = sql & pFieldsArray(i) & " = " & checkstring2(eval(pFieldsArray(i)&"_0"),(instr(nullpFields,"|"&pFieldsArray(i)&"|") > 0)) & ", "
    next
    sql = sql & " updated_by = " & checkstring(session("user_id"),50) & ", updated_dt = " & checkstring(now,50)
    sql = sql & " WHERE person_id = " & checkstring(person_id,50)
  end if
  '*******************************************************************************
  if pg = "instance" then
    sqltemp = "select coalesce(program_track_id,'0') program_track_id, coalesce(reg_status_id,'0') reg_status_id, year_of_study_id from student_instance where student_instance_id = " & checkstring(id,50)
    rs.open sqltemp,conn,1,1
    if not rs.eof then
      program_track_id = rs("program_track_id")
      reg_status_id = rs("reg_status_id")
      year_of_study_id = rs("year_of_study_id")
      rs.close
      if cstr(program_track_id) <> cstr(request("program_track_id_1")) then
        sql = sql & " UPDATE student_program_track SET active = '', end_date = " & checkstring(cdate(date-1),50) & ","
        sql = sql & " updated_by = " & checkstring(session("user_id"),50) & ", updated_dt = " & checkstring(now,50)
        sql = sql & " WHERE student_instance_id = " & checkstring(id,50)
        sql = sql & " INSERT INTO student_program_track (student_instance_id,program_track_id,active,completed_date,start_date,end_date,notes,created_by,created_dt,updated_by,updated_dt) "
        sql = sql & " VALUES (" & checkstring(id,50) & "," & checkstring2(request("program_track_id_1"),true) & ",'Y',null," & checkstring(date,50) & ",null,'',"
        sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ", "
        sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ")"
      end if
      sqltemp = "select program_id from programs where track_reg_hist = 'Y' and program_id = " & checkstring(request("program_id_1"),50)
      rs.open sqltemp,conn,1,1
      trackHist = not rs.eof
      rs.close
      if trackHist and (cstr(reg_status_id) <> cstr(request("reg_status_id_1"))) then
        theDate = date
        if request("reg_status_date_1") <> "" then
          theDate = cdate(request("reg_status_date_1"))
        end if
        term_id = "null"
        term_end = "null"
        'sqltemp = "select top 1 term_id, end_date from terms where start_date <= " & checkstring(date,50) & " and end_date >= " & checkstring(date,50)
        sqltemp = "term_id,null," & checkstring(request("program_id_1"),50) & "," & checkstring(year_of_study_id,50)
        sqltemp = "select top 1 term_id, dbo.GetRegHistDate('E',"&sqltemp&") end_date from terms where dbo.GetRegHistDate('S',"&sqltemp&") <= " & checkstring(theDate,50) & " and dbo.GetRegHistDate('E',"&sqltemp&") >= " & checkstring(theDate,50)
        rs.open sqltemp,conn,1,1
        if not rs.eof then
          term_id = checkstring2(rs("term_id"),true)
          term_end = checkstring2(rs("end_date"),true)
        end if
        rs.close
        reg_status_name = "''"
        sqltemp = "select reg_status_name from reg_status where reg_status_id = " & checkstring(request("reg_status_id_1"),50)
        rs.open sqltemp,conn,1,1
        if not rs.eof then
          reg_status_name = checkstring2(rs("reg_status_name"),false)
        end if
        rs.close
        if term_id <> "" and term_id <> "null" then
          sql = sql & " UPDATE student_reg_hist SET effective_end_date = " & checkstring(cdate(theDate-1),50) & " WHERE student_instance_id = " & checkstring(id,50) & " AND term_id = " & term_id & " and effective_end_date > " & checkstring(cdate(theDate-1),50)
        end if
        sql = sql & " INSERT INTO student_reg_hist (student_instance_id,term_id,reg_status_id,person_id,effective_start_date,effective_end_date,year_of_study_id,reg_status_name,notes,created_by,created_dt,updated_by,updated_dt) "
        sql = sql & " VALUES (" & checkstring(id,50) & "," & term_id & "," & checkstring2(request("reg_status_id_1"),true) & "," & checkstring2(person_id,true) & "," & checkstring(theDate,50) & "," & term_end & "," & checkstring2(year_of_study_id,true) & "," & reg_status_name & ",'',"
        sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ", "
        sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ")"
      end if
    else
      rs.close
    end if
  end if
  sql = sql & " END"
  sql = sql & " IF (@ErrorSave = 0)"
  sql = sql & "   COMMIT TRANSACTION"
  sql = sql & " ELSE"
  sql = sql & "   ROLLBACK TRANSACTION"
  sql = sql & " SELECT @ErrorSave AS theError"
  theError = 1
  'response.write sql & "<br>"
  'response.end
  rs.open sql,conn,1,1
  do until (rs is nothing)
    if (rs.state = 1) then
      if (not rs.eof) then
        theError = rs("theError")
      end if
    end if
    set rs = rs.NextRecordset()
  loop
  if pg = "correspond2" and ((request("cid") = "") or (hasChanged_1 = "3")) then
    response.redirect("p_details.asp?task=upframe&pg=correspond&id=" & server.urlencode(id))
  end if
  response.redirect("p_details.asp?task=edit&pg=" & server.urlencode(pg) & "&id=" & server.urlencode(id) & "&cid=" & server.urlencode(request("cid")) & "&message=" & theError)
  response.end
case else
end select
'*******************************************************************************

sub PrintDetails
%>
<html>
<META HTTP-EQUIV="Expires" Content="0">
<META HTTP-EQUIV="Pragma" Content="no-cache">
<META HTTP-EQUIV="Cache-Control" Content="Private">
<head>
<title>MADRIS</title>
<style type="text/css">
<!--
body { scrollbar-arrow-color: #990000; scrollbar-3dlight-color: #666666; scrollbar-highlight-color: #DDDDDD; scrollbar-face-color: #DDDDDD; scrollbar-shadow-color: #DDDDDD; scrollbar-track-color: #EEEEEE; scrollbar-darkshadow-color: #666666}
Textarea {font-family: Arial, Verdana}
A:link#alink {text-decoration: none;}
A:visited#alink {text-decoration: none;}
A:active#alink {text-decoration: none;}
A:hover#alink {text-decoration: underline; color: #990000}
A:link#alink2 {text-decoration: none; color: #990000}
A:visited#alink2 {text-decoration: none; color: #990000}
A:active#alink2 {text-decoration: none; color: #990000}
A:hover#alink2 {text-decoration: underline; color: #990000}
-->
</style>

<script>

function DoChange(which) {
  x = eval("document.dataform.hasChanged_"+which);
  if ((x.value == '')||(x.value == '0')) {
    x.value = 1;
  }
  parent.DoChange();
}
function ShowMessage(which) {
  if (which != '') {
    parent.ChangesMade = false;
    parent.document.footerform.bsave.style.fontWeight = 'normal';
  }
  if (which == '0') {alert('Your changes have been saved.');}
  if (which == '1') {alert('ERROR - Your modifications were NOT SAVED!\n(Someone else changed the data on this page.)');}
}
</script>
<!--#include file ="js_cal.asp"-->
<!--#include file ="formval.asp"-->
</head>
<body bgcolor="F9F9F9" text="000000" link="000000" vlink="000000" alink="000000" onLoad="ShowMessage('<%=request("message")%>');">
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td width=1><form name="dataform" method=POST action="p_details.asp">
<input type=hidden name="task" value="save">
<input type=hidden name="id" value="<%=server.htmlencode(id)%>">
<input type=hidden name="pg" value="<%=server.htmlencode(pg)%>">
<input type=hidden name="rows" value="<%=server.htmlencode(rows)%>">
<input type=hidden name="pgLoad" value="<%=server.htmlencode(pgLoad)%>">
</td>
<td align=center valign=middle>


<table border=0 cellspacing=0 cellpadding=0><tr><td align=center bgcolor=CCCCCC>
<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>

<font face='arial,helvetica'>

<font color=666666><b><%=full_name%></b></font><br>


<% select case pg %>

<% case "names" %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% for i = 0 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_name_id_<%=i%>" value="<%=eval("person_name_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="name_restriction_<%=i%>" value="<%=eval("name_restriction_"&i)%>">
<b><% if i = 0 then %>Primary Name:<% else %>Old Name <%=i%>:<% end if %></b><br>
<%
CanNotEdit = ""
BgStyle = name_restriction_style
if i > 0 then
  CanNotEdit = "ContentEditable=false"
  BgStyle = ";background-color:DDDDDD"
end if
%>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td rowspan=8>&nbsp;&nbsp;</td>
<td align=left><font face='arial,helvetica' size=-1>Last&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>First&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Middle&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Suffix</font></td>
</tr>
<tr>
<td align=left><input type=text <%=CanNotEdit%> name="last_name_<%=i%>" value="<%=eval("last_name_"&i)%>" size=15 maxlength=50 style="width:175px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text <%=CanNotEdit%> name="first_name_<%=i%>" value="<%=eval("first_name_"&i)%>" size=15 maxlength=50 style="width:175px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text <%=CanNotEdit%> name="middle_name_<%=i%>" value="<%=eval("middle_name_"&i)%>" size=15 maxlength=60 style="width:125px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text <%=CanNotEdit%> name="name_suffix_<%=i%>" value="<%=eval("name_suffix_"&i)%>" size=15 maxlength=50 style="width:100px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr><td colspan=7><img src="images/spacer.gif" border=0 width=1 height=1></td></tr>
<tr>
<td align=left><font face='arial,helvetica' size=-1>Phonetic&nbsp;Last&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Phonetic&nbsp;First&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Nick&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Maiden&nbsp;Name</font></td>
</tr>
<tr>
<td align=left><input type=text <%=CanNotEdit%> name="phonetic_last_name_<%=i%>" value="<%=eval("phonetic_last_name_"&i)%>" size=15 maxlength=150 style="width:175px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text <%=CanNotEdit%> name="phonetic_first_name_<%=i%>" value="<%=eval("phonetic_first_name_"&i)%>" size=15 maxlength=150 style="width:175px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text <%=CanNotEdit%> name="nick_name_<%=i%>" value="<%=eval("nick_name_"&i)%>" size=15 maxlength=50 style="width:125px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text <%=CanNotEdit%> name="maiden_name_<%=i%>" value="<%=eval("maiden_name_"&i)%>" size=15 maxlength=255 style="width:100px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr><td colspan=7><img src="images/spacer.gif" border=0 width=1 height=1></td></tr>
<tr>
<td align=left colspan=3><font face='arial,helvetica' size=-1>Full&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left colspan=3><font face='arial,helvetica' size=-1>Legal&nbsp;Name</font></td>
</tr>
<tr>
<td align=left colspan=3><input type=text <%=CanNotEdit%> name="full_name_<%=i%>" value="<%=eval("full_name_"&i)%>" size=15 maxlength=255 style="width:355px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=left colspan=3><input type=text <%=CanNotEdit%> name="legal_name_<%=i%>" value="<%=eval("legal_name_"&i)%>" size=15 maxlength=100 style="width:230px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<% if i > 0 then %>
<tr><td colspan=8><img src="images/spacer.gif" border=0 width=1 height=1></td></tr>
<tr>
<td rowspan=2>&nbsp;&nbsp;</td>
<td align=left colspan=3><font face='arial,helvetica' size=-1>Notes</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Start&nbsp;Date</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>End&nbsp;Date</font></td>
</tr>
<tr>
<td align=left colspan=3><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:355px" onchange="DoChange(<%=i%>);"></td>
<td align=left><input type=text name="start_date_<%=i%>" value="<%=eval("start_date_"&i)%>" size=15 maxlength=50 style="width:100px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.start_date_"&i, "DoChange("&i&")"%></td>
<td align=left><input type=text name="end_date_<%=i%>" value="<%=eval("end_date_"&i)%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.end_date_"&i, "DoChange("&i&")"%></td>
</tr>
<% end if %>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=10><br><% end if %>
<% next %>






<% case "addresses", "emergency" %>

<%
sql = "select abbrev from my_states order by abbrev"
rs.open sql,conn,1,1
while not rs.eof
  state_list = state_list & rs("abbrev") & "|"
  rs.movenext
wend
rs.close
states = split(state_list,"|")
%>
<script>
function PickCountry(x) {
  eval("document.dataform."+x+".value = '';");
  window.open("g_country.asp?fn="+x,"CountryWindow","width=400,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no <% if pg = "emergency" then %>emergency contacts<% else %>addresses<% end if %> listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<% BgStyle = address_restriction_style %>
<% if i > 1 then %><hr noshade size=1><% end if %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<b>Address <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<input type=hidden name="person_address_id_<%=i%>" value="<%=eval("person_address_id_"&i)%>">
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Address&nbsp;Type:&nbsp;</font></td>
<td align=left><select name="address_type_<%=i%>" style="width:175px<%=BgStyle%>" onchange="DoChange(<%=i%>);parent.DoRefresh(<%=i%>);">
<option value="" <% if eval("address_type_"&i) = "" then %>selected<% end if %>>(delete)</option>
<%
sql = "select code from lookups where lookup_type = 'address_types' AND is_active='Y' AND (zone_name=" & checkstring(session("zone"),50) & " OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if rs("code") = eval("address_type_"&i) then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>


<!--<% if pg = "emergency" then %>
<option value="Emergency contact" <% if eval("address_type_"&i) = "Emergency contact" then %>selected<% end if %>>Emergency contact</option>
<% else %>
<option value="Alternate Mailing" <% if eval("address_type_"&i) = "Alternate Mailing" then %>selected<% end if %>>Alternate Mailing</option>
<option value="Emergency contact" <% if eval("address_type_"&i) = "Emergency contact" then %>selected<% end if %>>Emergency contact</option>
<option value="Exclerk Certification" <% if eval("address_type_"&i) = "Exclerk Certification" then %>selected<% end if %>>Exclerk Certification</option>
<option value="Father Address" <% if eval("address_type_"&i) = "Father Address" then %>selected<% end if %>>Father Address</option>
<option value="Gradesheet address" <% if eval("address_type_"&i) = "Gradesheet address" then %>selected<% end if %>>Gradesheet address</option>
<option value="Local/Mailing" <% if eval("address_type_"&i) = "Local/Mailing" then %>selected<% end if %>>Local/Mailing</option>
<option value="Mother Address" <% if eval("address_type_"&i) = "Mother Address" then %>selected<% end if %>>Mother Address</option>
<option value="Permanent" <% if eval("address_type_"&i) = "Permanent" then %>selected<% end if %>>Permanent</option>
<% end if %>-->

</select>
</td>
<td align=right><img src="images/spacer.gif" border=0 width=50 height=1></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Active:&nbsp;</font></td>
<td align=left colspan=2><nobr><font face='arial,helvetica' size=-1>
<input type=checkbox name="active_<%=i%>" value="Y" <% if eval("active_"&i) = "Y" then %>checked<% end if %> onchange="DoChange(<%=i%>);">
<% if pg <> "emergency" then %>
&nbsp;&nbsp;&nbsp;&nbsp;Primary:&nbsp;<input type=radio name="primary_flag" value="<%=eval("person_address_id_"&i)%>" <% if (primary_flag = "" and eval("primary_flag_"&i) = "Y") or (primary_flag <> "" and primary_flag = eval("person_address_id_"&i)) then %>checked<% end if %> onchange="DoChange(<%=i%>);">
<% end if %>
</font></nobr></td>
</tr>
</table>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Line&nbsp;1:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="line1_<%=i%>" value="<%=eval("line1_"&i)%>" size=15 style="width:275px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>Valid&nbsp;From:&nbsp;</font></td>
<td align=left><input type=text name="valid_from_<%=i%>" value="<%if eval("valid_from_"&i)="NEW" then response.write(formatdatetime(now,2)) else response.write(eval("valid_from_"&i)) end if%>" size=15 style="width:90px<%=BgStyle%>" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.valid_from_"&i, "DoChange("&i&")"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Line&nbsp;2:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="line2_<%=i%>" value="<%=eval("line2_"&i)%>" size=15 style="width:275px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Valid&nbsp;To:&nbsp;</font></td>
<td align=left><input type=text name="valid_to_<%=i%>" value="<%=eval("valid_to_"&i)%>" size=15 style="width:90px<%=BgStyle%>" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.valid_to_"&i, "DoChange("&i&")"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Line&nbsp;3:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="line3_<%=i%>" value="<%=eval("line3_"&i)%>" size=15 style="width:275px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>Phone:&nbsp;</font></td>
<td align=left><input type=text name="phone_number_<%=i%>" value="<%=eval("phone_number_"&i)%>" size=15 style="width:115px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>City:&nbsp;</font></td>
<td align=left><input type=text name="city_<%=i%>" value="<%=eval("city_"&i)%>" size=15 style="width:100px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>State:&nbsp;</font></td>
<!--
<td align=left><input type=text name="state_<%=i%>" value="<%=eval("state_"&i)%>" size=15 style="width:100px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
-->
<td align=left><select name="state_<%=i%>" style="width:100px<%=BgStyle%>" onchange="DoChange(<%=i%>);"><option value=""></option>
<%
for j = 0 to ubound(states)-1
  isselected = ""
  if states(j) = eval("state_"&i) then isselected = " selected"
  response.write("<option value=""" & states(j) & """" & isselected & ">" & states(j) & "</option>")
next
%>
</select>
</td>


<td align=right><font face='arial,helvetica' size=-1>Zip:&nbsp;</font></td>
<td align=left><input type=text name="zip_<%=i%>" value="<%=eval("zip_"&i)%>" size=15 style="width:115px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Province:&nbsp;</font></td>
<td align=left><input type=text name="province_<%=i%>" value="<%=eval("province_"&i)%>" size=15 style="width:100px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Country:&nbsp;</font></td>
<td align=left><input type=text name="country_<%=i%>" value="<%=eval("country_"&i)%>" size=15 ContentEditable="false" style="width:80px<%=BgStyle%>" onchange="DoChange(<%=i%>);"><a href="JavaScript:PickCountry('country_<%=i%>');DoChange(<%=i%>);"><img src="images/search.gif" border=0></a></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;Foreign&nbsp;Other:&nbsp;</font></td>
<td align=left><input type=text name="foreign_other_<%=i%>" value="<%=eval("foreign_other_"&i)%>" size=15 style="width:115px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<% if eval("address_type_"&i) = "Exclerk Certification" then %>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Cert&nbsp;Institution:&nbsp;</font></td>
<td align=left><input type=text name="cert_institution_<%=i%>" value="<%=eval("cert_institution_"&i)%>" size=15 style="width:440px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Cert&nbsp;Name:&nbsp;</font></td>
<td align=left><input type=text name="cert_name_<%=i%>" value="<%=eval("cert_name_"&i)%>" size=15 style="width:440px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Cert&nbsp;Person&nbsp;Title:&nbsp;</font></td>
<td align=left><input type=text name="cert_person_title_<%=i%>" value="<%=eval("cert_person_title_"&i)%>" size=15 style="width:440px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Cert&nbsp;Email:&nbsp;</font></td>
<td align=left><input type=text name="cert_email_<%=i%>" value="<%=eval("cert_email_"&i)%>" size=15 style="width:440px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<% end if %>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Notes:&nbsp;</font></td>
<td align=left><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:500px<%=BgStyle%>" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
</td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=10><br><% end if %>
<% next %>






<% case "ethnicity" %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<b>Ethnicity <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<input type=hidden name="person_ethnicity_id_<%=i%>" value="<%=eval("person_ethnicity_id_"&i)%>">
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Ethnicity:&nbsp;</font></td>
<td align=left><select name="ethnicity_id_<%=i%>" style="width:200px" onchange="DoChange(<%=i%>);"><option value="">(delete)</option>
<%
sql = "select distinct ethnicity_id, ethnicity_code from ethnicity where ethnicity_code is not null order by ethnicity_code"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if toNum(rs("ethnicity_id")) = toNum(eval("ethnicity_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & rs("ethnicity_id") & """" & isselected & ">" & rs("ethnicity_code") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Primary:&nbsp;</font></td>
<td align=left colspan=2><nobr><font face='arial,helvetica' size=-1><input type=radio name="primary_flag" value="<%=eval("person_ethnicity_id_"&i)%>" <% if (primary_flag = "" and eval("primary_flag_"&i) = "Y") or (primary_flag <> "" and primary_flag = eval("person_ethnicity_id_"&i)) then %>checked<% end if %> onchange="DoChange(<%=i%>);"></font></td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=10><br><% end if %>
<% next %>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<input type=hidden name="hasChanged_0" value="<%=hasChanged_0%>">
<b>Other Ethnicity</b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Ethnic:&nbsp;</font></td>
<td align=left><input type=text name="ethnic_0" value="<%=ethnic_0%>" size=15 maxlength=30 style="width:300px" onchange="DoChange(0);"></td>
</tr>
</table>



<% case "ethnicityNew" %>
<%
hChanged = 2
sql = "SELECT s.student_instance_id, s.person_id,  e.person_ethnicity_id,  e.hispanic_flag, e.amind_flag, e.asian_flag, e.black_flag, e.pacif_flag, e.white_flag, e.created_by, e.created_dt, e.updated_by, e.updated_dt FROM student_instance S LEFT OUTER JOIN dbo.person_ethnicity_new e on s.person_id = e.person_id  WHERE s.student_instance_id = " & Checkstring(id,50)
rs.open sql,conn,1,1
while not rs.eof
	hispanic_flag = rs("hispanic_flag")
	amind_flag = rs("amind_flag")
	asian_flag = rs("asian_flag")
	black_flag = rs("black_flag")
	pacif_flag = rs("pacif_flag")
	white_flag = rs("white_flag")
	person_id_x = rs("person_id")
	person_ethnicity_id_1  =  rs("person_ethnicity_id")
	if rs("person_ethnicity_id") <> "" then
		hChanged = 1
	end if
	rs.movenext
wend
rs.close

dim NullString
NullString = ""

%>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% i = 1  %>

<input type=hidden name="person_id_<%=i%>" value="<%=person_id_x %>">
<b>Ethnicity [IPEDS]</b><br><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<input type=hidden name="person_ethnicity_id_<%=i%>" value="<%=person_ethnicity_id_1%>">
<table border=0 cellspacing=0 cellpadding=0>
<tr>


<table border=0 cellspacing=0 cellpadding=0>
<tr>
	<td align=left><b><font face='arial,helvetica' size=-1>Do you consider yourself to be Hispanic/Latino &nbsp;<br /> </b>
		&nbsp;&nbsp;<input name="hispanic_flag_<%=i%>" type="radio" value="Y"  onClick="DoChange(1)"  <%= iif(hispanic_flag = "Y","checked",NullString)  %>/>&nbsp;Yes<br />
		&nbsp;&nbsp;<input name="hispanic_flag_<%=i%>"  type="radio" value ="N" onClick="DoChange(1)"   <%= iif(hispanic_flag = "N","checked",NullString)  %> />&nbsp;No<br /></font></td>
</tr>
<tr>
	<td align=left><b><font face='arial,helvetica' size=-1>In Addition, select one or more of the following racial categories to describe yourself:</font></b></td>
</tr>
<tr><td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input name="amind_flag_<%=i%>" type="checkbox" onClick="DoChange(1)" <%= iif(amind_flag = "Y","checked",NullString)  %> value="Y" />&nbsp;American Indian or Alaska Native</font></td></tr>
<tr><td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input name="asian_flag_<%=i%>" type="checkbox" onClick="DoChange(1)" <%= iif(asian_flag = "Y","checked",NullString)  %> value="Y" />&nbsp;Asian</font></td></tr>
<tr><td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input name="black_flag_<%=i%>" type="checkbox" onClick="DoChange(1)" <%= iif(black_flag = "Y","checked",NullString)  %> value="Y" />&nbsp;Black or African American</font></td></tr>
<tr><td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input name="pacif_flag_<%=i%>" type="checkbox" onClick="DoChange(1)" <%= iif(pacif_flag = "Y","checked",NullString)  %> value="Y" />&nbsp;Native Hawaiian or Pacific Islander</font></td></tr>
<tr><td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input name="white_flag_<%=i%>" type="checkbox" onClick="DoChange(1)" <%= iif(white_flag = "Y","checked",NullString)  %> value="Y" />&nbsp;White</font></td></tr>

</table>
<img src="images/spacer.gif" border=0 width=1 height=10><br>

</table>
<input type=hidden name="hasChanged_<%=i%>" value="<%= hChanged %>">



<% case "citizenship" %>

<script>
function PickCountry(x) {
  eval("document.dataform."+x+".value = '';");
  window.open("g_country.asp?fn="+x,"CountryWindow","width=400,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<input type=hidden name="hasChanged_0" value="<%=hasChanged_0%>">
<input type=hidden name="person_id_0" value="<%=person_id_0%>">
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=4><font face='arial,helvetica' size=-1><b>Current Citizenship</b><br><img src="images/spacer.gif" border=0 width=1 height=2></font></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Citizenship:&nbsp;</font></td>
<td align=left><select name="citizenship_0" style="width:200px" onchange="DoChange(0)"><option value=""></option>
<%
sql = "select code from lookups where lookup_type = 'citizenship' AND is_active='Y' AND (zone_name=" & checkstring(session("zone"),50) & " OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if rs("code") = citizenship_0 then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Visa&nbsp;Type:&nbsp;</font></td>
<td align=left><select name="visa_type_0" style="width:100px" onchange="DoChange(0)"><option value=""></option>
<%
sql = "select code from lookups where lookup_type = 'visa_type' AND is_active='Y' AND (zone_name=" & checkstring(session("zone"),50) & " OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if rs("code") = visa_type_0 then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Country&nbsp;Of&nbsp;Citizenship:&nbsp;</font></td>
<td align=left><input type=text name="country_of_citizenship_0" value="<%=country_of_citizenship_0%>" size=15 ContentEditable="false" maxlength=30 style="width:175px" onchange="DoChange(0);"><a href="JavaScript:PickCountry('country_of_citizenship_0');DoChange(0);"><img src="images/search.gif" border=0></a></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Visa&nbsp;Date:&nbsp;</font></td>
<td align=left><input type=text name="visa_date_0" value="<%=visa_date_0%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(0);"><%DrawCal "dataform.visa_date_0", "DoChange(0)"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Current&nbsp;Res&nbsp;Country:&nbsp;</font></td>
<td align=left><input type=text name="current_res_country_0" value="<%=current_res_country_0%>" size=15 ContentEditable="false" maxlength=30 style="width:175px" onchange="DoChange(0);"><a href="JavaScript:PickCountry('current_res_country_0');DoChange(0);"><img src="images/search.gif" border=0></a></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Visa&nbsp;Number:&nbsp;</font></td>
<td align=left><input type=text name="visa_number_0" value="<%=visa_number_0%>" size=15 maxlength=20 style="width:100px" onchange="DoChange(0);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Perm&nbsp;Res&nbsp;Country:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="perm_res_country_0" value="<%=perm_res_country_0%>" size=15 ContentEditable="false" maxlength=30 style="width:175px" onchange="DoChange(0);"><a href="JavaScript:PickCountry('perm_res_country_0');DoChange(0);"><img src="images/search.gif" border=0></a></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Country&nbsp;of&nbsp;Origin:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="country_of_origin_0" value="<%=country_of_origin_0%>" size=15 ContentEditable="false" maxlength=150 style="width:365px" onchange="DoChange(0);"><a href="JavaScript:PickCountry('country_of_origin_0');DoChange(0);"><img src="images/search.gif" border=0></a></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Visa&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="visa_notes_0" value="<%=visa_notes_0%>" size=15 style="width:390px" onchange="DoChange(0);"></td>
</tr>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="citizenship_history_id_<%=i%>" value="<%=eval("citizenship_history_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<tr>
<td align=left colspan=4><font face='arial,helvetica' size=-1><img src="images/spacer.gif" border=0 width=1 height=10><br><b>Citizenship History <%=i%></b><br><img src="images/spacer.gif" border=0 width=1 height=2></font></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Citizenship:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="citizenship_<%=i%>" value="<%=eval("citizenship_"&i)%>" size=15 maxlength=30 style="width:200px;background-color:DDDDDD" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Visa&nbsp;Type:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="visa_type_<%=i%>" value="<%=eval("visa_type_"&i)%>" size=15 maxlength=30 style="width:100px;background-color:DDDDDD" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Country&nbsp;Of&nbsp;Citizenship:&nbsp;</font></td>
<td align=left colspan=3><input type=text ContentEditable=false name="country_of_citizenship_<%=i%>" value="<%=eval("country_of_citizenship_"&i)%>" size=15 maxlength=30 style="width:200px;background-color:DDDDDD" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Visa&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:390px" onchange="DoChange(<%=i%>);"></td>
</tr>
<% next %>
</table>






<% case "family" %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no relations listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="family_info_id_<%=i%>" value="<%=eval("family_info_id_"&i)%>">
<b>Relation <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Relation&nbsp;Type:&nbsp;</font></td>
<td align=left><select name="relation_type_<%=i%>" style="width:150px" onchange="DoChange(<%=i%>);">
<option value="" <% if eval("relation_type_"&i) = "" then %>selected<% end if %>>(delete)</option>
<option value="Mother" <% if eval("relation_type_"&i) = "Mother" then %>selected<% end if %>>Mother</option>
<option value="Father" <% if eval("relation_type_"&i) = "Father" then %>selected<% end if %>>Father</option>
<option value="Spouse" <% if eval("relation_type_"&i) = "Spouse" then %>selected<% end if %>>Spouse</option>
<option value="Child 1" <% if eval("relation_type_"&i) = "Child 1" then %>selected<% end if %>>Child 1</option>
<option value="Child 2" <% if eval("relation_type_"&i) = "Child 2" then %>selected<% end if %>>Child 2</option>
<option value="Child 3" <% if eval("relation_type_"&i) = "Child 3" then %>selected<% end if %>>Child 3</option>
<option value="Child 4" <% if eval("relation_type_"&i) = "Child 4" then %>selected<% end if %>>Child 4</option>
<option value="Child 5" <% if eval("relation_type_"&i) = "Child 5" then %>selected<% end if %>>Child 5</option>
</select>
</td>
</tr>
</table>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left><font face='arial,helvetica' size=-1>Last&nbsp;Name:&nbsp;</font></td>
<td align=left><input type=text name="last_name_<%=i%>" value="<%=eval("last_name_"&i)%>" size=15 maxlength=50 style="width:170px" onchange="DoChange(<%=i%>);"></td>
<td align=left><font face='arial,helvetica' size=-1>&nbsp;&nbsp;First&nbsp;Name:&nbsp;</font></td>
<td align=left><input type=text name="first_name_<%=i%>" value="<%=eval("first_name_"&i)%>" size=15 maxlength=50 style="width:170px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Line&nbsp;1:&nbsp;</font></td>
<td align=left><input type=text name="line1_<%=i%>" value="<%=eval("line1_"&i)%>" size=15 maxlength=60 style="width:200px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;City:&nbsp;</font></td>
<td align=left><input type=text name="city_<%=i%>" value="<%=eval("city_"&i)%>" size=15 maxlength=30 style="width:100px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Country:&nbsp;</font></td>
<td align=left><input type=text name="country_<%=i%>" value="<%=eval("country_"&i)%>" size=15 maxlength=50 style="width:110px" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Line&nbsp;2:&nbsp;</font></td>
<td align=left><input type=text name="line2_<%=i%>" value="<%=eval("line2_"&i)%>" size=15 maxlength=60 style="width:200px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;State:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="state_<%=i%>" value="<%=eval("state_"&i)%>" size=15 maxlength=2 style="width:100px" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Line&nbsp;3:&nbsp;</font></td>
<td align=left><input type=text name="line3_<%=i%>" value="<%=eval("line3_"&i)%>" size=15 maxlength=60 style="width:200px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Zip:&nbsp;</font></td>
<td align=left><input type=text name="zip_<%=i%>" value="<%=eval("zip_"&i)%>" size=15 maxlength=20 style="width:100px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Phone:&nbsp;</font></td>
<td align=left><input type=text name="telephone_<%=i%>" value="<%=eval("telephone_"&i)%>" size=15 maxlength=20 style="width:110px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
</td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=10><br><% end if %>
<% next %>





<% case "requirements" %>

<%
valid_days = ""
valid_days2 = ""
sql = "select * from requirements"
if session("zone") <> "" then sql = sql & " where zone_name='' OR zone_name = " & checkstring(session("zone"),50)
sql = sql & "  order by requirement_name"
rs.open sql,conn,1,1
while not rs.eof
  requirement_id_list = requirement_id_list & rs("requirement_id") & "|"
  requirement_name_list = requirement_name_list & rs("requirement_name") & "|"
  if rs("valid_days") <> "" then
    valid_days = valid_days & " if (x == '" & rs("requirement_id") & "') {e.value = '" & (date + rs("valid_days")) & "';} "
    valid_days2 = valid_days2 & " if (x == '" & rs("requirement_id") & "') {e.value = AddDays(d.value," & rs("valid_days") & ");} "
    end if
  rs.movenext
wend
rs.close
requirement_ids = split(requirement_id_list,"|")
requirement_names = split(requirement_name_list,"|")
requirements = ubound(requirement_ids)-1
%>
<script>
function CheckCompleted(i) {
  r = eval('document.dataform.requirement_id_'+i);
  c = eval('document.dataform.completed_'+i);
  d = eval('document.dataform.completed_date_'+i);
  e = eval('document.dataform.expires_date_'+i);
  x = r.options[r.selectedIndex].value;
  if (c.checked) {
    d.value = '<%=date%>';
    <%=valid_days%>
  } else {
    d.value = '';
    e.value = '';
  }
}

function ChngCompleted(i) {
  r = eval('document.dataform.requirement_id_'+i);
  c = eval('document.dataform.completed_'+i);
  d = eval('document.dataform.completed_date_'+i);
  e = eval('document.dataform.expires_date_'+i);
  x = r.options[r.selectedIndex].value;
  if (d.value!='' || e.value != '') {
     <%=valid_days2%>
   } else {
    e.value = '';
  }
}

function AddDays(startDate,numDays) {
 var d = new Date;
 if (startDate == '2/29/00') {
  startDate = '2/29/2000';
 }
 var m = Date.parse(startDate);
 d.setTime(m);
 if (d.getYear() < 100) {
  if (m >= Date.parse('3/1/00')) {
   m += (100*365+25)*24*60*60*1000;
  } else {
   m += (100*365+24)*24*60*60*1000;
  }
 }
 d.setTime(m + numDays*24*60*60*1000);
 var s = (d.getMonth()+1) + '/' + d.getDate() + '/' + d.getFullYear();
 return s;
}

</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no requirements listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="student_requirement_id_<%=i%>" value="<%=eval("student_requirement_id_"&i)%>">
<b>Requirement <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Type:&nbsp;</font></td>
<td align=left><select name="requirement_id_<%=i%>" style="width:150px" onchange="DoChange(<%=i%>);"><option value="">(delete)</option>
<%
for j = 0 to requirements
  isselected = ""
  if toNum(requirement_ids(j)) = toNum(eval("requirement_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & requirement_ids(j) & """" & isselected & ">" & requirement_names(j) & "</option>")
next
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Completed:&nbsp;</font></td>
<td align=left><font face='arial,helvetica' size=-1><input type=checkbox name="completed_<%=i%>" value="Y" <% if eval("completed_"&i) = "Y" then %>checked<% end if %> onclick="CheckCompleted(<%=i%>);" onchange="DoChange(<%=i%>);"></font></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Completed&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="completed_date_<%=i%>" value="<%=eval("completed_date_"&i)%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);ChngCompleted(<%=i%>);DoChange(<%=i%>);"><%DrawCal "dataform.completed_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Expires&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="expires_date_<%=i%>" value="<%=eval("expires_date_"&i)%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.expires_date_"&i, "DoChange("&i&")"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:440px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Status:&nbsp;</font></td>
<td align=left><input type=text name="status_<%=i%>" value="<%=eval("status_"&i)%>" size=15 maxlength=50 style="width:100px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
</td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=10><br><% end if %>
<% next %>





<% case "exams" %>

<%
Dim exam_part()
Dim exam_parts()
sql = "select * from exams"
if session("zone") <> "" then sql = sql & " where zone_name='' OR zone_name = " & checkstring(session("zone"),50)
sql = sql & " order by exam_name"
rs.open sql,conn,1,1
ReDim exam_part(rs.RecordCount,10)
ReDim exam_parts(rs.RecordCount)
i = 0
while not rs.eof
  exam_id_list = exam_id_list & rs("exam_id") & "|"
  exam_name_list = exam_name_list & rs("exam_name") & "|"
  exam_parts(i) = -1
  for j = 1 to 10
    n = j
    if n < 10 then n = "0" & n
    x = rs("score" & n & "_desc")
    if isnull(x) then x = ""
    exam_part(i,j-1) = x
    if x <> "" then exam_parts(i) = j-1
  next
  i = i + 1
  rs.movenext
wend
rs.close
exam_ids = split(exam_id_list,"|")
exam_names = split(exam_name_list,"|")
exams = ubound(exam_ids)-1
%>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no exams listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<% if i > 1 then %><hr noshade size=1><% end if %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="student_exam_id_<%=i%>" value="<%=eval("student_exam_id_"&i)%>">
<b>Exam <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Exam:&nbsp;</font></td>
<td align=left><select name="exam_id_<%=i%>" style="width:140px" onchange="DoChange(<%=i%>);parent.DoRefresh();"><option value="">(delete)</option>
<%
theExamID = -1
for j = 0 to exams
  isselected = ""
  if toNum(exam_ids(j)) = toNum(eval("exam_id_"&i)) then
    isselected = " selected"
    theExamID = j
  end if
  response.write("<option value=""" & exam_ids(j) & """" & isselected & ">" & exam_names(j) & "</option>")
next
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Score:&nbsp;</font></td>
<td align=left><input type=text name="score_<%=i%>" value="<%=eval("score_"&i)%>" size=15 maxlength=50 style="width:95px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Pass/Fail:&nbsp;</font></td>
<td align=left><select name="pass_fail_<%=i%>" style="width:95px" onchange="DoChange(<%=i%>);">
<option value="" <% if eval("pass_fail_"&i) = "" then %>selected<% end if %>></option>
<option value="P" <% if eval("pass_fail_"&i) = "P" then %>selected<% end if %>>P</option>
<option value="F" <% if eval("pass_fail_"&i) = "F" then %>selected<% end if %>>F</option>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Ref&nbsp;Num:&nbsp;</font></td>
<td align=left><input type=text name="reference_number_<%=i%>" value="<%=eval("reference_number_"&i)%>" size=15 maxlength=50 style="width:95px" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right colspan=3><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Application&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="application_date_<%=i%>" value="<%=eval("application_date_"&i)%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.application_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Exam&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="exam_date_<%=i%>" value="<%=eval("exam_date_"&i)%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.exam_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Scores&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="scores_recorded_date_<%=i%>" value="<%=eval("scores_recorded_date_"&i)%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.scores_recorded_date_"&i, "DoChange("&i&")"%></td>
</tr>
</table>
<%
if theExamID >= 0 then
  if exam_parts(theExamID) >= 0 then
    response.write("<img src=""images/spacer.gif"" border=0 width=1 height=2><br>")
    response.write("<table border=0 cellspacing=0 cellpadding=0><tr>")
    response.write("<td align=center valign=top><font face='arial,helvetica' size=-1>Scores:&nbsp;</font></td>")
    response.write("<td align=center valign=top bgcolor=AAAAAA><table border=0 cellspacing=1 cellpadding=2><tr>")
    for j = 0 to exam_parts(theExamID)
      n = j+1
      if n < 10 then n = "0" & n
      response.write("<td align=center valign=bottom bgcolor=EEEEEE><font face='arial,helvetica' size=-2>&nbsp;" & exam_part(theExamID,j) & "&nbsp;<br>")
      response.write("<input type=text name=""score" & n & "_" & i & """ value=""" & (eval("score"&n&"_"&i)) & """ size=15 style=""width:50px"" onchange=""DoChange("&i&");"">")
      response.write("</font></td>")
    next
    response.write("</tr></table></td>")
    response.write("</tr></table>")
  end if
end if
%>
</td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=10><br><% end if %>
<% next %>





<% case "other" %>

<img src="images/spacer.gif" border=0 width=600 height=10><br>
<font size=-1>
<input type=hidden name="hasChanged_1" value="<%=hasChanged_1%>">
<input type=hidden name="person_id_1" value="<%=person_id_1%>">
<b>General</b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td rowspan=5>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Harvard&nbsp;ID:&nbsp;</font></td>
<td align=left><input type=text name="harvard_id_1" id="harvard_id_1" value="<%=harvard_id_1%>" size=15 maxlength=8 style="width:95px" onChange="validateHrvdId(this, false);DoChange(1);"></td>
<td rowspan=3>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Marital:&nbsp;</font></td>
<td align=left><select name="marital_1" style="width:150px" onchange="DoChange(1);">
<option value="" <% if marital_1 = "" then %>selected<% end if %>></option>
<option value="Divorced" <% if marital_1 = "Divorced" then %>selected<% end if %>>Divorced</option>
<option value="Married" <% if marital_1 = "Married" then %>selected<% end if %>>Married</option>
<option value="Separated" <% if marital_1 = "Separated" then %>selected<% end if %>>Separated</option>
<option value="Single" <% if marital_1 = "Single" then %>selected<% end if %>>Single</option>
<option value="Widowed" <% if marital_1 = "Widowed" then %>selected<% end if %>>Widowed</option>
</select>
</td>
<td rowspan=3>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Gender:&nbsp;</font></td>
<td align=left><select name="gender_1" style="width:100px" onchange="DoChange(1);">
<option value="" <% if gender_1 = "" then %>selected<% end if %>></option>
<option value="F" <% if gender_1 = "F" then %>selected<% end if %>>Female</option>
<option value="M" <% if gender_1 = "M" then %>selected<% end if %>>Male</option>
<option value="U" <% if gender_1 = "U" then %>selected<% end if %>>Unknown</option>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>SSN:&nbsp;</font></td>
<td align=left><input type=text name="ssn_1" id="ssn_1" value="<%=ssn_1%>" size=15 maxlength=9 style="width:95px" onchange="validateSSN(this, false);DoChange(1);"></td>
<td align=right><font face='arial,helvetica' size=-1>Birthplace:&nbsp;</font></td>
<td align=left><input type=text name="birthplace_1" value="<%=birthplace_1%>" size=30 maxlength=30 style="width:150px" onchange="DoChange(1);"></td>
<td align=right><font face='arial,helvetica' size=-1>Deceased:&nbsp;</font></td>
<td align=left><font face='arial,helvetica' size=-1><input type=checkbox name="deceased_1" value="Y" <% if deceased_1 = "Y" then %>checked<% end if %> onchange="DoChange(1);"></font></td>
</tr>
<tr>
<td align="right"><font face='arial,helvetica' size="-1">Birthday:&nbsp;</font></td>
<td align="left"><input type="text" name="date_of_birth_1" value="<%=date_of_birth_1%>" size="15" maxlength="50" style="width:75px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.date_of_birth_1", "DoChange(1)"%></td>
<% if session("zone") = "Dental" then %>
	<td align="right"><font face='arial,helvetica' size="-1">DentPin:&nbsp;</font></td>
	<td align="left" colspan="3"><input type="text" id="dentpin_1"  name="dentpin_1" value="<%=dentpin_1%>" size="15" maxlength="8" style="width:75px" onchange="DoChange(1);" /></td>
<% elseif session("zone") = "Medical" then %>
	<td align="right"><font face='arial,helvetica' size="-1">AAMC Id:&nbsp;</font></td>
	<td align="left" colspan="3"><input type="text" id="aamcId"  name="aamcId" value="<%=aamcId%>" size="15" maxlength="8" style="width:75px" onchange="DoChange(1);" /></td>
<% else %>
<% end if %>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Email:&nbsp;</font></td>
<td align=left colspan=7><input type=text name="email_address_1" value="<%=email_address_1%>" size=50 maxlength=255 style="width:500px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Alt. Email:&nbsp;</font></td>
<td align=left colspan=7><input type=text name="alt_email_address_1" value="<%=alt_email_address_1%>" size=50 maxlength=255 style="width:500px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td colspan=9><img src="images/spacer.gif" border=0 width=1 height=10></td>
</tr>
<tr>
<td colspan=9><font face='arial,helvetica' size=-1><b>Archive</b></font></td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Barcode:&nbsp;</font></td>
<td align=left colspan=7><input type=text name="archive_barcode_1" value="<%=archive_barcode_1%>" size=150 maxlength=150 style="width:500px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Box:&nbsp;</font></td>
<td align=left colspan=7><input type=text name="archive_box_1" value="<%=archive_box_1%>" size=150 maxlength=150 style="width:500px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Years:&nbsp;</font></td>
<td align=left colspan=7><input type=text name="archive_years_1" value="<%=archive_years_1%>" size=150 maxlength=150 style="width:500px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td colspan=9><img src="images/spacer.gif" border=0 width=1 height=10></td>
</tr>
<tr>
<td colspan=9><font face='arial,helvetica' size=-1><b>Notes</b></font></td>
</tr>
<tr>
<td align=right>&nbsp;&nbsp;</td>
<td align=left colspan=8><textarea rows=3 cols=50 name="notes_1" style="width:600px" onchange="DoChange(1);"><%=notes_1%></textarea></td>
</tr>
</table>







<% case "ferpa" %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>
<input type=hidden name="hasChanged_1" value="<%=hasChanged_1%>">
<input type=hidden name="person_id_1" value="<%=person_id_1%>">
<font size=-1>
<b>Name</b><br>
&nbsp;&nbsp;Restricted:&nbsp;<input type=checkbox name="name_restriction_1" value="Y" <% if name_restriction_1 = "Y" then %>checked<% end if %> onchange="DoChange(1);">Yes<br>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<b>Address</b><br>
&nbsp;&nbsp;Restricted:&nbsp;<input type=checkbox name="address_restriction_1" value="Y" <% if address_restriction_1 = "Y" then %>checked<% end if %> onchange="DoChange(1);">Yes<br>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<b>Photo</b><br>
&nbsp;&nbsp;Restricted:&nbsp;<input type=checkbox name="photo_restriction_1" value="Y" <% if photo_restriction_1 = "Y" then %>checked<% end if %> onchange="DoChange(1);">Yes<br>




<% case "instances" %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no instances listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% else %>
<center>
<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table border=0 cellspacing=1 cellpadding=2>
<tr bgcolor=F7F7F7>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;Program&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;Track&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;Zone&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;Status&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;Matric&nbsp;Dt&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;Grad&nbsp;Dt&nbsp;</font></td>
</tr>
<%
for i = 1 to rows
  sid = eval("student_instance_id_"&i)
  sql = "select i.student_instance_id, g.program_name, t.track_name, g.zone_name, s.reg_status_name, convert(varchar(50),i.matric_date,1) matric_date, convert(varchar(50),i.actual_grad_date,1) actual_grad_date "
  sql = sql & " from student_instance i, programs g, program_track t, reg_status s "
  sql = sql & " where i.program_id = g.program_id and i.program_track_id = t.program_track_id and i.reg_status_id = s.reg_status_id "
  sql = sql & " and i.student_instance_id = " & checkstring(sid,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    if session("zone") = "" or session("zone") = rs("zone_name") then
      response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFFCC';"" onMouseOut=""this.bgColor='#FFFFFF';"" onClick=""parent.document.location='p_summary.asp?id=" & sid & "';"">")
    else
      response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFDDDD';"" onMouseOut=""this.bgColor='#FFFFFF';"" onClick=""alert('This instance does not belong to your selected zone.');"">")
    end if
    theStar = ""
    if cstr(rs("student_instance_id")) = cstr(id) then theStar = "&nbsp;*"
    response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("program_name") & theStar & "&nbsp;</font></td>")
    response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("track_name") & theStar & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("zone_name") & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("reg_status_name") & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("matric_date") & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("actual_grad_date") & "&nbsp;</font></td>")
    response.write("</tr>")
  end if
  rs.close
next
%>
</table>
</td></tr></table>
</center>
<% end if %>

<%
sql = " select g.degree_name, ds.degree_status_name, t.institution_name, d.start_date, d.end_date, d.completed_date, d.completed_month, d.completed_year,"
sql = sql & " d.student_instance_id, i.matric_date, i.actual_grad_date, i.student_instance_id, p.zone_name "
sql = sql & " from student_degrees d left outer join degree_status ds on d.degree_status_id = ds.degree_status_id"
sql = sql & " left outer join institution t on d.institution_id = t.institution_id"
sql = sql & " left outer join student_instance i on d.student_instance_id = i.student_instance_id "
sql = sql & " left outer join programs p on i.program_id = p.program_id, degrees g "
sql = sql & " where d.person_id = " & checkstring(person_id,50) & " and g.degree_id = d.degree_id"
sql = sql & " order by (case when d.student_instance_id = i.student_instance_id then 0 else 1 end), "
sql = sql & " (case when (select x.degree_name from degrees x where x.degree_id = d.degree_id) = 'Non-degree' then -1000 else d.student_degree_id end) desc "
rs.open sql,conn,1,1
if not rs.eof then
  response.write("<center>")
  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")
  response.write("<table border=0 cellspacing=0 cellpadding=0><tr>")
  response.write("<td bgcolor=999999>")
  response.write("<table border=0 cellspacing=1 cellpadding=1>")
  response.write("<tr bgcolor=F7F7F7>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Degree&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Status&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Institution&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Start&nbsp;Dt&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;End&nbsp;Dt&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Completed&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Matric&nbsp;Dt&nbsp;</font></td>")
  response.write("<td align=center valign=middle><font face='arial,helvetica' size=-1>&nbsp;Grad&nbsp;Dt&nbsp;</font></td>")
  response.write("</tr>")
  while not rs.eof
    temp = rs("completed_month")
    if isnull(temp) then temp = ""
    if temp = "" then temp = "?"
    month_year = temp & "/"
    temp = rs("completed_year")
    if isnull(temp) then temp = ""
    if temp = "" then temp = "?"
    month_year = month_year & temp
    if month_year = "?/?" then month_year = "&nbsp;"
    if rs("completed_date") <> "" then month_year = FormatSmallDate(rs("completed_date"))

    if isnull(rs("student_instance_id")) then
      response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFDDDD';"" onMouseOut=""this.bgColor='#FFFFFF';"" onClick=""alert('This degree is not associated with a student instance.');"">")
    elseif session("zone") = "" or session("zone") = rs("zone_name") then
      response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFFCC';"" onMouseOut=""this.bgColor='#FFFFFF';"" onClick=""parent.document.location='p_summary.asp?id=" & rs("student_instance_id") & "';"">")
    else
      response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFDDDD';"" onMouseOut=""this.bgColor='#FFFFFF';"" onClick=""alert('This instance does not belong to your selected zone.');"">")
    end if
    if toStr(rs("student_instance_id")) = toStr(id) then
      response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;*&nbsp;" & rs("degree_name") & "&nbsp;*&nbsp;</font></td>")
    else
      response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("degree_name") & "&nbsp;</font></td>")
    end if
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("degree_status_name") & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("institution_name") & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & FormatSmallDate(rs("start_date")) & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & FormatSmallDate(rs("end_date")) & "&nbsp;</font></td>")
    response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & month_year & "&nbsp;</font></td>")
    if toStr(rs("student_instance_id")) = toStr(id) then
      response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & FormatSmallDate(rs("matric_date")) & "&nbsp;</font></td>")
      response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & FormatSmallDate(rs("actual_grad_date")) & "&nbsp;</font></td>")
    else
      response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;</font></td>")
      response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;</font></td>")
    end if
    response.write("</tr>")
    rs.movenext
  wend
  response.write("</table>")
  response.write("</td>")
  response.write("</tr></table>")
  response.write("</center>")
end if
rs.close
%>
<center>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font face='arial,helvetica' size=1>* Indicates currently selected student instance or degree.</font><br>
</center>


<% case "instance" %>

<script>
function PickInstitution(x,y) {
  eval("document.dataform."+x+".value = '';");
  eval("document.dataform."+y+".value = '';");
  window.open("g_institution.asp?fi="+x+"&fn="+y,"InstitutionWindow","width=400,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
function PickCountry(x) {
  eval("document.dataform."+x+".value = '';");
  window.open("g_country.asp?fn="+x,"CountryWindow","width=400,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<input type=hidden name="hasChanged_1" value="<%=hasChanged_1%>">
<input type=hidden name="student_instance_id_1" value="<%=student_instance_id_1%>">
<input type=hidden name="person_id_1" value="<%=person_id_1%>">
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td valign=top align=left>

<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=5><font face='arial,helvetica' size=-1><b>General&nbsp;Info</b></font></td>
</tr>
<tr>
<td rowspan=5>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Program:&nbsp;</font></td>
<td align=left colspan=3><select name="program_id_1" style="width:325px" onchange="DoChange(1);parent.DoRefresh(<%=i%>);"><option value=""></option>
<%
if session("zone") = "" then
  sql = "select program_id, program_name from programs order by program_name"
else
  sql = "select program_id, program_name from programs where zone_name = " & checkstring(session("zone"),50) & " order by program_name"
end if
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if toNum(rs("program_id")) = toNum(program_id_1) then isselected = " selected"
  response.write("<option value=""" & rs("program_id") & """" & isselected & ">" & rs("program_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Track:&nbsp;</font></td>
<td align=left colspan=3><select name="program_track_id_1" style="width:325px" onchange="DoChange(1);parent.DoRefresh(<%=i%>);"><option value=""></option>
<%
sql = "select program_track_id, track_name from program_track where program_id = " & checkstring(program_id_1,50) & " order by track_name"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if toNum(rs("program_track_id")) = toNum(program_track_id_1) then
    isselected = " selected"
    track_name = rs("track_name")
  end if
  response.write("<option value=""" & rs("program_track_id") & """" & isselected & ">" & rs("track_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Status:&nbsp;</font></td>
<td align=left colspan=3><select name="reg_status_id_1" style="width:325px" onchange="DoChange(1);document.dataform.reg_status_date_1.value='<%=date%>';"><option value=""></option>
<%
sql = "select reg_status_id, reg_status_name from reg_status order by reg_status_name"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if toNum(rs("reg_status_id")) = toNum(reg_status_id_1) then isselected = " selected"
  response.write("<option value=""" & rs("reg_status_id") & """" & isselected & ">" & rs("reg_status_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
</table>
</td>
<td><img src="images/spacer.gif" border=0 width=15 height=1></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=5><font face='arial,helvetica' size=-1><b>&nbsp;</b></font></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Society:&nbsp;</font></td>
<td align=left><select name="society_1" style="width:120px" onchange="DoChange(1);"><option value=""></option>
<%
tempZone = ""
sql = "select zone_name from programs where program_id = " & checkstring(program_id_1,50)
rs.open sql,conn,1,1
if not rs.eof then
  tempZone = rs("zone_name")
end if
rs.close
sql = "select code from lookups where lookup_type = 'society' and zone_name = " & checkstring(tempZone,50) & " order by code"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if rs("code") = society_1 then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Time Status:&nbsp;</font></td>
<td align=left><select name="fulltime_1" style="width:120px" onchange="DoChange(1);">
<option value="" <% if fulltime_1 = "" then %>selected<% end if %>></option>
<option value="F" <% if fulltime_1 = "F" then %>selected<% end if %>>Full Time</option>
<option value="H" <% if fulltime_1 = "H" then %>selected<% end if %>>Half Time</option>
<option value="L" <% if fulltime_1 = "L" then %>selected<% end if %>>Less Than Half</option>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Status&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="reg_status_date_1" value="<%=reg_status_date_1%>" size=15 style="width:85px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.reg_status_date_1", "DoChange(1)"%></td>
</tr>
</table>
</td>
</tr>
</table>
<% if track_name = "Exchange Clerk" then %>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=5><font face='arial,helvetica' size=-1><b>Exchange&nbsp;Clerk</b></font></td>
</tr>
<%
if xclerk_school_id_1 <> "" and xclerk_school_1 = "" then
  sql = "select institution_name from institution where institution_id = " & checkstring(xclerk_school_id_1,50)
  rs.open sql,conn,1,1
  if not rs.eof then xclerk_school_1 = rs("institution_name")
  rs.close
end if
%>
<tr>
<td rowspan=5>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>School:&nbsp;</font></td>
<td align=left><input type=hidden name="xclerk_school_id_1" value="<%=xclerk_school_id_1%>"><input type=text ContentEditable=false name="xclerk_school_1" value="<%=xclerk_school_1%>" size=15 style="width:215px" onchange="DoChange(1);"><a href="JavaScript:PickInstitution('xclerk_school_id_1','xclerk_school_1');DoChange(1);"><img src="images/search.gif" border=0></a></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;App&nbsp;Sent&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="date_application_package_sent_1" value="<%=date_application_package_sent_1%>" size=15 style="width:100px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.date_application_package_sent_1", "DoChange(1)"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Country:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="xclerk_school_country_1" value="<%=xclerk_school_country_1%>" size=15 style="width:215px" onchange="DoChange(1);"><a href="JavaScript:PickCountry('xclerk_school_country_1');DoChange(1);"><img src="images/search.gif" border=0></a></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Ineligible&nbsp;Reason:&nbsp;</font></td>
<td align=left><input type=text name="ineligible_reason_1" value="<%=ineligible_reason_1%>" size=15 style="width:150px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Program:&nbsp;</font></td>
<td align=left><select name="xclerk_program_1" style="width:215px" onchange="DoChange(1);"><option value=""></option>
<%
sql = "SELECT code FROM lookups WHERE lookup_type = 'exclerk_special_program' AND is_active='Y' AND (zone_name='" & session("zone") & "' OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if rs("code") = xclerk_program_1 then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>Exp&nbsp;Grad&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="expected_grad_date_1" value="<%=expected_grad_date_1%>" size=15 maxlength=50 style="width:100px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.expected_grad_date_1", "DoChange(1)"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Billing&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="xclerk_billing_notes_1" value="<%=xclerk_billing_notes_1%>" size=15 style="width:505px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Sched&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="xclerk_scheduling_notes_1" value="<%=xclerk_scheduling_notes_1%>" size=15 style="width:505px" onchange="DoChange(1);"></td>
</tr>
</table>
<% else %>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=6><font face='arial,helvetica' size=-1><b>Other&nbsp;Info</b></font></td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
<td valign=top align=left>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Matric&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="matric_date_1" value="<%=matric_date_1%>" size=15 maxlength=50 style="width:75px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.matric_date_1", "DoChange(1)"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Orig&nbsp;Class:&nbsp;</font></td>
<td align=left><select name="original_class_1" style="width:95px" onchange="DoChange(1);"><option value=""></option>
<%
sql = "select code from lookups where lookup_type = 'original_class' AND is_active='Y' AND (zone_name='" & session("zone") & "' OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if cstr(rs("code")) = cstr(original_class_1) then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Current&nbsp;Class:&nbsp;</font></td>
<td align=left><select name="current_class_1" style="width:95px" onchange="DoChange(1);"><option value=""></option>
<%
for i = 2030 to 1960 step -1
  isselected = ""
  if i = toNum(current_class_1) then isselected = " selected"
  response.write("<option value=""" & i & """" & isselected & ">" & i & "</option>")
next
%>
</select>
</td>
</tr>
</table>
</td>
<td><img src="images/spacer.gif" border=0 width=20 height=1></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Year&nbsp;of&nbsp;Study:&nbsp;</font></td>
<td align=left><select name="year_of_study_id_1" style="width:50px" onchange="DoChange(1);"><option value=""></option>
<%
sql = "select year_of_study_id, year_of_study_name from year_of_study order by year_of_study_name"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if rs("year_of_study_id") = toNum(year_of_study_id_1) then isselected = " selected"
  response.write("<option value=""" & rs("year_of_study_id") & """" & isselected & ">" & rs("year_of_study_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Tuition&nbsp;Count:&nbsp;</font></td>
<td align=left><input type=text name="tuition_count_1" value="<%=tuition_count_1%>" size=15 maxlength=10 style="width:50px" onchange="IsNumeric(this, false);DoChange(1);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Years&nbsp;in&nbsp;Prgm:&nbsp;</font></td>
<td align=left><input type=text name="years_in_program_1" value="<%=years_in_program_1%>" size=15 maxlength=10 style="width:50px" onchange="IsNumeric(this, false);DoChange(1);"></td>
</tr>
</table>
</td>
<td><img src="images/spacer.gif" border=0 width=20 height=1></td>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Anticipated&nbsp;YoS:&nbsp;</font></td>

<td align=left><select name="anticipated_year_of_study_1" style="width:95px" onchange="DoChange(1);"><option value=""></option>
<%
sql = "select code from lookups where lookup_type = 'anticipated_year_of_study' AND is_active='Y' AND (zone_name='" & session("zone") & "' OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if cstr(rs("code")) = cstr(anticipated_year_of_study_1) then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
  </select>
  </td>

<!--<td align=left><input type=text name="anticipated_year_of_study_1" value="<%=anticipated_year_of_study_1%>" size=15 maxlength=150 style="width:120px" onchange="DoChange(1);"></td>-->
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Anticipate&nbsp;5&nbsp;Yr:&nbsp;</font></td>
<td align=left><input type=text name="anticipate_5_year_program_1" value="<%=anticipate_5_year_program_1%>" size=15 maxlength=150 style="width:120px" onchange="DoChange(1);"></td>
</tr>

<tr>
<td align=right><font face='arial,helvetica' size=-1>PCE&nbsp;Assignment:&nbsp;</font></td>
<td align=left><select name="pce_assignment_1" style="width:95px" onchange="DoChange(1);"><option value=""></option>
<%
sql = "select code from lookups where lookup_type = 'PCE_assignment' AND is_active='Y' AND (zone_name='" & session("zone") & "' OR zone_name='') order by sort_order"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if cstr(rs("code")) = cstr(pce_assignment_1) then isselected = " selected"
  response.write("<option value=""" & rs("code") & """" & isselected & ">" & rs("code") & "</option>")
  rs.movenext
wend
rs.close
%>
  </select>
  </td>
  </tr>
  </table>
</td>
</tr>



</table>
<%
if cross_reg_school_id_1 <> "" and cross_reg_school_1 = "" then
  sql = "select institution_name from institution where institution_id = " & checkstring(cross_reg_school_id_1,50)
  rs.open sql,conn,1,1
  if not rs.eof then cross_reg_school_1 = rs("institution_name")
  rs.close
end if
%>
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;&nbsp;</td>
    <td align=right><font face='arial,helvetica' size=-1>Cross&nbsp;Reg&nbsp;School:&nbsp;</font></td>
    <td align=left><input type=hidden name="cross_reg_school_id_1" value="<%=cross_reg_school_id_1%>">
    <input type="text" ContentEditable="false" name="cross_reg_school_1" value="<%=cross_reg_school_1%>" size="15" maxlength="150" style="width:252px" onchange="DoChange(1);"><a href="JavaScript:PickInstitution('cross_reg_school_id_1','cross_reg_school_1');DoChange(1);"><img src="images/search.gif" border=0></a>
    &nbsp;&nbsp;&nbsp;&nbsp;</td>

  <td align=right><font face='arial,helvetica' size=-1>PCE&nbsp;Year:&nbsp;</font></td>
<td align=left><select name="pce_assignment_year_id_1" style="width:95px" onchange="DoChange(1);"><option value=""></option>
<%
sql = "select acad_year_id, acad_year_name from acad_year ORDER by acad_year_id desc"
rs.open sql,conn,1,1
while not rs.eof
  isselected = ""
  if cstr(rs("acad_year_id")) = cstr(pce_assignment_year_id_1) then isselected = " selected"
  response.write("<option value=""" & rs("acad_year_id") & """" & isselected & ">" & rs("acad_year_name") & "</option>")
  rs.movenext
wend
rs.close
%>
  </select>
  </td>


  </tr>
  <!-- Adding for OASIS CPP Flag sjt5 2/19/2016 -->
	<tr>
		<td>&nbsp;&nbsp;</td>
		<td align=right><font face='arial,helvetica' size=-1>In&nbsp;OASIS:&nbsp;</font></td>
<%
		in_oasis = ""
		sql = "select in_oasis from mycourses.dbo.person P where person_id  IN (SELECT person_id from madris.dbo.student_instance where student_instance_id =  "  & checkstring(id,50) & ")"
		rs.open sql,conn,1,1
		if not rs.eof then in_oasis = rs("in_oasis")
		rs.close
	%>
		<td align=left><font face='arial,helvetica' size=-1><%=in_oasis%></font>&nbsp;&nbsp;&nbsp;<font face='arial,helvetica' size=-1>Exclude&nbsp;from&nbsp;OASIS&nbsp;CPP feed:</font><input type=checkbox name="exclude_from_oasis_CPP_file_1" value="Y" <% if in_oasis <> "Y" then%> disabled <%end if %> <% if exclude_from_oasis_CPP_file_1 = "Y" then %>checked<% end if %> size=15 onchange="DoChange(1);"></td>

	</tr>



</table>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=5><font face='arial,helvetica' size=-1><b>Graduation</b></font></td>
</tr>
<tr>
<td rowspan=4>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Exp&nbsp;Grad&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="expected_grad_date_1" value="<%=expected_grad_date_1%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.expected_grad_date_1", "DoChange(1)"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Graduated:&nbsp;</font></td>
<td align=left><font face='arial,helvetica' size=-1><input type=checkbox name="graduated_1" value="Y" <% if graduated_1 = "Y" then %>checked<% end if %> size=15 onchange="DoChange(1);"></font></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Actual&nbsp;Grad&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="actual_grad_date_1" value="<%=actual_grad_date_1%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(1);"><%DrawCal "dataform.actual_grad_date_1", "DoChange(1)"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Graduation&nbsp;Cleared:&nbsp;</font></td>
<td align=left><font face='arial,helvetica' size=-1><input type=checkbox name="graduation_cleared_1" value="Y" <% if graduation_cleared_1 = "Y" then %>checked<% end if %> size=15 onchange="DoChange(1);"></font></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Diploma&nbsp;Name:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="diploma_name_1" value="<%=diploma_name_1%>" size=15 maxlength=255 style="width:495px" onchange="DoChange(1);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Residency&nbsp;Info:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="residency_information_1" value="<%=residency_information_1%>" size=15 maxlength=150 style="width:495px" onchange="DoChange(1);"></td>
</tr>
</table>
<% end if %>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=2><font face='arial,helvetica' size=-1><b>Notes</b></font></td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
<td align=left><textarea name="notes_1" rows=4 cols=50 style="width:590px" onchange="DoChange(1);"><%=notes_1%></textarea></td>
</tr>
</table>




<% case "advisors" %>

<%
sql = "SELECT * FROM advisor_roles order by role_name"
rs.open sql,conn,1,1
while not rs.eof
  role_id_list = role_id_list & rs("role_id") & "|"
  role_name_list = role_name_list & rs("role_name") & "|"
  rs.movenext
wend
rs.close
role_ids = split(role_id_list,"|")
role_names = split(role_name_list,"|")
roles = ubound(role_ids)-1
sql = "select project_id, coalesce(title,'Untitled Project') as title from student_projects where student_instance_id = " & id & " order by title"
rs.open sql,conn,1,1
while not rs.eof
  project_id_list = project_id_list & rs("project_id") & "|"
  project_name_list = project_name_list & rs("title") & "|"
  rs.movenext
wend
rs.close
project_ids = split(project_id_list,"|")
project_names = split(project_name_list,"|")
projects = ubound(project_ids)-1
%>
<table border=0 cellspacing=0 cellpadding=0 width=100%><tr><td align=right>
<table width=100 border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="JavaScript:parent.GoPage('p_details.asp?id=<%=id%>&pg=projects');" id="alink2">Projects</a></b></font></td></tr></table></td></tr></table>
</td></tr></table>
<script>
function PickPerson(x) {
  eval("document.dataform.advisor_username_"+x+".value='';");
  eval("document.dataform.advisor_name_"+x+".value='';");
  eval("document.dataform.advisor_phone_"+x+".value='';");
  eval("document.dataform.advisor_email_"+x+".value='';");
  window.open("g_person.asp?fi=advisor_username_"+x+"&fn=advisor_name_"+x+"&fp=advisor_phone_"+x+"&fe=advisor_email_"+x,"PersonWindow","width=550,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<font size=-1>
<% if rows = 0 then %>
<center><i>no advisors listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="advisor_id_<%=i%>" value="<%=eval("advisor_id_"&i)%>">
<b>Advisor <%=i%></b><br>
<%
faculty_name = ""
faculty_phone = ""
faculty_email = ""
if eval("advisor_username_"&i) <> "" then
  sql = "select last_first, telephone, email from faculty where username = " & checkstring(eval("advisor_username_"&i),50)
  rs.open sql,conn,1,1
  if not rs.eof then
    faculty_name = rs("last_first")
    faculty_phone = rs("telephone")
    faculty_email = rs("email")
  end if
  rs.close
end if
%>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Role:&nbsp;</font></td>
<td align=left><select name="role_id_<%=i%>" style="width:200px" onchange="DoChange(<%=i%>);"><option value="">(delete)</option>
<%
for j = 0 to roles
  isselected = ""
  if toNum(role_ids(j)) = toNum(eval("role_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & role_ids(j) & """" & isselected & ">" & role_names(j) & "</option>")
next
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Start&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="start_date_<%=i%>" value="<%=eval("start_date_"&i)%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.start_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;End&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="end_date_<%=i%>" value="<%=eval("end_date_"&i)%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.end_date_"&i, "DoChange("&i&")"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Advisor:&nbsp;</font></td>
<td align=left colspan=5>
<input type=hidden name="advisor_username_<%=i%>" value="<%=eval("advisor_username_"&i)%>">
<input type=text name="advisor_name_<%=i%>" value="<%=faculty_name%>" size=20 style="width:175px" onchange="DoChange(<%=i%>);" contenteditable=false>
<input type=text name="advisor_phone_<%=i%>" value="<%=faculty_phone%>" size=20 style="width:100px" onchange="DoChange(<%=i%>);" contenteditable=false>
<input type=text name="advisor_email_<%=i%>" value="<%=faculty_email%>" size=20 style="width:215px" onchange="DoChange(<%=i%>);" contenteditable=false>
<a href="JavaScript:PickPerson('<%=i%>');DoChange(<%=i%>);"><img src="images/search.gif" border=0></a>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Project:&nbsp;</font></td>
<td align=left colspan=5><select name="project_id_<%=i%>" style="width:525px" onchange="DoChange(<%=i%>);"><option value=""></option>
<%
for j = 0 to projects
  isselected = ""
  if toNum(project_ids(j)) = toNum(eval("project_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & project_ids(j) & """" & isselected & ">" & project_names(j) & "</option>")
next
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:525px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=5><br><% end if %>
<% next %>




<% case "projects" %>

<table border=0 cellspacing=0 cellpadding=0 width=100%><tr><td align=right>
<table width=100 border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="JavaScript:parent.GoPage('p_details.asp?id=<%=id%>&pg=advisors');" id="alink2">Advisors</a></b></font></td></tr></table></td></tr></table>
</td></tr></table>
<script>
function PickInstitution(x) {
  eval("document.dataform.institution_id_"+x+".value = '';");
  eval("document.dataform.institution_name_"+x+".value = '';");
  window.open("g_institution.asp?fi=institution_id_"+x+"&fn=institution_name_"+x,"InstitutionWindow","width=400,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<font size=-1>
<% if rows = 0 then %>
<center><i>no projects listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="project_id_<%=i%>" value="<%=eval("project_id_"&i)%>">
<%
institution_name = ""
if eval("institution_id_"&i) <> "" then
  sql = "select institution_name from institution where institution_id = " & checkstring(eval("institution_id_"&i),50)
  rs.open sql,conn,1,1
  if not rs.eof then
    institution_name = rs("institution_name")
  end if
  rs.close
end if
%>
<b>Project <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Title:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="title_<%=i%>" value="<%=eval("title_"&i)%>" size=15 maxlength=255 style="width:500px" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Transcript:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="transcript_title_<%=i%>" value="<%=eval("transcript_title_"&i)%>" size=15 maxlength=255 style="width:500px" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Institution:&nbsp;</font></td>
<td align=left><input type=hidden name="institution_id_<%=i%>" value="<%=eval("institution_id_"&i)%>"><input type=text name="institution_name_<%=i%>" value="<%=institution_name%>" size=15 style="width:200px;font-color:#000000" ContentEditable=false onchange="DoChange(<%=i%>);" <%=disabled%>><a href="JavaScript:PickInstitution('<%=i%>');DoChange(<%=i%>);"><img src="images/search.gif" border=0></a></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Date:&nbsp;</font></td>
<td align=left><input type=text name="project_date_<%=i%>" value="<%=eval("project_date_"&i)%>" size=15 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.project_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Approved:&nbsp;</font></td>
<td align=left><nobr><font face='arial,helvetica' size=-1><input type=checkbox name="approved_<%=i%>" value="Y" <% if eval("approved_"&i) = "Y" then %>checked<% end if %> onchange="DoChange(<%=i%>);"></font></nobr></td>
</tr>
<tr>
<td align=right valign=top><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Description:&nbsp;</font></td>
<td align=left colspan=5><textarea name="description_<%=i%>" rows=3 cols=50 style="width:500px" onchange="DoChange(<%=i%>);"><%=eval("description_"&i)%></textarea></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:500px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=5><br><% end if %>
<% next %>





<% case "reghist" %>

<%
sql = "select * from terms order by start_date desc"
rs.open sql,conn,1,1
while not rs.eof
  term_id_list = term_id_list & rs("term_id") & "|"
  term_name_list = term_name_list & rs("term_name") & "|"
  rs.movenext
wend
rs.close
term_ids = split(term_id_list,"|")
term_names = split(term_name_list,"|")
terms = ubound(term_ids)-1
sql = "select * from reg_status order by reg_status_name"
rs.open sql,conn,1,1
while not rs.eof
  status_id_list = status_id_list & rs("reg_status_id") & "|"
  status_name_list = status_name_list & rs("reg_status_name") & "|"
  rs.movenext
wend
rs.close
status_ids = split(status_id_list,"|")
status_names = split(status_name_list,"|")
statuses = ubound(status_ids)-1
sql = "select * from year_of_study order by year_of_study_id"
rs.open sql,conn,1,1
while not rs.eof
  yos_id_list = yos_id_list & rs("year_of_study_id") & "|"
  yos_name_list = yos_name_list & rs("year_of_study_name") & "|"
  rs.movenext
wend
rs.close
yos_ids = split(yos_id_list,"|")
yos_names = split(yos_name_list,"|")
yoses = ubound(yos_ids)-1
%>
<script>
function ChangeTerm(which) {
  eval("x = document.dataform.term_id_"+which+";");
  eval("y = document.dataform.effective_start_date_"+which+";");
  eval("z = document.dataform.effective_end_date_"+which+";");
<%
sql = "select t.term_id, dbo.GetRegHistDate('S',t.term_id,null,i.program_id,i.year_of_study_id) start_date, "
sql = sql & " dbo.GetRegHistDate('E',t.term_id,null,i.program_id,i.year_of_study_id) end_date "
sql = sql & " from terms t, student_instance i where i.student_instance_id = " & checkstring(id,50) & " order by t.acad_year_id, t.term_name desc"
rs.open sql,conn,1,1
while not rs.eof
  response.write("if (x.value=='" & rs("term_id") & "') {y.value='" & rs("start_date") & "'; z.value='" & rs("end_date") & "';} ")
  rs.movenext
wend
rs.close
%>
}
</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no registration histories listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% else %>
<table border=0 cellspacing=2 cellpadding=0>
<tr>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>Term</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>Start&nbsp;Dt</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>End&nbsp;Dt</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>YoS</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>Status</font></td>
</tr>
<% for i = 1 to rows %>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="student_reg_hist_id_<%=i%>" value="<%=eval("student_reg_hist_id_"&i)%>">
<b>Reg Hist <%=i%></b>
</font></td>
<td align=center><select name="term_id_<%=i%>" style="width:115px" onchange="ChangeTerm(<%=i%>);DoChange(<%=i%>);"><option value=""></option>
<%
for j = 0 to terms
  isselected = ""
  if toNum(term_ids(j)) = toNum(eval("term_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & term_ids(j) & """" & isselected & ">" & term_names(j) & "</option>")
next
%>
</select>
</td>
<td align=center><input type=text name="effective_start_date_<%=i%>" value="<%=eval("effective_start_date_"&i)%>" size=15 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.effective_start_date_"&i, "DoChange("&i&")"%></td>
<td align=center><input type=text name="effective_end_date_<%=i%>" value="<%=eval("effective_end_date_"&i)%>" size=15 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.effective_end_date_"&i, "DoChange("&i&")"%></td>
<td align=center><select name="year_of_study_id_<%=i%>" style="width:50px" onchange="DoChange(<%=i%>);"><option value=""></option>
<%
for j = 0 to yoses
  isselected = ""
  if toNum(yos_ids(j)) = toNum(eval("year_of_study_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & yos_ids(j) & """" & isselected & ">" & yos_names(j) & "</option>")
next
%>
</select>
</td>
<td align=left><select name="reg_status_id_<%=i%>" style="width:150px" onchange="DoChange(<%=i%>);"><option value="">(delete)</option>
<%
for j = 0 to statuses
  isselected = ""
  if toNum(status_ids(j)) = toNum(eval("reg_status_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & status_ids(j) & """" & isselected & ">" & status_names(j) & "</option>")
next
%>
</select>
</td>
</tr>
<% next %>
</table>
<% end if %>






<% case "correspond" %>

<%
sql = "select g.zone_name from programs g, student_instance i where i.program_id = g.program_id and i.student_instance_id = " & checkstring(id,50)
zone_name = ""
rs.open sql,conn,1,1
if not rs.eof then
  zone_name = rs("zone_name")
end if
rs.close
sql = "select * from letter_defs WHERE (active='Y' OR letter_id IN (SELECT letter_id from correspondence where student_instance_id = " & checkstring(id,50) & "))"
if session("zone") <> "" then sql = sql & " AND (zone_name='' OR zone_name = " & checkstring(session("zone"),50) & ")"
sql = sql & " order by letter_type, letter_name"
rs.open sql,conn,1,1
while not rs.eof
  letter_id_list = letter_id_list & rs("letter_id") & "|"
  letter_name_list = letter_name_list & rs("letter_type") & " - " & rs("letter_name") & "|"
  rs.movenext
wend
rs.close
letter_ids = split(letter_id_list,"|")
letter_names = split(letter_name_list,"|")
letters = ubound(letter_ids)-1
%>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no correspondences on record</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% else %>
<table border=0 cellspacing=2 cellpadding=0>
<tr>
<td align=center><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>Letter</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>Status</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=center><font face='arial,helvetica' size=-1>Print&nbsp;Dt</font></td>
<td align=center rowspan=<%=rows+1%>><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
</tr>
<% for i = 1 to rows %>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="zone_name_<%=i%>" value="<%=zone_name%>">
<input type=hidden name="correspondence_id_<%=i%>" value="<%=eval("correspondence_id_"&i)%>">
<% if eval("correspondence_id_"&i) <> "" then %>
<b><a href="JavaScript:parent.GoPage('p_details.asp?pg=correspond2&id=<%=id%>&cid=<%=eval("correspondence_id_"&i)%>');" id="alink2">Correspondence <%=i%></a></b>
<% else %>
<b>Correspondence <%=i%></b>
<% end if %>
</font></td>
<td align=center><select name="letter_id_<%=i%>" style="width:250px" onchange="DoChange(<%=i%>);"><option value="">(delete)</option>
<%
for j = 0 to letters
  isselected = ""
  if toNum(letter_ids(j)) = toNum(eval("letter_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & letter_ids(j) & """" & isselected & ">" & letter_names(j) & "</option>")
next
%>
</select>
</td>
<td align=left><select name="status_<%=i%>"" style="width:100px" onchange="DoChange(<%=i%>);">
<option value="" <% if eval("status_"&i) = "" then %>selected<% end if %>></option>
<option value="Sent" <% if eval("status_"&i) = "Sent" then %>selected<% end if %>>Sent</option>
<option value="Queued" <% if eval("status_"&i) = "Queued" then %>selected<% end if %>>Queued</option>
</select>
</td>
<td align=center><input type=text name="printed_date_<%=i%>" value="<% if ((not isnull(eval("printed_date_"&i))) AND (eval("printed_date_"&i)<>"")) then response.write(FormatDateTime(eval("printed_date_"&i),2)) end if %>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.printed_date_"&i, "DoChange("&i&")"%></td>
</tr>
<% next %>
</table>
<% end if %>





<% case "correspond2" %>

<%
sql = "select g.zone_name from programs g, student_instance i where i.program_id = g.program_id and i.student_instance_id = " & checkstring(id,50)
zone_name = ""
rs.open sql,conn,1,1
if not rs.eof then
  zone_name = rs("zone_name")
end if
rs.close
sql = "select * from letter_defs WHERE active='Y'"
if session("zone") <> "" then sql = sql & " AND (zone_name='' OR zone_name = " & checkstring(session("zone"),50) & ")"
sql = sql & " order by letter_type, letter_name"
rs.open sql,conn,1,1
while not rs.eof
  letter_id_list = letter_id_list & rs("letter_id") & "|"
  letter_name_list = letter_name_list & rs("letter_type") & " - " & rs("letter_name") & "|"
  rs.movenext
wend
rs.close
letter_ids = split(letter_id_list,"|")
letter_names = split(letter_name_list,"|")
letters = ubound(letter_ids)-1
i = 1
if request("cid") = "" and letter_id_1 = "" then
  hasChanged_1 = 2
  person_id_1 = person_id
  student_instance_id_1 = id
  copies_1 = "1"
  requested_date_1 = date
  status_1 = "Queued"

  ' Po: 05/24/2017 - Add addition query to return 'Mailing' typed address.
  sql = "select person_address_id from person_address where active = 'Y' and address_type = 'Mailing' and person_id = " & checkstring(person_id,50)
  rs.open sql,conn,1,1
  if not rs.eof then person_address_id_1 = rs("person_address_id")
  rs.close
  ' Po: 05/24/2017 - No 'Mailing' typed address found.  Get first active and primary address
  ' to set as default in drop-down list
  if Len(person_address_id_1) = 0 then
      sql = "select person_address_id from person_address where active = 'Y' and primary_flag = 'Y' and person_id = " & checkstring(person_id,50)
      rs.open sql,conn,1,1
      if not rs.eof then person_address_id_1 = rs("person_address_id")
      rs.close
  end if

end if
%>
<script>
function CheckAddress(which) {
  if (which == 1) {
    if (document.dataform.person_address_id_1.selectedIndex != 0) {
      document.dataform.other_address_1.value = '';
    }
  }
  if (which == 2) {
    if (document.dataform.other_address_1.value != '') {
      document.dataform.person_address_id_1.selectedIndex = 0;
    }
  }
}
</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<input type=hidden name="cid" value="<%=request("cid")%>">
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="zone_name_<%=i%>" value="<%=zone_name%>">
<input type=hidden name="correspondence_id_<%=i%>" value="<%=eval("correspondence_id_"&i)%>">
<!--
<b>Correspondence <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
-->
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Letter:&nbsp;</font></td>
<td align=left colspan=3><select name="letter_id_<%=i%>" style="width:305px" onchange="DoChange(<%=i%>);parent.DoRefresh();"><option value="">(delete)</option>
<%
for j = 0 to letters
  isselected = ""
  if toNum(letter_ids(j)) = toNum(eval("letter_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & letter_ids(j) & """" & isselected & ">" & letter_names(j) & "</option>")
next
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>Status:&nbsp;</font></td>
<td align=left><select name="status_<%=i%>" style="width:100px" onchange="DoChange(<%=i%>);">
<option value="" <% if eval("status_"&i) = "" then %>selected<% end if %>></option>
<option value="Sent" <% if eval("status_"&i) = "Sent" then %>selected<% end if %>>Sent</option>
<option value="Queued" <% if eval("status_"&i) = "Queued" then %>selected<% end if %>>Queued</option>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Req&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="requested_date_<%=i%>" value="<%=eval("requested_date_"&i)%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.requested_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Sched&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="scheduled_date_<%=i%>" value="<%=eval("scheduled_date_"&i)%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.scheduled_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Print&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="printed_date_<%=i%>" value="<%=eval("printed_date_"&i)%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.printed_date_"&i, "DoChange("&i&")"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Copies:&nbsp;</font></td>
<td align=left><input type=text name="copies_<%=i%>" value="<%=eval("copies_"&i)%>" size=15 maxlength=50 style="width:80px" onchange="DoChange(<%=i%>);"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:325px" onchange="DoChange(<%=i%>);"></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Address:&nbsp;</font></td>
<td align=left colspan=5><select name="person_address_id_<%=i%>" style="width:520px" onchange="CheckAddress(1);DoChange(<%=i%>);"><option value="">(other address)</option>
<%
sql = "select person_address_id, address_type + ': ' + "
sql = sql & "coalesce(line1,'') + ' ' + "
sql = sql & "coalesce(line2,'') + ' ' + "
sql = sql & "coalesce(line3,'') + ' ' + "
sql = sql & "coalesce(city,'') + ' ' + "
sql = sql & "coalesce(state,'') + ' ' + "
sql = sql & "coalesce(zip,'') + ' ' + "
sql = sql & "coalesce(province,'') + ' ' + "
sql = sql & "coalesce(country,'') address_name "
sql = sql & " from person_address where active = 'Y' and person_id = " & checkstring(person_id,50)
rs.open sql,conn,1,1

while not rs.eof
  isselected = ""
  response.Write(eval(i))
  if toNum(rs("person_address_id")) = toNum(eval("person_address_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & rs("person_address_id") & """" & isselected & ">" & rs("address_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>
<tr>
<td align=right valign=top><font face='arial,helvetica' size=-1><br>&nbsp;&nbsp;Other&nbsp;&nbsp;&nbsp;&nbsp;<br>Address:&nbsp;</font></td>
<td align=left valign=top colspan=5><textarea rows=4 cols=50 name="other_address_<%=i%>" style="width:520px" onchange="CheckAddress(2);DoChange(<%=i%>);"><%=eval("other_address_"&i)%></textarea></td>
</tr>
</table>
<%
if letter_id_1 <> "" then
  varlist = ""
  sql = "select * from letter_defs where letter_id = " & checkstring(letter_id_1,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    for i = 1 to 10
      var_desc = rs("var" & right("0"&i,2) & "_desc")
      if isnull(var_desc) then var_desc = ""
      if var_desc <> "" then
        if var_desc = "exclerk_month_id" then
          emi = cstr(eval("var"&right("0"&i,2)&"_1"))
          varlist = varlist & "<tr>"
          varlist = varlist & "<td valign=middle align=right><font face='arial,helvetica' size=-1>Month&nbsp;&nbsp;</font></td>"
          varlist = varlist & "<td valign=middle align=left><select name=""var" & right("0"&i,2) & "_1"" onchange=""DoChange(1);UpdateNotes"&i&"();""><option value="""">(choose)</option>"
          set rs2 = Server.CreateObject("ADODB.RecordSet")
          sql = "select m.exclerk_month_id, dbo.GetSectionName(e.section_id,'Medical') section_name, t.time_period_name + ' ' + convert(varchar(50),c.start_date,1) + ' - ' + convert(varchar(50),c.end_date,1) month_name "
          sql = sql & " from exclerk_user u, exclerk_months m left outer join enrollments e on m.enrollment_id = e.enrollment_id, calendar c, time_periods t "
          sql = sql & " where u.exclerk_user_id = m.exclerk_user_id and m.calendar_id = c.calendar_id and c.time_period_id = t.time_period_id "
          sql = sql & " and u.person_id = " & checkstring(person_id,50)
          sql = sql & " order by c.start_date "
          rs2.Open sql,conn,1,1
          while not rs2.EOF
            fieldnm = "var"&right("0"&i,2)&"_1"
            unjs = unjs & "if (document.dataform." & fieldnm & ".options[document.dataform." & fieldnm & ".selectedIndex].value == '" & rs2("exclerk_month_id") & "') {x.value = '" & rs2("section_name") & "';};"
            if cstr(rs2("exclerk_month_id")) = emi then
              varlist = varlist & "<option value=""" & rs2("exclerk_month_id") & """ selected>" & rs2("month_name") & "</option>"
            else
              varlist = varlist & "<option value=""" & rs2("exclerk_month_id") & """>" & rs2("month_name") & "</option>"
            end if
            rs2.MoveNext
          wend
          rs2.close
          varlist = varlist & "</select><script>"
          varlist = varlist & "function UpdateNotes"&i&"() {x = document.dataform.notes_1;" & unjs & "}"
          varlist = varlist & "</script></td>"
          varlist = varlist & "</tr>"
        elseif var_desc = "enrollment_id" then
          eid = cstr(eval("var"&right("0"&i,2)&"_1"))
          varlist = varlist & "<tr>"
          varlist = varlist & "<td valign=middle align=right><font face='arial,helvetica' size=-1>Section&nbsp;&nbsp;</font></td>"
          varlist = varlist & "<td valign=middle align=left><select name=""var" & right("0"&i,2) & "_1"" onchange=""DoChange(1);UpdateNotes"&i&"();""><option value="""">(choose)</option>"
          set rs2 = Server.CreateObject("ADODB.RecordSet")
          sql = "select e.enrollment_id, dbo.GetSectionName(e.section_id,'Medical') section_name "
          sql = sql & " from enrollments e inner join sections s on e.section_id = s.section_id inner join calendar c on s.calendar_id = c.calendar_id "
          sql = sql & " where e.student_instance_id = " & checkstring(id,50)
          sql = sql & " and exists (select * from student_grades g where g.enrollment_id = e.enrollment_id and g.active = 'Y' and g.grade_type = 'Final') "
          sql = sql & " order by c.start_date "
          'response.Write sql
          rs2.Open sql,conn,1,1
          while not rs2.EOF
            fieldnm = "var"&right("0"&i,2)&"_1"
            unjs = unjs & "if (document.dataform." & fieldnm & ".options[document.dataform." & fieldnm & ".selectedIndex].value == '" & rs2("enrollment_id") & "') {x.value = '" & rs2("section_name") & "';};"
            if cstr(rs2("enrollment_id")) = eid then
              varlist = varlist & "<option value=""" & rs2("enrollment_id") & """ selected>" & rs2("section_name") & "</option>"
            else
              varlist = varlist & "<option value=""" & rs2("enrollment_id") & """>" & rs2("section_name") & "</option>"
            end if
            rs2.MoveNext
          wend
          rs2.close
          varlist = varlist & "</select><script>"
          varlist = varlist & "function UpdateNotes"&i&"() {x = document.dataform.notes_1;" & unjs & "}"
          varlist = varlist & "</script></td>"
          varlist = varlist & "</tr>"
        else
          varlist = varlist & "<tr>"
          varlist = varlist & "<td valign=middle align=right><font face='arial,helvetica' size=-1>" & var_desc & "&nbsp;&nbsp;</font></td>"
          varlist = varlist & "<td valign=middle align=left><input type=text name=""var" & right("0"&i,2) & "_1"" value=""" & eval("var"&right("0"&i,2)&"_1") & """ size=15 style=""width:350px"" onchange=""DoChange(1);""></td>"
          varlist = varlist & "</tr>"
        end if
      end if
    next
  end if
  rs.close
  if varlist <> "" then
    response.write("<center><table border=0 cellspacing=0 cellpadding=0>")
    response.write(varlist)
    response.write("</table></center>")
  end if
end if
%>





<% case "degrees" %>

<%
sql = "select * from degrees order by degree_name"
rs.open sql,conn,1,1
while not rs.eof
  degree_id_list = degree_id_list & rs("degree_id") & "|"
  degree_name_list = degree_name_list & rs("degree_name") & "|"
  rs.movenext
wend
rs.close
degree_ids = split(degree_id_list,"|")
degree_names = split(degree_name_list,"|")
degrees = ubound(degree_ids)-1
sql = "select * from degree_status order by degree_status_name"
rs.open sql,conn,1,1
while not rs.eof
  status_id_list = status_id_list & rs("degree_status_id") & "|"
  status_name_list = status_name_list & rs("degree_status_name") & "|"
  rs.movenext
wend
rs.close
status_ids = split(status_id_list,"|")
status_names = split(status_name_list,"|")
statuses = ubound(status_ids)-1
sql = "select p.program_name, t.track_name, i.student_instance_id from student_instance as i "
sql = sql & " left outer join programs as p on p.program_id = i.program_id "
sql = sql & " left outer join program_track as t on t.program_track_id = i.program_track_id "
sql = sql & " where i.person_id = " & checkstring(person_id,50)
if session("zone") <> "" then sql = sql & " and p.zone_name = " & checkstring(session("zone"),50)
rs.open sql,conn,1,1
while not rs.eof
  instance_id_list = instance_id_list & rs("student_instance_id") & "|"
  instance_name_list = instance_name_list & rs("program_name") & " (" & rs("track_name") & ")" & "|"
  rs.movenext
wend
rs.close
instance_ids = split(instance_id_list,"|")
instance_names = split(instance_name_list,"|")
instances = ubound(instance_ids)-1
%>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no degrees listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<script>
function PickInstitution(x) {
  eval("document.dataform.institution_id_"+x+".value = '';");
  eval("document.dataform.institution_name_"+x+".value = '';");
  window.open("g_institution.asp?fi=institution_id_"+x+"&fn=institution_name_"+x,"InstitutionWindow","width=400,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<% for i = 1 to rows %>
<%
institution_name = ""
if eval("institution_id_"&i) <> "" then
  sql = "select institution_name from institution where institution_id = " & checkstring(eval("institution_id_"&i),50)
  rs.open sql,conn,1,1
  if not rs.eof then
    institution_name = rs("institution_name")
  end if
  rs.close
end if
program_track = ""
otherZone = false
if (eval("student_instance_id_"&i) <> "") and (session("zone") <> "") then
  sql = "select p.zone_name, p.program_name, t.track_name from student_instance as i "
  sql = sql & " left outer join programs as p on p.program_id = i.program_id "
  sql = sql & " left outer join program_track as t on t.program_track_id = i.program_track_id "
  sql = sql & " where i.student_instance_id = " & checkstring(eval("student_instance_id_"&i),50)
  rs.open sql,conn,1,1
  if not rs.eof then
    if session("zone") <> rs("zone_name") then
      program_track = rs("program_name") & " (" & rs("track_name") & ")"
      otherZone = true
    end if
  end if
  rs.close
end if
disabled = ""
if otherZone then disabled = " disabled"
%>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="student_degree_id_<%=i%>" value="<%=eval("student_degree_id_"&i)%>">
<b>Degree <%=i%></b><br>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Degree:&nbsp;</font></td>
<td align=left><select name="degree_id_<%=i%>" style="width:100px" onchange="DoChange(<%=i%>);" <%=disabled%>><option value="-1">(delete)</option>
<%
if toNum(eval("degree_id_"&i)) < 1 then
  response.write("<option value="""" selected></option>")
else
  response.write("<option value=""""></option>")
end if
for j = 0 to degrees
  isselected = ""
  if toNum(degree_ids(j)) = toNum(eval("degree_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & degree_ids(j) & """" & isselected & ">" & degree_names(j) & "</option>")
next
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Start&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="start_date_<%=i%>" value="<%=eval("start_date_"&i)%>" size=15 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);" <%=disabled%>><%if not otherZone then DrawCal "dataform.start_date_"&i, "DoChange("&i&")" end if%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;End&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="end_date_<%=i%>" value="<%=eval("end_date_"&i)%>" size=15 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);" <%=disabled%> id="Text1"><%if not otherZone then DrawCal "dataform.end_date_"&i, "DoChange("&i&")" end if%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Status:&nbsp;</font></td>
<td colspan="2" align=left><select name="degree_status_id_<%=i%>" style="width:200px" onchange="DoChange(<%=i%>);" <%=disabled%> id="Select1"><option value=""></option>
<%
for j = 0 to statuses
  isselected = ""
  if toNum(status_ids(j)) = toNum(eval("degree_status_id_"&i)) then isselected = " selected"
  response.write("<option value=""" & status_ids(j) & """" & isselected & ">" & status_names(j) & "</option>")
next
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Completed:&nbsp;</font></td>
<td colspan="2" align=left><input type=text name="completed_date_<%=i%>" value="<%=eval("completed_date_"&i)%>" size=15 style="width:75px" onchange="validateDate(this, false);DoChange(<%=i%>);" <%=disabled%> id="Text1"><%if not otherZone then DrawCal "dataform.completed_date_"&i, "DoChange("&i&")" end if%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Institution:&nbsp;</font></td>
<td colspan="2" align=left><input type=hidden name="institution_id_<%=i%>" value="<%=eval("institution_id_"&i)%>" id="Hidden1"><input type=text name="institution_name_<%=i%>" value="<%=institution_name%>" size=15 style="width:175px;font-color:#000000" contenteditable=false onchange="DoChange(<%=i%>);" <%=disabled%> id="Text2">
<% if not otherZone then%><a href="JavaScript:PickInstitution('<%=i%>');DoChange(<%=i%>);"><img src="images/search.gif" border=0></a><% end if %></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Completed:&nbsp;</font></td>
<td colspan="2" align=left>

<table border=0 cellspacing=0 cellpadding=0><tr>
<td valign=middle><font face='arial,helvetica' size=-1>Month&nbsp;</font></td>
<td valign=middle><input type=text name="completed_month_<%=i%>" value="<%=eval("completed_month_"&i)%>" size=15 style="width:35px" maxlength="2" onchange="IsNumeric(this, false);DoChange(<%=i%>);" <%=disabled%>></td>
<td valign=middle><font face='arial,helvetica' size=-1>Year&nbsp;</font></td>
<td valign=middle><input type=text name="completed_year_<%=i%>" value="<%=eval("completed_year_"&i)%>" size=15 style="width:50px" maxlength="4" onchange="IsNumeric(this, false);DoChange(<%=i%>);" <%=disabled%>></td>
</tr></table>

</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Instance:&nbsp;</font></td>
<td align=left colspan="2"><select name="student_instance_id_<%=i%>" style="width:255px" onchange="DoChange(<%=i%>);" <%=disabled%> id="Select2"><option value=""></option>
<%
if otherZone then
  response.write("<option value=""" & eval("student_instance_id_"&i) & """ selected>" & program_track & "</option>")
else
  for j = 0 to instances
    isselected = ""
    if toNum(instance_ids(j)) = toNum(eval("student_instance_id_"&i)) then isselected = " selected"
    response.write("<option value=""" & instance_ids(j) & """" & isselected & ">" & instance_names(j) & "</option>")
  next
end if
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Honors:&nbsp;</font></td>
<td colspan="2" align=left><input type=text name="honors_<%=i%>" value="<%=eval("honors_"&i)%>" size=15 style="width:170px" onchange="DoChange(<%=i%>);" <%=disabled%>></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:530px" onchange="DoChange(<%=i%>);" <%=disabled%>></td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=5><br><% end if %>
<% next %>


<% case "termbill" %>

<%
theTerm = -1
sql = "select * from terms as t order by t.start_date desc"
rs.open sql,conn,1,1
while not rs.eof
  term_id_list = term_id_list & rs("term_id") & "|"
  term_name_list = term_name_list & rs("term_name") & "|"
  if (date >= cdate(rs("start_date"))) and (date < cdate(rs("end_date"))) then
    theTerm = rs("term_id")
  end if
  rs.movenext
wend
rs.close
if request("term_filter") <> "" then
  theTerm = toNum(request("term_filter"))
end if
term_ids = split(term_id_list,"|")
term_names = split(term_name_list,"|")
terms = ubound(term_ids)-1
for i = 0 to terms
  term_ids(i) = toNum(term_ids(i))
next
sql = "select b.* from term_billing_charge_items as b, terms as t where b.term_id = t.term_id"
if session("zone") <> "" then sql = sql & " and b.zone_name = " & checkstring(session("zone"),50)
sql = sql & " order by t.start_date, b.zone_name, b.charge_item_name, b.amount desc"
rs.open sql,conn,1,1
while not rs.eof
  charge_id_list = charge_id_list & rs("possible_charge_id") & "|"
  x = rs("charge_item_name")
  if session("zone") = "" then x = "[" & rs("zone_name") & "] " & x
  if not isnull(rs("charge_type")) then x = x & " (" & rs("charge_type") & ")"
  x = x & " $" & rs("amount")
  charge_name_list = charge_name_list & x & "|"
  charge_term_list = charge_term_list & rs("term_id") & "|"
  rs.movenext
wend
rs.close
charge_ids = split(charge_id_list,"|")
charge_names = split(charge_name_list,"|")
charge_terms = split(charge_term_list,"|")
charges = ubound(charge_ids)-1
for i = 0 to charges
  charge_ids(i) = toNum(charge_ids(i))
  charge_terms(i) = toNum(charge_terms(i))
next
dim termCount()
redim termCount(terms+1)
for i = 0 to terms
  termCount(i) = 0
next
dim thisTerm()
redim thisTerm(rows)
for i = 1 to rows
  thisTerm(i) = theTerm
  for j = 0 to charges
    if charge_ids(j) = toNum(eval("possible_charge_id_"&i)) then thisTerm(i) = charge_terms(j)
  next
  for j = 0 to terms
    if term_ids(j) = thisTerm(i) then
      termCount(j) = termCount(j) + 1
    end if
  next
next
%>
<script>
current_term = 0;
function ChangeTermFilter() {
  if (parent.ChangesMade) {
    if (!confirm('You have not saved your changes. Would you like to continue anyway?')) {
      document.dataform.term_filter.selectedIndex = current_term;
      return;
    } else {
      parent.ChangesMade = false;
      <% if isRole("B") then %>
      parent.footerform.bsave.style.fontWeight = 'normal';
      <% end if %>
    }
  }
  document.location = 'p_details.asp?pg=termbill&id=<%=id%>&task=edit&term_filter='+document.dataform.term_filter.options[document.dataform.term_filter.selectedIndex].value;
}
<% if theTerm < 0 then %>
function ChangeTerm(fTerm,fItems) {
  term_id = fTerm.options[fTerm.selectedIndex].value;
  x = fItems;
  x.options.length=0;
<%
prevTerm = -1
option_list = """"""
value_list = """"""
for i = 0 to charges
  if (charge_terms(i) <> prevTerm) and (prevTerm >= 0) then
%>
  if (term_id == '<%=prevTerm%>') {
    a = new Array(<%=option_list%>);
    b = new Array(<%=value_list%>);
    for (i = 0; i < a.length; i++) {
      x.options[i] = new Option(a[i]);
      x.options[i].value = b[i];
    }
  }
<%
    option_list = """"""
    value_list = """"""
  end if
  prevTerm = charge_terms(i)
  option_list = option_list & ",""" & charge_names(i) & """"
  value_list = value_list & ",""" & charge_ids(i) & """"
next
%>
  if (term_id == '<%=prevTerm%>') {
    a = new Array(<%=option_list%>);
    b = new Array(<%=value_list%>);
    for (i = 0; i < a.length; i++) {
      x.options[i] = new Option(a[i]);
      x.options[i].value = b[i];
    }
  }
  x.selectedIndex = 0;
}
<% end if %>
</script>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<center>
<table border=0 cellspacing=0 cellpadding=1><tr><td bgcolor=AAAAAA>
<table border=0 cellspacing=0 cellpadding=5>
<tr>
<td align=right bgcolor=DDDDDD><font face='arial,helvetica' size=-1><b>Term&nbsp;Filter</b>:</font></td>
<td align=left bgcolor=DDDDDD><select name="term_filter" style="width:150px" onchange="ChangeTermFilter();">
<option value="-1" <% if theTerm < 0 then %>selected<% end if %>>ALL  [<%=rows%>]</option>
<%
x = 0
theTermName = ""
for j = 0 to terms
  isselected = ""
  if term_ids(j) = theTerm then
    isselected = " selected"
    theTermName = term_names(j)
    x = j+1
  end if
  response.write("<option value=""" & term_ids(j) & """" & isselected & ">" & term_names(j) & "  [" & termCount(j) & "]</option>")
next
%>
</select>
</td>
</tr>
</table>
</td></tr></table>
<script>current_term = <%=x%>;</script>
</center>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<%
ii = 0
for i = 1 to rows
  if (theTerm >= 0) and (theTerm <> thisTerm(i)) and (thisTerm(i) > 0) then
%>
<input type=hidden name="possible_charge_id_<%=i%>" value="<%=eval("possible_charge_id_"&i)%>">
<input type=hidden name="date_sent_<%=i%>" value="<%=eval("date_sent_"&i)%>">
<input type=hidden name="notes_<%=i%>" value="<%=eval("notes_"&i)%>">
<%
  else
    ii = ii + 1
%>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<b>Charge Item <%=ii%></b><br>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="term_billing_id_<%=i%>" value="<%=eval("term_billing_id_"&i)%>">
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Term:&nbsp;</font></td>
<% if theTerm > 0 then %>
<td align=left><input type=text name="term_name_<%=i%>" value="<%=theTermName%>" size=15 style="width:115px" onchange="DoChange(<%=i%>);" disabled></td>
<% else %>
<td align=left><select name="term_id_<%=i%>" style="width:115px" onchange="ChangeTerm(document.dataform.term_id_<%=i%>,document.dataform.possible_charge_id_<%=i%>);DoChange(<%=i%>);"><option value=""></option>
<%
for j = 0 to terms
  isselected = ""
  if term_ids(j) = thisTerm(i) then isselected = " selected"
  response.write("<option value=""" & term_ids(j) & """" & isselected & ">" & term_names(j) & "</option>")
next
%>
</select>
</td>
<% end if %>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Item:&nbsp;</font></td>
<td align=left><select name="possible_charge_id_<%=i%>" style="width:350px" onchange="DoChange(<%=i%>);"><option value="">(delete)</option>
<%
for j = 0 to charges
  isselected = ""
  if charge_ids(j) = toNum(eval("possible_charge_id_"&i)) then isselected = " selected"
  if charge_terms(j) = thisTerm(i) then
    response.write("<option value=""" & charge_ids(j) & """" & isselected & ">" & charge_names(j) & "</option>")
  end if
next
%>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Sent&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="date_sent_<%=i%>" value="<%=eval("date_sent_"&i)%>" size=15 maxlength=50 style="width:90px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.date_sent_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:350px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<%
  end if
next
if ii = 0 then
%>
<img src="images/spacer.gif" border=0 width=1 height=5><br>
<center><i>no term bill items listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>






<% case "tracks" %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>
<font size=-1>
<% if rows = 0 then %>
<center><i>no program tracks listed</i></center>
<img src="images/spacer.gif" border=0 width=550 height=1><br>
<% end if %>
<% for i = 1 to rows %>
<input type=hidden name="hasChanged_<%=i%>" value="<%=eval("hasChanged_"&i)%>">
<input type=hidden name="person_id_<%=i%>" value="<%=eval("person_id_"&i)%>">
<input type=hidden name="student_instance_id_<%=i%>" value="<%=eval("student_instance_id_"&i)%>">
<input type=hidden name="student_program_track_id_<%=i%>" value="<%=eval("student_program_track_id_"&i)%>">
<b>Program Track <%=i%></b><br>
<%
program_name = ""
track_name = ""
sql = "select g.program_name, g.program_id, t.track_name from programs g, program_track t where g.program_id = t.program_id and program_track_id = " & checkstring(eval("program_track_id_"&i),50)
rs.open sql,conn,1,1
if not rs.eof then
  program_name = rs("program_name")
  track_name = rs("track_name")
end if
rs.close
%>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Program:&nbsp;</font></td>
<td align=left><input type=text name="program_name_<%=i%>" style="width:185px;background-color:DDDDDD" ContentEditable=false value="<%=program_name%>"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Start&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="start_date_<%=i%>" value="<%=eval("start_date_"&i)%>" size=15 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.start_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Active:&nbsp;</font></td>
<td align=left><nobr><font face='arial,helvetica' size=-1><% if eval("active_"&i) = "Y" then %>yes<% else %>no<% end if %><input type=hidden name="active_<%=i%>" value="<%=eval("active_"&i)%>"></font></nobr></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Track:&nbsp;</font></td>
<td align=left><input type=text name="track_name_<%=i%>" style="width:185px;background-color:DDDDDD" ContentEditable=false value="<%=track_name%>"><input type=hidden name="program_track_id_<%=i%>" value="<%=eval("program_track_id_"&i)%>"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;End&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="end_date_<%=i%>" value="<%=eval("end_date_"&i)%>" size=15 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.end_date_"&i, "DoChange("&i&")"%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Complete&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="completed_date_<%=i%>" value="<%=eval("completed_date_"&i)%>" size=15 style="width:80px" onchange="validateDate(this, false);DoChange(<%=i%>);"><%DrawCal "dataform.completed_date_"&i, "DoChange("&i&")"%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="notes_<%=i%>" value="<%=eval("notes_"&i)%>" size=15 style="width:540px" onchange="DoChange(<%=i%>);"></td>
</tr>
</table>
<% if i < rows then %><img src="images/spacer.gif" border=0 width=1 height=5><br><% end if %>
<% next %>






<% case else %>

<img src="images/spacer.gif" border=0 width=600 height=150><br>
<center>
<font size=-1>Under Construction</font>
</center>
<img src="images/spacer.gif" border=0 width=600 height=150><br>

<% end select %>

</font>


</td></tr></table>
</td></tr></table>


</td>
<td width=1></form></td>
</tr>
</table>
</body>
</html>
<%
end sub






sub PrintShell
  session("pagetab") = "People"
  if pg = "correspond2" then
    session("history") = "Menu|p_menu.asp|Select a Student|p_select.asp|Summary|p_summary.asp?id=" & id & "|Correspondences|p_details.asp?pg=correspond&id=" & id & "|" & pgname
  else
    session("history") = "Menu|p_menu.asp|Select a Student|p_select.asp|Summary|p_summary.asp?id=" & id & "|" & pgname
  end if
  PageHeader
%>
<script>
ChangesMade = false;
function DoBack() {
<% if pg = "correspond2" then %>
  GoPage('p_details.asp?pg=correspond&id=<%=id%>');
<% else %>
  GoPage('p_summary.asp?id=<%=id%>');
<% end if %>
}
function DoChange() {
  if (!ChangesMade) {
    ChangesMade = true;
    <% if isRole("B") then %>
    document.footerform.bsave.disabled = false;
    document.footerform.breset.disabled = false;
    document.footerform.bsave.style.fontWeight = 'bold';
    <% end if %>
  }
}
function UndoChanges() {
  //ChangesMade = false;
  //details.document.dataform.reset();
  //document.footerform.bsave.style.fontWeight = 'normal';
  document.location = 'p_details.asp?pg=<%=server.urlencode(pg)%>&id=<%=server.urlencode(id)%>&cid=<%=server.urlencode(request("cid"))%>';
}
function NewRecord() {
<% if pg = "correspond" then %>
  document.location = 'p_details.asp?pg=correspond2&id=<%=server.urlencode(id)%>';
<% else %>
  details.document.dataform.task.value="add";
  details.document.dataform.submit();
<% end if %>
  DoChange();
}
function DoRefresh() {
  details.document.dataform.task.value="refresh";
  details.document.dataform.submit();
  DoChange();
}
function SaveChanges() {
  var f = details.document.dataform;
  var e;
  var n = 0;
  for (var i = 0; i < f.elements.length; i++) {
    e = f.elements[i];
    if (e.type == 'select-one') {
      if (e.options[e.selectedIndex].text == '(delete)') {
        n++;
      }
    }
  }
  if (n > 0) {
    if (confirm('You are about to delete one or more records. Are you sure you want to continue?')) {
      details.document.dataform.task.value="save";
      details.document.dataform.submit();
    }
  } else {
    details.document.dataform.task.value="save";
    details.document.dataform.submit();
  }
}
</script>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0><% if false and request("message") <> "" then %>
<tr><td align=center valign=top>
<table border=0 cellspacing=5 cellpadding=0 width=100%><tr><td align=center bgcolor=000000>
<% if request("message") = "0" then %>
<table border=0 cellspacing=1 cellpadding=2 width=100%><tr><td align=center bgcolor=FFFF99><b><font face='arial,helvetica' size=-1>Your changes have been saved.</font></b></td></tr></table>
<% else %>
<table border=0 cellspacing=1 cellpadding=2 width=100%><tr><td align=center bgcolor=FFCCCC><b><font face='arial,helvetica' size=-1>ERROR - Changes have occurred since you began editing this page.</font></b></td></tr></table>
<% end if %>
</td></tr></table>
</td></tr>
<% end if %><tr><td align=center valign=middle height=100%>
<iframe name="details" src="p_details.asp?id=<%=id%>&pg=<%=pg%>&task=edit&cid=<%=request("cid")%>" scrolling=auto width=100% height=100% frameborder=0></iframe>
</td></tr></table>
<% PageMiddle %>
<td align=center valign=middle>
<% if isRole("B") and (pg <> "instances") then %>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="SaveChanges();" name="bsave" value="Save">
<img src="images/spacer.gif" border=0 width=20 height=1>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="UndoChanges();" name="breset" value="Reset">
<img src="images/spacer.gif" border=0 width=20 height=1>
<% if bNew then %>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="NewRecord();" name="bnew" value="New">
<img src="images/spacer.gif" border=0 width=20 height=1>
<% end if %>
<% end if %>
<% if request("isnew") = "Y" then %>
<input style="width:75px; background-color:#CCCCCC" type=button onClick="DoBack();" name="bcancel" value="Continue">
<% else %>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="DoBack();" name="bcancel" value="Back">
<% end if %>
</td>
<%
  PageFooter
end sub
%>