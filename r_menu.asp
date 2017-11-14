<!--#include file ="checkusr.asp"-->
<!--#include file ="vb.asp"-->
<!--#include file ="db.asp"-->
<%

function NumGradesReceived(secID)
  sql2 = sql2 & " SELECT count(g.student_grade_id) as count FROM student_grades g, enrollments e"
  sql2 = sql2 & " WHERE e.enrollment_id = g.enrollment_id AND g.active = 'Y' "
  sql2 = sql2 & " AND g.approved = 'Y' AND grade_value_id IS NOT null"
  sql2 = sql2 & " AND e.section_id=" & rs("section_id")
  set rs2 = Server.CreateObject("ADODB.RecordSet")
  rs2.open sql2,conn,1,1
  MyGradesReceived = -1
  if not rs2.EOF then
    MyGradesReceived = rs2("count")  
  end if
  rs2.close
  NumGradesReceived = MyGradesReceived
end function

function NumEnrolled(secID)
  sql2 = sql2 & " SELECT count(e.enrollment_id) as count FROM enrollments e"
  sql2 = sql2 & " WHERE e.enrollment_status_id = 18 AND e.section_id=" & rs("section_id")
  set rs2 = Server.CreateObject("ADODB.RecordSet")
  rs2.open sql2,conn,1,1
  myEnrolled = 0
  if not rs2.EOF then
    myEnrolled = rs2("count")  
  end if
  rs2.close
  NumEnrolled = myEnrolled
end function

set rs = Server.CreateObject("ADODB.RecordSet")

task = request("task")
report = request("report")
term_id = request("term_id")
mycoursetype = request.QueryString("coursetype")
if mycoursetype<>"" then 
mycoursetype="&coursetype=" & mycoursetype 
end if
if task = "shell" then
  drawSessionZone = true
  changeSessionZone = "document.location='r_menu.asp?task=shell&report=" & report & mycoursetype & "'"
  drawSessionYear = true
  changeSessionYear = "document.location='r_menu.asp?task=shell&report=" & report & mycoursetype & "'"
  PrintShell
elseif task = "details" then
  PrintDetails
else
  drawSessionZone = true
  changeSessionZone = "document.location='r_menu.asp'"
  PrintMenu
end if

sub PrintMenu
session("pagetab") = "Reports"
session("history") = "Menu"
PageHeader
m = ""
m = m & "Over/Under|r_menu.asp?task=shell&report=overunder|"
m = m & "Temp Enrollments|r_menu.asp?task=shell&report=tempenrollments|"
m = m & "Enrollment Conflicts|r_menu.asp?task=shell&report=enrollmentconflicts|"
m = m & "Queued Letters|queuedletters.asp|"
m = m & "Overdue Grades|r_menu.asp?task=shell&report=missinggrades|"
m = m & "Grades Not Approved|r_menu.asp?task=shell&report=gradesnotapproved|"
m = m & "Incomplete Grades|r_menu.asp?task=shell&report=incompletegrades|"
m = m & "Printed Reports|reportpage.asp|"
'm = m & "Add/Drop Requests|JavaScript:alert(\'Under Construction\');|"
m = split(m,"|")
n = ubound(m)/2
%>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td valign='center' align='center'>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr>

<td width=10%><img src="images/spacer.gif" border=0 width=20 height=1></td>

<td width=40% align="center" valign="top">
<% for i = 0 to int((n-1)/2) %>
<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align="center" valign="middle" onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='<%=m(i*2+1)%>'">
<font face='arial,helvetica'><b><%=m(i*2)%></b><br></font>
</td></tr></table>
</td></tr></table>
<img src="images/spacer.gif" border=0 width=1 height=20><br>
<% next %>
</td>

<td><img src="images/spacer.gif" border=0 width=20 height=1></td>

<td width=40% align="center" valign="top">
<% for i = int((n-1)/2)+1 to n-1 %>
<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align="center" valign="middle" onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='<%=m(i*2+1)%>'">
<font face='arial,helvetica'><b><%=m(i*2)%></b><br></font>
</td></tr></table>
</td></tr></table>
<img src="images/spacer.gif" border=0 width=1 height=20><br>
<% next %>
</td>

<td width=10%><img src="images/spacer.gif" border=0 width=20 height=1></td>

</tr>
</table>
</td></tr></table>
<%
PageMiddle
PageFooter
end sub

sub PrintShell
  session("pagetab") = "Reports"
  select case report
  case "enrollmentconflicts"
    session("history") = "Menu|r_menu.asp|Enrollment Conflicts"
  case "overunder"
    session("history") = "Menu|r_menu.asp|Over/Under"
  case "tempenrollments"
    session("history") = "Menu|r_menu.asp|Temp Enrollments"
  case "missinggrades"
    session("history") = "Menu|r_menu.asp|Overdue Grades"
  case "gradesnotapproved"
    session("history") = "Menu|r_menu.asp|Grades Not Approved"
  case "incompletegrades"
    session("history") = "Menu|r_menu.asp|Incomplete Grades"
  case "queuedletters"
    session("history") = "Menu|r_menu.asp|Queued Letters"
  case else
    response.redirect("r_menu.asp")
  end select
  
  PageHeader  
%>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
 <td align="right" valign="top">
  <input style="background-color:#CCCCCC" type="button" onclick="parent.details.focus();parent.details.print();" name="bprint" value="Print Report">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr>
<td valign="middle" height="100%">
<iframe name="details" id="details" src="r_menu.asp?task=details&printview=yes&report=<%=report%><%=mycoursetype%>" scrolling=auto width=100% height=100% frameborder=0></iframe>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align="center" valign="middle"><nobr>
<font face='arial,helvetica' size='-1'><div id="searchstatus"><b>Generating&nbsp;Report...</b></div></font>
</nobr></td>
<%
PageFooter
end sub


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
-->
</style>
<script>
function DoOnLoad() {
  parent.document.getElementById('searchstatus').innerHTML = '<input style="width:100px; background-color:#CCCCCC" type=button onclick="GoPage(\'r_menu.asp\');" name="bcancel" value=" Back ">';
}
</script>
</head>
<body bgcolor="F9F9F9" text="000000" link="000000" vlink="000000" alink="000000" onLoad="DoOnLoad();">

<script language=javascript>
function usefilter(rpt){
str = dataform.drpcourse_type.value;
if (rpt=='missinggrades'){
parent.location = "r_menu.asp?task=shell&report=missinggrades&coursetype=" + str;}
if (rpt=='overunder'){
parent.location = "r_menu.asp?task=shell&report=overunder&coursetype=" + str;}
}
</script>

<%

  response.write("<center>")
  response.write("<table border=0 cellspacing=0 cellpadding=0><tr>")
  response.write("<td><form name=dataform method=""GET"" action=""r_menu.asp""><input type=hidden name=""task"" value=""details""><input type=hidden name=""report"" value=""" & report & """></td>")
  response.write("<td valign=top align='center'>")

  select case report
  case "enrollmentconflicts"
    response.write("<font face='arial,helvetica' color='#666666'><b>Enrollment Conflicts</b></font>")
  case "overunder"
    response.write("<font face='arial,helvetica' color='#666666'><b>Over/Under</b></font>")
  case "tempenrollments"
    response.write("<font face='arial,helvetica' color='#666666'><b>Temp Enrollments</b></font>")
  case "missinggrades"
    response.write("<font face='arial,helvetica' color='#666666'><b>Overdue Grades</b></font>")
  case "gradesnotapproved"
    response.write("<font face='arial,helvetica' color='#666666'><b>Grades Not Approved</b></font>")
  case "incompletegrades"
    response.write("<font face='arial,helvetica' color='#666666'><b>Incomplete Grades</b></font>")
  case "queuedletters"
    response.write("<font face='arial,helvetica' color='#666666'><b>Queued Letters</b></font>")
  end select
  
  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=10><br>")
  response.write("<font face='arial,helvetica' size='-1'>")
  response.write("<b>Today</b>: " & date)
  response.write("&nbsp;&nbsp;&nbsp;&nbsp;")
  zone = session("zone")
  if zone = "" then zone = "Any"
  response.write("<b>Zone</b>: " & zone)
  if instr("|enrollmentconflicts|missinggrades|","|"&report&"|") > 0 OR instr("|incompletegrades|","|"&report&"|") > 0 then
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;")
    theyear = "Any"
    if session("year") <> "" then
      sql = "select acad_year_name from acad_year where acad_year_id = '" & session("year") & "'"
      rs.open sql,conn,1,1
      if not rs.eof then
        theyear = rs("acad_year_name")
      end if
      rs.close
    end if
    response.write("<b>Year</b>: " & theyear)
  end if
  
  if instr("|missinggrades|","|"&report&"|") > 0 OR instr("|overunder|","|"&report&"|") > 0 then
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;")
    thecoursetype = server.HTMLEncode(request.QueryString("coursetype"))
    response.write("<b>Course Type</b>: ")
    response.Write("<select name='drpcourse_type' style='width:175px' onchange=""usefilter('" & report & "');""><option value='All'>All</option>")
    if session("zone") = "" then isdistinct = "distinct"
     sql = "select " & isdistinct & " course_type from course_type"
     if session("zone") <> "" then sql = sql & " WHERE zone_name = " & checkstring(session("zone"),50)
      sql = sql & " order by course_type"
      rs.open sql,conn,1,1
      while not rs.eof
       isselected = ""
        if rs("course_type") = server.HTMLEncode(thecoursetype) then isselected = " selected"
         response.write("<option value=""" & rs("course_type") & """" & isselected & ">" & rs("course_type") & "</option>")
         rs.movenext
      wend
      rs.close
      response.Write("</select>")
     end if
  
  response.write("<br>")

  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=10><br>")

  select case report
  case "enrollmentconflicts"
    sql = "SELECT p.full_name, a.student_instance_id, a.section_id section_id1, b.section_id section_id2,  "
    sql = sql & "dbo.GetSectionName(a.section_id,g.zone_name) section_name1, dbo.GetSectionName(b.section_id,g.zone_name) section_name2 "
    sql = sql & "FROM person p, student_instance i, programs g, "
    sql = sql & "( "
    sql = sql & "SELECT e.enrollment_id, e.student_instance_id, e.section_id,  "
    sql = sql & "(t.start_hour*60+t.start_minute) start_time, "
    sql = sql & "(t.end_hour*60+t.end_minute) end_time, "
    sql = sql & "d.day_of_week_code "
    sql = sql & "FROM enrollments e, enrollment_status u, section_times t, days_of_week d, sections s, offerings o "
    sql = sql & "WHERE e.section_id = t.section_id AND t.day_of_week_id = d.day_of_week_id "
    sql = sql & "AND e.enrollment_status_id = u.enrollment_status_id and u.enrollment_group = 'E' "
    sql = sql & "AND e.section_id = s.section_id AND s.offering_id = o.offering_id "
    if session("year") <> "" then
      sql = sql & "AND o.acad_year_id = '" & session("year") & "' "
    end if
    sql = sql & "AND ((t.start_hour*60+t.start_minute) IS NOT null) AND ((t.end_hour*60+t.end_minute) IS NOT null) "
    sql = sql & "AND d.day_of_week_code > 0 "
    sql = sql & ") a, "
    sql = sql & "( "
    sql = sql & "SELECT e.enrollment_id, e.student_instance_id, e.section_id,  "
    sql = sql & "(t.start_hour*60+t.start_minute) start_time, "
    sql = sql & "(t.end_hour*60+t.end_minute) end_time, "
    sql = sql & "d.day_of_week_code "
    sql = sql & "FROM enrollments e, enrollment_status u, section_times t, days_of_week d, sections s, offerings o "
    sql = sql & "WHERE e.section_id = t.section_id AND t.day_of_week_id = d.day_of_week_id "
    sql = sql & "AND e.enrollment_status_id = u.enrollment_status_id and u.enrollment_group = 'E' "
    sql = sql & "AND e.section_id = s.section_id AND s.offering_id = o.offering_id "
    if session("year") <> "" then
      sql = sql & "AND o.acad_year_id = '" & session("year") & "' "
    end if
    sql = sql & "AND ((t.start_hour*60+t.start_minute) IS NOT null) AND ((t.end_hour*60+t.end_minute) IS NOT null) "
    sql = sql & "AND d.day_of_week_code > 0 "
    sql = sql & ") b "
    sql = sql & "WHERE a.student_instance_id = b.student_instance_id AND a.section_id = b.section_id "
    sql = sql & "AND a.enrollment_id < b.enrollment_id "
    sql = sql & "AND a.start_time < b.end_time AND a.end_time > b.start_time "
    sql = sql & "AND (a.day_of_week_code & b.day_of_week_code > 0) "
    sql = sql & "AND p.person_id = i.person_id AND i.student_instance_id = a.student_instance_id "
    sql = sql & "AND i.program_id = g.program_id "
    if session("zone") <> "" then
      sql = sql & "AND g.zone_name = '" & session("zone") & "' "
    end if
    sql = sql & "ORDER BY p.full_name, section_name1, section_name2 "
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("<i>no enrollment conflicts</i>")
    else
      response.write("<table border='0' cellspacing='0' cellpadding='0'><tr><td bgcolor='#CCCCCC' valign='center'>")
      response.write("<table border='0' cellspacing='1' cellpadding='1'>")
      response.write("<tr>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Student&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Section&nbsp;1&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Section&nbsp;2&nbsp;</b></font></td>")
      response.write("</tr>")
      while not rs.eof
        response.write("<tr>")
        response.write("<td bgcolor='#FFFFFF' onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" align='center' onClick=""parent.location = 'p_summary.asp?id=" & rs("student_instance_id") & "'""><font face='arial,helvetica' size='-1'>&nbsp;" & rs("full_name") & "&nbsp;</font></td>")
        response.write("<td bgcolor='#FFFFFF' onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" align='center' onClick=""parent.location = 'c_enrollment.asp?as=Y&id=" & rs("section_id1") & "'""><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_name1") & "&nbsp;</font></td>")
        response.write("<td bgcolor='#FFFFFF' onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" align='center' onClick=""parent.location = 'c_enrollment.asp?as=Y&id=" & rs("section_id2") & "'""><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_name2") & "&nbsp;</font></td>")
        response.write("</tr>")
        rs.movenext
      wend
      response.write("</table>")
      response.write("</td></tr></table>")
    end if
    rs.close
  case "overunder"
    sql = " select t.*, (case when o.transcript_title <> '' then o.transcript_title else c.course_title end) section_title "
    sql = sql & " from ( "
    'sql = sql & " select dbo.GetSectionName(s.section_id,'" & session("zone") & "') section_name, s.class_start, s.section_id, "
    sql = sql & " select dbo.GetSectionName(s.section_id,'') section_name, dbo.GetOfferingName(s.offering_id,'') courseid,  s.class_start, s.section_id, "
   
    sql = sql & " max(min_enrollment) min_enrollment, max(max_enrollment) max_enrollment, max(offering_id) offering_id, "
    sql = sql & " sum(case when t.enrollment_group in ('E') then 1 else 0 end) enrolled_e,  "
    sql = sql & " sum(case when t.enrollment_group in ('T') then 1 else 0 end) enrolled_et, "
    sql = sql & "(Select convert(int,StudentCnt) from dbo.OASIS_Enrollments oe WHERE oe.CID = dbo.GetOfferingName(s.offering_id,'') AND oe.Start_dt =s.class_start) enrolled_oasis "
    sql = sql & " from sections s, calendar d, enrollments e, enrollment_status t "
    sql = sql & " where s.calendar_id = d.calendar_id "
    sql = sql & " and coalesce(s.class_start, d.start_date) >= GetDate() "
    sql = sql & " and coalesce(s.class_start, d.start_date) <= (DATEADD(m,14,GetDate())) "
    sql = sql & " and s.section_id = e.section_id and e.enrollment_status_id = t.enrollment_status_id "
    sql = sql & " group by s.class_start, s.section_id, s.offering_id "
    sql = sql & " ) t, offerings o, course c "
    sql = sql & " where t.offering_id = o.offering_id and o.course_id = c.course_id "
    if thecoursetype <> "" AND thecoursetype <> "All" then
      sql = sql & " AND c.course_type='" & thecoursetype & "'" 
  end if
    if session("zone") <> "" then sql = sql & " and " & session("zone") & "_number is not null "
    sql = sql & " and (enrolled_e < coalesce(min_enrollment,0) "
    sql = sql & " or enrolled_e > coalesce(max_enrollment,999999) "
    sql = sql & " or enrolled_et < coalesce(min_enrollment,0) "
    sql = sql & " or enrolled_et > coalesce(max_enrollment,999999) "
    sql = sql & " or enrolled_oasis < coalesce(min_enrollment,0) "
    sql = sql & " or enrolled_oasis > coalesce(max_enrollment,999999)) "
    sql = sql & " order by class_start, section_name "
    'response.Write sql
    'response.end
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("<i>no under/over sections</i>")
    else
   
      response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign='center'>")
      response.write("<table border=0 cellspacing=1 cellpadding=1>")
      response.write("<tr>")
      response.write("<td bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Section&nbsp;</b></font></td>")
      response.write("<td bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Title&nbsp;</b></font></td>")
      
    response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Enrolled&nbsp;MADRIS&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Enrolled&nbsp;OASIS&nbsp;</b></font></td>")
        
    response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Temp&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Min&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Max&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Unfilled&nbsp;</b></font></td>")
      response.write("</tr>")
   
      while not rs.eof
        enrolled_e = cint(rs("enrolled_e"))
        enrolled_et = cint(rs("enrolled_et"))
        enrolled_oa = rs("enrolled_oasis")
    if isnull(enrolled_oa) then
        enrolled_o = 0
        else
        enrolled_o = cint(rs("enrolled_oasis"))
    end if

        min_enrollment = rs("min_enrollment")
        max_enrollment = rs("max_enrollment")
        'oasis
        

        if isnull(min_enrollment) then
          min_enrollment = "&nbsp;"
          min_e = 0
        else
          min_e = cint(min_enrollment)
        end if
        if isnull(max_enrollment) then
          max_enrollment = "&nbsp;"
          max_e = 999999
          unfilled = "&nbsp;"
        else
          max_e = cint(max_enrollment)
          unfilled = max_e - enrolled_e - enrolled_o
        end if
        response.write("<tr bgcolor='#FFFFFF' onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" onClick=""parent.location = 'c_enrollment.asp?as=Y&id=" & rs("section_id") & "'"">")
        response.write("<td><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_name") & "&nbsp;</font></td>")
        response.write("<td><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_title") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & enrolled_e & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & enrolled_o & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & enrolled_et & "&nbsp;</font></td>")
        if enrolled_e < min_e then
          response.write("<td align='right' bgcolor='#c0c0c0'><font face='arial,helvetica' size='-1' color=990000><b>" & min_enrollment & "</b>&nbsp;</font></td>")
        else
          response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & min_enrollment & "&nbsp;</font></td>")
        end if
        if enrolled_e > max_e then
          response.write("<td align='right' bgcolor='#c0c0c0'><font face='arial,helvetica' size='-1' color='#990000'><b>" & max_enrollment & "</b>&nbsp;</font></td>")
          response.write("<td align='right' bgcolor='#c0c0c0'><font face='arial,helvetica' size='-1' color='#990000'><b>" & unfilled & "</b>&nbsp;</font></td>")
        else
          response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & max_enrollment & "&nbsp;</font></td>")
          response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & unfilled & "&nbsp;</font></td>")
        end if
        response.write("</tr>")


    
        rs.movenext
      wend
      response.write("</table>")
      response.write("</td></tr></table>")
      response.write("<img src=""images/spacer.gif"" border=0 width=1 height=5><br>")
      'response.write("<font face='arial,helvetica' size='1'>Temporary enrollments are in parentheses.</font><br>")
    
    end if
    rs.close
  case "tempenrollments"
    sql = " select t.*, (case when o.transcript_title <> '' then o.transcript_title else c.course_title end) section_title "
    sql = sql & " from ( "
    sql = sql & " select dbo.GetSectionName(s.section_id,'Medical') section_name, s.section_id, "
    sql = sql & " max(min_enrollment) min_enrollment, max(max_enrollment) max_enrollment, max(offering_id) offering_id, "
    sql = sql & " sum(case when e.enrollment_status_id = 0 then 1 else 0 end) enrolled_temporary,  "
    sql = sql & " sum(case when e.enrollment_status_id = 18 then 1 else 0 end) enrolled_enrolled,  "
    sql = sql & " sum(case when e.enrollment_status_id = 21 then 1 else 0 end) enrolled_waitlist,  "
    sql = sql & " sum(case when e.enrollment_status_id = 449 then 1 else 0 end) enrolled_pending, "
    ' added sjt5 8/18/2011 for awaiting approval
    sql = sql & " sum(case when e.enrollment_status_id = 450 then 1 else 0 end) enrolled_awaiting "
    sql = sql & " from sections s, calendar d, enrollments e "
    sql = sql & " where s.calendar_id = d.calendar_id "
    sql = sql & " and coalesce(s.class_start, d.start_date) >= GetDate() "
    sql = sql & " and coalesce(s.class_start, d.start_date) <= (GetDate()+60) "
    sql = sql & " and s.section_id = e.section_id "
    sql = sql & " group by s.section_id "
    sql = sql & " ) t, offerings o, course c "
    sql = sql & " where t.offering_id = o.offering_id and o.course_id = c.course_id "
    sql = sql & " and (enrolled_temporary > 0 or enrolled_waitlist > 0 or enrolled_pending > 0 or enrolled_awaiting > 0) "
    sql = sql & " order by section_name "
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("<i>no temporary, waitlist, pending, or awaiting approval enrollments</i>")
    else
      response.write("<table border='0' cellspacing='0' cellpadding='0'><tr><td bgcolor=CCCCCC valign='center'>")
      response.write("<table border='0' cellspacing='1' cellpadding='1'>")
      response.write("<tr>")
      response.write("<td bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Section&nbsp;</b></font></td>")
      response.write("<td bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Title&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Enrolled&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Temporary&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Waitlist&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Pending&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Awaiting Approval&nbsp;</b></font></td>")
      response.write("</tr>")
      while not rs.eof
        response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" onClick=""parent.location = 'c_enrollment.asp?as=Y&id=" & rs("section_id") & "'"">")
        response.write("<td><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_name") & "&nbsp;</font></td>")
        response.write("<td><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_title") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & rs("enrolled_enrolled") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & rs("enrolled_temporary") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & rs("enrolled_waitlist") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & rs("enrolled_pending") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>" & rs("enrolled_awaiting") & "&nbsp;</font></td>")
        response.write("</tr>")
        rs.movenext
      wend
      response.write("</table>")
      response.write("</td></tr></table>")
    end if
    rs.close
  case "missinggrades","gradesnotapproved","incompletegrades"
    if session("zone") <> "" then
      sql = "SELECT e.section_id, o.transcript_title, dbo.GetSectionName(e.section_id, '" & session("zone") & "') section_name, count(*) n "
    else
      sql = "SELECT e.section_id, o.transcript_title, dbo.GetSectionName(e.section_id, c.zone_name) section_name, count(*) n "
    end if
    sql = sql & " FROM enrollments e, sections s, offerings o, course c, calendar d "
    sql = sql & " WHERE e.section_id = s.section_id AND s.offering_id = o.offering_id AND o.course_id = c.course_id "
    sql = sql & " AND s.calendar_id = d.calendar_id AND e.enrollment_status_id = '18' "
    if thecoursetype <> "" AND thecoursetype <> "All" then
      sql = sql & " AND c.course_type='" & thecoursetype & "'" 
    end if
    if session("year") <> "" then
      sql = sql & " AND o.acad_year_id = " & checkstring(session("year"),50)
    end if
    sql = sql & " AND c.course_id NOT IN (4501,4502,4503,5741,4504,4506,4505,4507,4508,4511,4512,5523) "
    if report = "missinggrades" then
      sql = sql & " AND COALESCE(e.end_date,s.class_end,d.end_date) < (GetDate()-21) "
      sql = sql & " AND NOT EXISTS ( "
      sql = sql & " SELECT g.student_grade_id FROM student_grades g WHERE e.enrollment_id = g.enrollment_id "
      sql = sql & " AND g.active = 'Y' AND g.approved = 'Y' AND grade_value_id IS NOT null "
    elseif report = "incompletegrades" then
      sql = sql & " AND EXISTS ( "
      sql = sql & " SELECT g.student_grade_id FROM student_grades g WHERE e.enrollment_id = g.enrollment_id "
      sql = sql & " AND g.active = 'Y' AND grade_value_id IN (14, 16, 329) "    
    else
      sql = sql & " AND EXISTS ( "
      sql = sql & " SELECT g.student_grade_id FROM student_grades g WHERE e.enrollment_id = g.enrollment_id "
      sql = sql & " AND g.active = 'Y' AND (g.approved IS NULL or g.approved <> 'Y') AND grade_value_id IS NOT null "
    end if
    sql = sql & " ) "
    if session("zone") <> "" then
      sql = sql & " AND c." & session("zone") & "_number is not null "
    end if
      sql = sql & " GROUP BY e.section_id, o.transcript_title, c.zone_name, c.course_number, s.class_start "
      sql = sql & " ORDER BY c.course_number, s.class_start "
    
    'response.Write sql
    'response.end
    rs.open sql,conn,1,1
    if rs.eof then
      if report = "missinggrades" then
        response.write("<i>Their are no overdue grades</i>")
      elseif report = "incompletegrades" then
        response.write("<i>Their are no incomplete grades</i>")
      else
        response.write("<i>All grades have been approved</i>")
      end if
    else
      if report = "missinggrades" then
        response.write("<font face='arial,helvetica' size='-1'>Sections listed below ended more than 21 days ago.</font><br>")
        response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")
      end if
      response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign='center'>")
      response.write("<table border=0 cellspacing=1 cellpadding=1>")
      response.write("<tr>")
      response.write("<td bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Section&nbsp;</b></font></td>")
      response.write("<td bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Title&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Enrolled&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Received&nbsp;</b></font></td>")
      if report = "missinggrades" then
        response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Overdue&nbsp;</b></font></td>")
      elseif report = "incompletegrades" then
        response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Incomplete&nbsp;</b></font></td>")
      else
        response.write("<td align='center' bgcolor='#EEEEEE'><font face='arial,helvetica' size='-1'><b>&nbsp;Unapproved&nbsp;</b></font></td>")
      end if
      response.write("</tr>")
      
      while not rs.eof
        numrec = NumGradesReceived(rs("section_id"))
        totnum = NumEnrolled(rs("section_id"))
        response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" onClick=""parent.location = 'c_grades.asp?id=" & rs("section_id") & "'"">")
        response.write("<td><font face='arial,helvetica' size='-1'>&nbsp;" & rs("section_name") & "&nbsp;</font></td>")
        response.write("<td><font face='arial,helvetica' size='-1'>&nbsp;" & rs("transcript_title") & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>&nbsp;" & totnum & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>&nbsp;" & numrec & "&nbsp;</font></td>")
        response.write("<td align='right'><font face='arial,helvetica' size='-1'>&nbsp;" & rs("n") & "&nbsp;</font></td>")
        response.write("</tr>")
        rs.movenext
      wend
      response.write("</table>")
      response.write("</td></tr></table>")
    end if
    rs.close
  case "queuedletters"
    sql = " select distinct p.full_name, i.student_instance_id, l.letter_name, c.scheduled_date, g.program_name, t.track_name, c.correspondence_id "
    sql = sql & " from correspondence c, letter_defs l, student_instance i, person p, programs g, program_track t "
    sql = sql & " where c.letter_id = l.letter_id and c.status = 'Queued' "
    sql = sql & " and c.student_instance_id = i.student_instance_id and i.person_id = p.person_id "
    sql = sql & " and i.program_id = g.program_id and i.program_track_id = t.program_track_id "
    if session("zone") <> "" then
      sql = sql & " and g.zone_name = " & checkstring(session("zone"),50)
    end if
    select case request("sort")
    case "1"
      sql = sql & " order by p.full_name, g.program_name, t.track_name, l.letter_name "
    case "2"
      sql = sql & " order by l.letter_name, p.full_name, i.student_instance_id, c.scheduled_date "
    case "3"
      sql = sql & " order by g.program_name, t.track_name, p.full_name, l.letter_name "
    case else
      sql = sql & " order by c.scheduled_date, p.full_name, i.student_instance_id, l.letter_name "
    end select
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("<i>no queued letters</i>")
    else
      response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign='center'>")
      response.write("<table border=0 cellspacing=1 cellpadding=1>")
      response.write("<tr>")
      response.write("<td align='center' bgcolor='#EEEEEE' onMouseOver=""this.bgColor='#CCCCCC'"" onMouseOut=""this.bgColor='#EEEEEE'"" onClick=""document.location='r_menu.asp?task=details&report=queuedletters&sort=1'""><font face='arial,helvetica' size='-1'><b>&nbsp;Person&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE' onMouseOver=""this.bgColor='#CCCCCC'"" onMouseOut=""this.bgColor='#EEEEEE'"" onClick=""document.location='r_menu.asp?task=details&report=queuedletters&sort=2'""><font face='arial,helvetica' size='-1'><b>&nbsp;Letter&nbsp;</b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE' onMouseOver=""this.bgColor='#CCCCCC'"" onMouseOut=""this.bgColor='#EEEEEE'"" onClick=""document.location='r_menu.asp?task=details&report=queuedletters&sort=3'""><font face='arial,helvetica' size='-1'><b><nobr>&nbsp;Program&nbsp;(Track)&nbsp;</nobr></b></font></td>")
      response.write("<td align='center' bgcolor='#EEEEEE' onMouseOver=""this.bgColor='#CCCCCC'"" onMouseOut=""this.bgColor='#EEEEEE'"" onClick=""document.location='r_menu.asp?task=details&report=queuedletters&sort=4'""><font face='arial,helvetica' size='-1'><b>&nbsp;Scheduled&nbsp;</b></font></td>")
      response.write("</tr>")
      while not rs.eof
        response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" onClick=""parent.location = 'p_details.asp?pg=correspond&id=" & rs("student_instance_id") & "'"">")
        response.write("<td><font face='arial,helvetica' size='-1'><nobr>&nbsp;" & rs("full_name") & "&nbsp;</nobr></font></td>")
        response.write("<td align='center'><font face='arial,helvetica' size='-1'><nobr>&nbsp;" & rs("letter_name") & "&nbsp;</nobr></font></td>")
        response.write("<td><font face='arial,helvetica' size='-1'><nobr>&nbsp;" & rs("program_name") & " (" & rs("track_name") & ")&nbsp;</nobr></font></td>")
        response.write("<td align='center'><font face='arial,helvetica' size='-1'><nobr>&nbsp;" & rs("scheduled_date") & "&nbsp;</nobr></font></td>")
        response.write("</tr>")
        rs.movenext
      wend
      response.write("</table>")
      response.write("</td></tr></table>")
    end if
    rs.close
  end select

  response.write("</td>")
  response.write("<td></form></td>")
  response.write("</tr></table>")
  response.write("<br></body></html>")
end sub




%>