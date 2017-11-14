<!--#include file ="checkusr.asp"-->
<!--#include file ="vb.asp"-->
<!--#include file ="db.asp"-->
<%

set rs = Server.CreateObject("ADODB.RecordSet")
session("pagetab") = "People"
session("history") = "Menu|p_menu.asp|Select a Student|p_select.asp|Summary"

PageHeader

id = request("id")

'2013-07-25- Rob added this to get the amcas Id to be displayed on Medical Zone students only

'2/11/2014 sjt5 simplified query
sql = "SELECT mal.amcas_id FROM Madris.dbo.student_instance si INNER JOIN MADRIS.dbo.Madris_Amcas_Lookup mal ON si.person_id = mal.person_id WHERE si.student_instance_id = " & checkstring(id,50)

rs.open sql,conn,1,1
if not rs.eof then
    strAmcasId = rs("amcas_id")
end if
rs.close

sql = "select * from student_instance where student_instance_id = " & checkstring(id,50)
rs.open sql,conn,1,1
if not rs.eof then
  person_id = rs("person_id")
  reg_status_id = rs("reg_status_id")
  program_id = rs("program_id")
  program_track_id = rs("program_track_id")
  if isnull(program_track_id) then program_track_id = ""
  expected_grad_date = rs("expected_grad_date")
  reg_status_date = ""
  year_of_study_id = rs("year_of_study_id")
  matric_date = rs("matric_date")
  original_class = rs("original_class")
  if isnull(original_class) then original_class = ""
  original_class = ""
  society = rs("society")
  if isnull(society) then society = ""
  current_class = rs("current_class")
  if isnull(current_class) then current_class = ""
end if
rs.close
if program_track_id <> "" then
  sql = "select track_name from program_track where program_track_id = " & checkstring(program_track_id,50)
  rs.open sql,conn,1,1
  if not rs.eof then track_name = rs("track_name")
  rs.close
end if
if program_id <> "" then
  sql = "select zone_name from programs where program_id = " & checkstring(program_id,50)
  rs.open sql,conn,1,1
  if not rs.eof then program_zone = rs("zone_name")
  rs.close
end if
sql = "select * from person where person_id = " & checkstring(person_id,50)
rs.open sql,conn,1,1
if not rs.eof then
  last_name = rs("last_name")
  first_name = rs("first_name")
  middle_name = rs("middle_name")
  nick_name = rs("nick_name")
  harvard_id = rs("harvard_id")
  date_of_birth = rs("date_of_birth")
  ssn = rs("ssn")
  dentpin = rs("dentpin")
  gender = rs("gender")
  citizenship = rs("citizenship")
  if isnull(gender) then gender = ""
  if isnull(citizenship) then citizenship = ""
  if rs("name_restriction") = "Y" then
    name_restriction = true
    name_restriction_style = ";background-color:FFFF99"
  else
    name_restriction = false
  end if
  if rs("address_restriction") = "Y" then
    address_restriction = true
    address_restriction_style = ";background-color:FFFF99"
  else
    address_restriction = false
  end if
  email_address = rs("email_address")
  if isnull(email_address) then email_address = ""
end if
rs.close
'05/30/2017: Check for default active 'Mailing' address.  Use original 'Local/Mailing' address logic if none found.
sql = "select * from person_address where address_type = 'Mailing' and active = 'Y' and person_id = " & checkstring(person_id,50)
rs.open sql,conn,1,1
if not rs.eof then
  line1 = rs("line1")
  line2 = rs("line2")
  line3 = rs("line3")
  city = rs("city")
  state = rs("state")
  zip = rs("zip")
  province = rs("province")
  country = rs("country")
  phone_number = rs("phone_number")

  rs.close
else
    rs.close 
    sql = "select * from person_address where address_type = 'Local/Mailing' and active = 'Y' and person_id = " & checkstring(person_id,50)
    rs.open sql,conn,1,1
    if not rs.eof then
      line1 = rs("line1")
      line2 = rs("line2")
      line3 = rs("line3")
      city = rs("city")
      state = rs("state")
      zip = rs("zip")
      province = rs("province")
      country = rs("country")
      phone_number = rs("phone_number")
    end if
    rs.close
end if

sql = "select e.ethnicity_id, e.ethnicity_code from person_ethnicity as p, ethnicity as e "
sql = sql & " where p.ethnicity_id = e.ethnicity_id and primary_flag = 'Y' and person_id = " & checkstring(person_id,50)
rs.open sql,conn,1,1
if not rs.eof then
  ethnicity_id = rs("ethnicity_id")
  ethnicity_code = rs("ethnicity_code")
  if isnull(ethnicity_code) then ethnicity_code = ""
end if
rs.close

%>
<script>
function DoChange() {
  if (!ChangesMade) {
    ChangesMade = true;
    document.footerform.bsave.disabled = false;
    document.footerform.breset.disabled = false;
    document.footerform.bsave.style.fontWeight = 'bold';
  }
}
function UndoChanges() {
  ChangesMade = false;
  document.dataform.reset();
  document.footerform.bsave.style.fontWeight = 'normal';
}
function SaveChanges() {
  alert('Saved');
}
function OpenPhoto() {
  window.open('photo.asp?id=<%=harvard_id%>','photowindow','width=250,height=250,menubar,resizable');
}
</script>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td width=1><form name="dataform" method=GET action="p_summary.asp"></td>
<td align=center valign=middle>


<table border=0 cellspacing=0 cellpadding=0><tr><td align=center bgcolor=CCCCCC>
<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>

<font face='arial,helvetica' size=-1>


<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=2><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=names&id=<%=id%>" id="alink2">Name</a></b>:&nbsp;&nbsp;Last&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>First&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Middle&nbsp;Name</font></td>
<td rowspan=2><img src="images/spacer.gif" border=0 width=5 height=1></td>
<td align=left><font face='arial,helvetica' size=-1>Nick&nbsp;Name</font></td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
<td align=left><input type=text ContentEditable=false name="last_name" value="<%=last_name%>" size=15 style="width:175px<%=name_restriction_style%>" onchange="DoChange();"></td>
<td align=left><input type=text ContentEditable=false name="first_name" value="<%=first_name%>" size=15 style="width:175px<%=name_restriction_style%>" onchange="DoChange();"></td>
<td align=left><input type=text ContentEditable=false name="middle_name" value="<%=middle_name%>" size=15 style="width:125px<%=name_restriction_style%>" onchange="DoChange();"></td>
<td align=left><input type=text ContentEditable=false name="nick_name" value="<%=nick_name%>" size=15 style="width:100px<%=name_restriction_style%>" onchange="DoChange();"></td>
</tr>
</table>

<img src="images/spacer.gif" border=0 width=1 height=5><br>

<b><a href="p_details.asp?pg=other&id=<%=id%>" id="alink2">Other&nbsp;Info</a></b><br>


<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td rowspan=3>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Harvard&nbsp;ID:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="harvard_id" value="<%=harvard_id%>" size=15 style="width:90px" onchange="DoChange();"></td>
<td rowspan=3>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Gender:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="gender" value="<%=gender%>" style="width:150px" onchange="DoChange();"></td>
<td rowspan=3>&nbsp;&nbsp;</td>
<td align=center><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=family&id=<%=id%>" id="alink2">Family&nbsp;Info</a></b></font></td></tr></table></td></tr></table></td>
<td rowspan=3>&nbsp;&nbsp;</td>
<td rowspan=3 valign=middle align=right><a href="JavaScript:OpenPhoto();"><img src="photo.asp?small=Y&id=<%=harvard_id%>" border=1 width=65 height=65></a></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>SSN:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="ssn" value="<%=ssn%>" size=15 style="width:90px" onchange="DoChange();"></td>
<td align=right><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=citizenship&id=<%=id%>" id="alink2">Citizenship</a></b>:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="citizenship" value="<%=citizenship%>" style="width:150px" onchange="DoChange();"></td>
<td align=center><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=requirements&id=<%=id%>" id="alink2">Requirements</a></b></font></td></tr></table></td></tr></table></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Birthday:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="date_of_birth" value="<%=date_of_birth%>" size=15 style="width:90px" onchange="DoChange();"></td>
<% if session("zone") = "Dental" then %>
	<td align="right"><font face='arial,helvetica' size="-1">DentPin:&nbsp;</font></td>
	<td align="left"><input type=text ContentEditable=false name="dentpin" value="<%=dentpin%>" size="15" style="width:90px" onchange="DoChange();" /></td>
<%elseif  session("zone") = "Medical" then %>
    <td align="right"><font face='arial,helvetica' size="-1">AAMC Id:&nbsp;</font></td>
    <td align="left"><input type=text ContentEditable=false name="strAmcasId" value="<%=strAmcasId%>" size="15" style="width:90px" onchange="DoChange();" /></td>
<% else %>
<td>&nbsp;</td>
<td>&nbsp;</td>
<% end if %>
<td align=center><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1>
<b>
<!--<a href="p_details.asp?pg=exams&id=<%=id%>" id="alink2">Exam&nbsp;Scores</a>-->
<a href="p_examscores.asp?siid=<%=id%>&pid=<%=person_id%>" id="alink2">Exam&nbsp;Scores</a>
</b>
</font></td></tr></table></td></tr></table></td>
</tr>
</table>


<img src="images/spacer.gif" border=0 width=1 height=5><br>


<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left colspan=2><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=addresses&id=<%=id%>" id="alink2">Address</a></b></font></td>
<td rowspan=10>&nbsp;&nbsp;</td>
<td align=left colspan=5><table border=0 cellspacing=0 cellpadding=0><tr>
<td valign=middle align=left><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=instance&id=<%=id%>" id="alink2">Student&nbsp;Instance</a></b></font></td>
<%
sql = "select count(*) n from student_instance where person_id = " & checkstring(person_id,50)
rs.open sql,conn,1,1
if not rs.eof then
  n = cint(rs("n"))
  response.write("<td valign=middle align=left><font face='arial,helvetica' size=-1 color=990000><b>&nbsp;&nbsp;(<a href=""p_details.asp?pg=instances&id=" & id & """ id=""alink2"">1 of " & n & "</a>)</b></font></td>")
end if
rs.close
%>
</tr></table></td>
</tr>
<tr>
<td rowspan=7>&nbsp;&nbsp;</td>
<td align=left><input type=text ContentEditable=false name="line1" value="<%=line1%>" size=15 style="width:300px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td rowspan=9>&nbsp;&nbsp;</td>
<td align=right><font face='arial,helvetica' size=-1>Program:&nbsp;</font></td>
<%
sql = "select program_name from programs where program_id = " & checkstring(program_id,50)
rs.open sql,conn,1,1
if not rs.eof then program_name = rs("program_name")
rs.close
%>
<td align=left colspan=3><input type=text ContentEditable=false name="program_id" value="<%=program_name%>" style="width:200px" onchange="DoChange();"></td>
</tr>
<tr>
<td align=left><input type=text ContentEditable=false name="line2" value="<%=line2%>" size=15 style="width:300px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td align=right><font face='arial,helvetica' size=-1>Track:&nbsp;</font></td>
<%
sql = "select track_name from program_track where program_track_id = " & checkstring(program_track_id,50)
rs.open sql,conn,1,1
if not rs.eof then track_name = rs("track_name")
rs.close
%>
<td align=left colspan=3><input type=text ContentEditable=false name="program_track_id" value="<%=track_name%>" style="width:200px" onchange="DoChange();"></td>
</tr>
<tr>
<td align=left><input type=text name="line3" value="<%=line3%>" size=15 style="width:300px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td align=right><font face='arial,helvetica' size=-1>Status:&nbsp;</font></td>
<%
sql = "select reg_status_name from reg_status where reg_status_id = " & checkstring(reg_status_id,50)
rs.open sql,conn,1,1
if not rs.eof then reg_status_name = rs("reg_status_name")
rs.close
%>
<td align=left colspan=3><input type=text ContentEditable=false name="reg_status_id" value="<%=reg_status_name%>" style="width:200px" onchange="DoChange();"></td>
</tr>
<tr>
<td align=left valign=middle rowspan=2>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left><font face='arial,helvetica' size=-1>City</font></td>
<td rowspan=2>&nbsp;</td>
<td align=left><font face='arial,helvetica' size=-1>State</font></td>
<td rowspan=2>&nbsp;</td>
<td align=left><font face='arial,helvetica' size=-1>Zip</font></td>
</tr>
<tr>
<td align=left><input type=text ContentEditable=false name="city" value="<%=city%>" size=15 style="width:160px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td align=left><input type=text ContentEditable=false name="state" value="<%=state%>" size=15 style="width:50px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td align=left><input type=text ContentEditable=false name="zip" value="<%=zip%>" size=15 style="width:80px<%=address_restriction_style%>" onchange="DoChange();"></td>
</tr>
</table>
</td>
<td align=right><font face='arial,helvetica' size=-1>Society:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="society" value="<%=society%>" size=15 style="width:80px" onchange="DoChange();"></td>
<td><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=left valign=bottom><font face='arial,helvetica' size=-1>Exp&nbsp;Grad&nbsp;Dt:&nbsp;</font></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Matric&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="matric_date" value="<%=matric_date%>" size=15 style="width:80px" onchange="DoChange();"></td>
<td><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="expected_grad_date" value="<%=expected_grad_date%>" size=15 style="width:80px" onchange="DoChange();"></td>
</tr>
<tr>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Province:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="province" value="<%=province%>" size=15 style="width:80px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Country:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="country" value="<%=country%>" size=15 style="width:100px<%=address_restriction_style%>" onchange="DoChange();"></td>
</tr>
</table>
</td>
<td align=right><font face='arial,helvetica' size=-1>Class:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="current_class" value="<%=current_class%>" style="width:80px" onchange="DoChange();"></td>
<td><font face='arial,helvetica' size=-1>&nbsp;</font></td>
<%
sql = "select year_of_study_name from year_of_study where year_of_study_id = " & checkstring(year_of_study_id,50)
rs.open sql,conn,1,1
if not rs.eof then year_of_study_name = rs("year_of_study_name")
rs.close
%>
<td align=left colspan=2><font face='arial,helvetica' size=-1>YoS:&nbsp;<input type=text ContentEditable=false name="year_of_study_id" value="<%=year_of_study_name%>" style="width:50px" onchange="DoChange();"></td>
</tr>
<tr>
<td align=left valign=top>
<table border=0 cellspacing=0 cellpadding=0 width=100%>
<tr>
<td align=right><font face='arial,helvetica' size=-1>Phone:&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="phone_number" value="<%=phone_number%>" size=15 style="width:125px<%=address_restriction_style%>" onchange="DoChange();"></td>
<td width=100%><font face='arial,helvetica' size=-1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
<td align=right><table border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1>&nbsp;<b><a href="p_details.asp?pg=ferpa&id=<%=id%>" id="alink2">FERPA&nbsp;Flags</a></b>&nbsp;</font></td></tr></table></td><td>&nbsp;&nbsp;</td></tr></table></td>
</tr>
</table>
</td>
<td colspan=4>
<table border=0 cellspacing=3 cellpadding=0 width=100%>
<tr>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=advisors&id=<%=id%>" id="alink2">Advisors</a></b></font></td></tr></table></td></tr></table></td>
<td align=center width=40%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_enrollment.asp?id=<%=id%>" id="alink2">Enrollment</a></b></font></td></tr></table></td></tr></table></td>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=reghist&id=<%=id%>" id="alink2">Reg&nbsp;Hist</a></b></font></td></tr></table></td></tr></table></td>
</tr>
</table>
</td>
</tr>
<tr>
<td align=left valign=bottom colspan=2 rowspan=2>
<table border=0 cellspacing=0 cellpadding=0 width=100%>
<tr>
<td align=left colspan=2 valign=bottom><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=other&id=<%=id%>" id="alink2">Email</a></b></font></td>
<td align=center><table border=0 cellspacing=0 cellpadding=0><tr><td><td align=center bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1>&nbsp;<b><a href="p_details.asp?pg=emergency&id=<%=id%>" id="alink2">Emergency&nbsp;Contact</a></b>&nbsp;</font></td></tr></table></td><td>&nbsp;&nbsp;</td></tr></table><img src="images/spacer.gif" border=0 width=1 height=5></td>
</tr>
</table>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left><font face='arial,helvetica' size=-1>&nbsp;&nbsp;</font></td>
<td align=left><input type=text ContentEditable=false name="email_address" value="<%=email_address%>" size=15 style="width:300px<%=email_restriction_style%>" onchange="DoChange();"></td>
</tr>
</table>
</td>
<td colspan=4>
<table border=0 cellspacing=3 cellpadding=0 width=100%>
<tr>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=projects&id=<%=id%>" id="alink2">Projects</a></b></font></td></tr></table></td></tr></table></td>
<td align=center width=40%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=correspond&id=<%=id%>" id="alink2">Correspond</a></b></font></td></tr></table></td></tr></table></td>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=degrees&id=<%=id%>" id="alink2">Degrees</a></b></font></td></tr></table></td></tr></table></td>
</tr>
</table>
</td>
</tr>
<tr>
<td colspan=4>
<table border=0 cellspacing=3 cellpadding=0 width=100%>
<tr>
<%if program_zone="DMS" then%>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="dms_allocations.asp?siid=<%=id%>" id="alink2">Billing</a></b></font></td></tr></table></td></tr></table></td>
<%elseif program_track_id=16 then%>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="x_sched_details.asp?task=edit&rpg=p&siid=<%=id%>" id="alink2">Billing</a></b></font></td></tr></table></td></tr></table></td>
<%else%>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=termbill&id=<%=id%>" id="alink2">Billing</a></b></font></td></tr></table></td></tr></table></td>
<%end if%>
<td align=center width=40%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=tracks&id=<%=id%>" id="alink2">Prgm&nbsp;Tracks</a></b></font></td></tr></table></td></tr></table></td>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_degaudit.asp?id=<%=id%>" id="alink2">Deg&nbsp;Audit</a></b></font></td></tr></table></td></tr></table></td>
</tr>

</table>
<table border=0 cellspacing=3 cellpadding=0 width=100%>
<tr>
<td align=center width=40%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=98% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=ethnicityNew&id=<%=id%>" id="alink2">Ethnicity&nbsp;[IPEDS]</a></b></font></td></tr></table></td></tr></table></td>
<td align=center width=30%><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td><td align=center width=100% bgcolor=CCCCCC><table width=100% border=0 cellspacing=1 cellpadding=1><tr><td align=center bgcolor=F9F9F9><font face='arial,helvetica' size=-1><b><a href="p_details.asp?pg=ethnicity&id=<%=id%>" id="alink2">Ethnicity&nbsp;(old)</a></b></font></td></tr></table></td></tr></table></td>
<td> </td>
</tr>
</table>
</td>
</tr>
</table>



</font>


</td></tr></table>
</td></tr></table>


</td>
<td width=1></form></td>
</tr>
</table>
<%



PageMiddle
%>
<td align=center valign=middle>
<% if request("back") <> "" then %>
<input style="width:75px; background-color:#CCCCCC" type=button onClick="document.location='<%=request("back")%>?id=<%=request("id2")%>';" name="bsave" value=" Back ">
<% else %>
<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=000000>
<table border=0 cellspacing=1 cellpadding=2><tr>
<td valign=middle bgcolor=FFFF99><font face='arial,helvetica' size=-1>&nbsp;<b>Note:</b>&nbsp;Yellow&nbsp;indicates&nbsp;FERPA&nbsp;protected&nbsp;(restricted)&nbsp;data.&nbsp;</font></td>
</tr></table>
</td></tr></table>
<% end if %>
</td>
<%
PageFooter

%>